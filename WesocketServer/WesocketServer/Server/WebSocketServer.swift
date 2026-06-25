import Foundation
import Network
import Security

/// WebSocket server on a single port.
///
/// Two roles share the same port, differentiated by their first message:
///   - **Agent** sends `{"type":"ingest","event":{...}}` — events are stored and fanned out.
///   - **Dashboard browser** sends `{"type":"subscribe"}` — receives every subsequent broadcast.
///
/// TLS (wss://) is enabled automatically when a P12 identity is found at
/// `~/.office-addin-dev-certs/localhost.p12`.  Generate it once with:
///   cd SkanOulook && npm run prepare-wss
final class WebSocketServer: @unchecked Sendable {

    private let port: NWEndpoint.Port
    private var listener: NWListener?
    private var connections: [String: NWConnection] = [:]
    private var subscriberIds: Set<String> = []
    private let queue = DispatchQueue(label: "oc.ws.server", qos: .userInteractive)

    /// Called on the internal queue whenever a validated `CaptureEvent` arrives.
    var onEvent: (@Sendable (CaptureEvent) -> Void)?

    init(port: UInt16 = 8765) {
        self.port = NWEndpoint.Port(rawValue: port)!
    }

    // MARK: - TLS identity loader

    /// Loads the dev-cert P12 from `~/.office-addin-dev-certs/localhost.p12` (empty passphrase).
    /// Returns `nil` and logs a hint if the file is missing — server falls back to plain ws://.
    static func loadIdentity() -> SecIdentity? {
        let p12URL = URL(fileURLWithPath: NSHomeDirectory())
            .appendingPathComponent(".office-addin-dev-certs")
            .appendingPathComponent("localhost.p12")

        guard let data = try? Data(contentsOf: p12URL) else {
            print("[WSServer] ⚠️  No P12 found at \(p12URL.path)")
            print("[WSServer]    Run: cd SkanOulook && npm run prepare-wss")
            print("[WSServer]    Falling back to plain  ws://localhost — add-in will block (mixed content)")
            return nil
        }

        let opts = [kSecImportExportPassphrase as String: ""] as CFDictionary
        var items: CFArray?
        guard SecPKCS12Import(data as CFData, opts, &items) == errSecSuccess,
              let arr = items as? [[String: Any]],
              let identity = arr.first?[kSecImportItemIdentity as String] as? SecIdentity
        else {
            print("[WSServer] ⚠️  P12 import failed — falling back to plain ws://")
            return nil
        }
        return identity
    }

    // MARK: - Start

    func start(identity: SecIdentity? = nil) throws {
        let params: NWParameters

        if let identity {
            // WSS — TLS with the provided certificate
            let tlsOpts = NWProtocolTLS.Options()
            let secId   = sec_identity_create(identity)
            sec_protocol_options_set_local_identity(tlsOpts.securityProtocolOptions, secId)
            sec_protocol_options_set_min_tls_protocol_version(tlsOpts.securityProtocolOptions, .TLSv12)
            params = NWParameters(tls: tlsOpts, tcp: NWProtocolTCP.Options())
        } else {
            // WS — plain TCP (add-in will be blocked by mixed-content rules)
            params = NWParameters.tcp
        }

        let wsOptions = NWProtocolWebSocket.Options()
        wsOptions.autoReplyPing = true
        params.defaultProtocolStack.applicationProtocols.insert(wsOptions, at: 0)

        listener = try NWListener(using: params, on: port)
        listener?.stateUpdateHandler = { [weak self] state in
            guard let self else { return }
            switch state {
            case .ready:
                print("[WSServer] Listening on port \(self.port)")
            case .failed(let error):
                print("[WSServer] Listener failed: \(error)")
            default:
                break
            }
        }
        listener?.newConnectionHandler = { [weak self] conn in
            self?.accept(conn)
        }
        listener?.start(queue: queue)
    }

    func stop() {
        listener?.cancel()
        queue.async { [weak self] in
            self?.connections.values.forEach { $0.cancel() }
            self?.connections.removeAll()
            self?.subscriberIds.removeAll()
        }
    }

    // MARK: - Private

    private func accept(_ connection: NWConnection) {
        let id = UUID().uuidString
        queue.async { [weak self] in
            self?.connections[id] = connection
        }
        connection.stateUpdateHandler = { [weak self] state in
            switch state {
            case .ready:
                self?.receive(from: connection, id: id)
            case .failed, .cancelled:
                self?.queue.async {
                    self?.connections.removeValue(forKey: id)
                    self?.subscriberIds.remove(id)
                }
            default:
                break
            }
        }
        connection.start(queue: queue)
    }

    private func receive(from connection: NWConnection, id: String) {
        connection.receiveMessage { [weak self] data, _, _, error in
            guard let self else { return }
            if let data, !data.isEmpty {
                self.handle(data: data, from: id, connection: connection)
            }
            if error == nil {
                self.receive(from: connection, id: id)
            }
        }
    }

    private func handle(data: Data, from id: String, connection: NWConnection) {
        guard let msg = try? JSONDecoder().decode(WSMessage.self, from: data) else {
            print("[WSServer] Unrecognised message from \(id)")
            return
        }
        switch msg.type {
        case .ingest:
            guard let event = msg.event else { return }
            onEvent?(event)
            send(WSMessage(type: .ack), to: connection)
            let broadcast = WSMessage(type: .broadcast, event: event)
            for sid in subscriberIds {
                if let sub = connections[sid] { send(broadcast, to: sub) }
            }
        case .subscribe:
            subscriberIds.insert(id)
            print("[WSServer] Dashboard subscriber registered: \(id)")
        default:
            break
        }
    }

    private func send(_ message: WSMessage, to connection: NWConnection) {
        guard let data = try? JSONEncoder().encode(message) else { return }
        let meta = NWProtocolWebSocket.Metadata(opcode: .text)
        let ctx = NWConnection.ContentContext(identifier: "oc.ws", metadata: [meta])
        connection.send(content: data, contentContext: ctx, isComplete: true, completion: .idempotent)
    }
}
