import Foundation
import Network
import Security

/// WebSocket server on a single port.
///
/// Two roles share the same port, differentiated by their first message:
///   - **Agent** sends `{"type":"ingest","event":{...}}` — events are stored and fanned out.
///   - **Dashboard browser** sends `{"type":"subscribe"}` — receives every subsequent broadcast.
///
/// TLS (wss://) is enabled automatically by reading the office-addin-dev-certs PEM files at
/// `~/.office-addin-dev-certs/localhost.{crt,key}` and converting them to a SecIdentity at
/// startup using /usr/bin/openssl (always present on macOS). No manual setup required.
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

    /// Derives a `SecIdentity` from the office-addin-dev-certs PEM files at startup.
    ///
    /// Strategy:
    ///   1. Locate `localhost.crt` and `localhost.key` under `~/.office-addin-dev-certs/`.
    ///   2. Shell out to `/usr/bin/openssl pkcs12 -export` (always available on macOS) to
    ///      produce a temporary P12 with a fixed internal passphrase.
    ///   3. Import with `SecPKCS12Import` and return the `SecIdentity`.
    ///   4. Delete the temp file.
    ///
    /// Using a non-empty passphrase avoids the macOS keychain-interaction dialog that
    /// `SecPKCS12Import` shows when passphrase is empty and no keychain is specified.
    static func loadIdentity() -> SecIdentity? {
        let certDir  = URL(fileURLWithPath: NSHomeDirectory())
            .appendingPathComponent(".office-addin-dev-certs")
        let certPath = certDir.appendingPathComponent("localhost.crt").path
        let keyPath  = certDir.appendingPathComponent("localhost.key").path

        guard FileManager.default.fileExists(atPath: certPath),
              FileManager.default.fileExists(atPath: keyPath) else {
            print("[WSServer] ⚠️  Dev certs not found at \(certDir.path)")
            print("[WSServer]    Install them with: cd SkanOulook && npx office-addin-dev-certs install")
            return nil
        }

        // Write P12 to a unique temp path so concurrent restarts don't collide.
        let tmpP12 = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("ws_tls_\(UUID().uuidString).p12")
        defer { try? FileManager.default.removeItem(at: tmpP12) }

        // Internal passphrase — only lives in memory for the duration of this call.
        let passphrase = "ws-tls-internal"

        let proc = Process()
        proc.executableURL = URL(fileURLWithPath: "/usr/bin/openssl")
        proc.arguments = [
            "pkcs12", "-export",
            "-out",     tmpP12.path,
            "-inkey",   keyPath,
            "-in",      certPath,
            "-passout", "pass:\(passphrase)",
        ]
        // Suppress openssl's own output so server logs stay clean.
        proc.standardOutput = FileHandle.nullDevice
        proc.standardError  = FileHandle.nullDevice

        do    { try proc.run(); proc.waitUntilExit() }
        catch { print("[WSServer] ⚠️  openssl exec failed: \(error)"); return nil }

        guard proc.terminationStatus == 0 else {
            print("[WSServer] ⚠️  openssl pkcs12 exited with status \(proc.terminationStatus)")
            return nil
        }

        guard let p12Data = try? Data(contentsOf: tmpP12) else {
            print("[WSServer] ⚠️  Could not read temp P12")
            return nil
        }

        return importP12(data: p12Data, passphrase: passphrase)
    }

    // MARK: - P12 import helper

    private static func importP12(data: Data, passphrase: String) -> SecIdentity? {
        let opts = [kSecImportExportPassphrase as String: passphrase] as CFDictionary
        var items: CFArray?
        let status = SecPKCS12Import(data as CFData, opts, &items)
        guard status == errSecSuccess else {
            print("[WSServer] ⚠️  SecPKCS12Import failed (OSStatus \(status))")
            return nil
        }
        guard let arr = items as? [[String: Any]],
              let rawIdentity = arr.first?[kSecImportItemIdentity as String] else {
            print("[WSServer] ⚠️  No identity found in P12")
            return nil
        }
        // kSecImportItemIdentity is documented as SecIdentity; CFTypeRef bridge requires force-cast.
        return (rawIdentity as! SecIdentity)
    }

    // MARK: - Start

    func start(identity: SecIdentity? = nil) throws {
        let params: NWParameters

        if let identity {
            // WSS — TLS with the provided certificate
            let tlsOpts = NWProtocolTLS.Options()
            // sec_identity_create returns sec_identity_t? — it is non-nil when given a valid SecIdentity.
            let secId   = sec_identity_create(identity)!
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
