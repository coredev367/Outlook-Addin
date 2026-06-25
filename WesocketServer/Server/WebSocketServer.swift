import Foundation
import Network

/// WebSocket server on a single port.
///
/// Two roles share the same port, differentiated by their first message:
///   - **Agent** sends `{"type":"ingest","event":{...}}` — events are stored and fanned out.
///   - **Dashboard browser** sends `{"type":"subscribe"}` — receives every subsequent broadcast.
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

    func start() throws {
        let params = NWParameters.tcp
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
