import Foundation
import Network

/// WebSocket client that connects to the ingest server and sends CaptureEvents.
/// Automatically reconnects with exponential back-off when the connection drops.
final class WebSocketClient: @unchecked Sendable {

    private let host: String
    private let port: UInt16
    private var connection: NWConnection?
    private let queue = DispatchQueue(label: "oc.ws.client", qos: .userInitiated)
    private var reconnectDelay: TimeInterval = 1.0
    private let maxDelay: TimeInterval = 30.0
    private var isStopped = false
    private var pendingEvents: [CaptureEvent] = []

    // MARK: - Lifecycle

    init(host: String = "127.0.0.1", port: UInt16 = 8765) {
        self.host = host
        self.port = port
    }

    func start() {
        isStopped = false
        connect()
    }

    func stop() {
        isStopped = true
        connection?.cancel()
    }

    // MARK: - Send

    func send(_ event: CaptureEvent) {
        queue.async { [weak self] in
            guard let self else { return }
            let msg = WSMessage(type: .ingest, event: event)
            guard let data = try? JSONEncoder().encode(msg) else { return }
            if let conn = self.connection, conn.state == .ready {
                self.sendData(data, to: conn)
            } else {
                // Buffer while connecting
                self.pendingEvents.append(event)
                if self.pendingEvents.count > 100 { self.pendingEvents.removeFirst() }
            }
        }
    }

    // MARK: - Connection management

    private func connect() {
        guard !isStopped else { return }

        let params = NWParameters.tcp
        let wsOptions = NWProtocolWebSocket.Options()
        wsOptions.autoReplyPing = true
        params.defaultProtocolStack.applicationProtocols.insert(wsOptions, at: 0)

        let endpoint = NWEndpoint.hostPort(
            host: NWEndpoint.Host(host),
            port: NWEndpoint.Port(rawValue: port)!
        )
        let conn = NWConnection(to: endpoint, using: params)
        self.connection = conn

        conn.stateUpdateHandler = { [weak self] state in
            guard let self else { return }
            switch state {
            case .ready:
                print("[WSClient] Connected to \(self.host):\(self.port)")
                self.reconnectDelay = 1.0
                self.flushPending(to: conn)
                self.receive(from: conn)
            case .failed(let error):
                print("[WSClient] Connection failed: \(error)")
                self.scheduleReconnect()
            case .cancelled:
                if !self.isStopped { self.scheduleReconnect() }
            default:
                break
            }
        }
        conn.start(queue: queue)
    }

    private func scheduleReconnect() {
        guard !isStopped else { return }
        let delay = reconnectDelay
        reconnectDelay = min(reconnectDelay * 2, maxDelay)
        print("[WSClient] Reconnecting in \(Int(delay))s…")
        queue.asyncAfter(deadline: .now() + delay) { [weak self] in
            self?.connect()
        }
    }

    private func flushPending(to conn: NWConnection) {
        let events = pendingEvents
        pendingEvents.removeAll()
        for event in events {
            let msg = WSMessage(type: .ingest, event: event)
            if let data = try? JSONEncoder().encode(msg) {
                sendData(data, to: conn)
            }
        }
    }

    private func sendData(_ data: Data, to conn: NWConnection) {
        let meta = NWProtocolWebSocket.Metadata(opcode: .text)
        let ctx  = NWConnection.ContentContext(identifier: "oc.ws.client", metadata: [meta])
        conn.send(content: data, contentContext: ctx, isComplete: true, completion: .idempotent)
    }

    private func receive(from conn: NWConnection) {
        conn.receiveMessage { [weak self] data, _, _, error in
            guard let self else { return }
            if let data, !data.isEmpty {
                // Server only sends `{"type":"ack"}` — no action needed.
                _ = try? JSONDecoder().decode(WSMessage.self, from: data)
            }
            if error == nil { self.receive(from: conn) }
        }
    }
}
