import Foundation
import Network

/// Minimal HTTP/1.1 server that serves the live dashboard, the event API, and screenshots.
///
/// Routes:
///   GET  /             → dashboard HTML (embedded)
///   GET  /api/events   → JSON array of recent CaptureEvents
///   GET  /screenshots/ → PNG files from EventStore.screenshotsDirectory
///   POST /ingest       → accepts a CaptureEvent JSON body (used by LaunchEvent handlers
///                        whose JS-only runtime does not support WebSocket)
final class HTTPServer: @unchecked Sendable {

    private let port: NWEndpoint.Port
    private var listener: NWListener?
    private let queue = DispatchQueue(label: "oc.http.server", qos: .userInteractive)

    var eventStore: EventStore?

    /// Called when a CaptureEvent arrives via POST /ingest (same contract as WebSocketServer.onEvent).
    var onEvent: (@Sendable (CaptureEvent) -> Void)?

    init(port: UInt16 = 8080) {
        self.port = NWEndpoint.Port(rawValue: port)!
    }

    func start() throws {
        listener = try NWListener(using: .tcp, on: port)
        listener?.stateUpdateHandler = { [weak self] state in
            if case .ready = state {
                print("[HTTPServer] Listening on port \(self?.port.rawValue ?? 0)")
            }
        }
        listener?.newConnectionHandler = { [weak self] conn in
            self?.accept(conn)
        }
        listener?.start(queue: queue)
    }

    func stop() { listener?.cancel() }

    // MARK: - Connection lifecycle

    private func accept(_ conn: NWConnection) {
        conn.stateUpdateHandler = { [weak self] state in
            if case .ready = state { self?.readRequest(conn) }
        }
        conn.start(queue: queue)
    }

    private func readRequest(_ conn: NWConnection) {
        conn.receive(minimumIncompleteLength: 1, maximumLength: 65_536) { [weak self] data, _, _, error in
            guard let self else { return }
            guard let data, !data.isEmpty, error == nil else { conn.cancel(); return }
            guard let text = String(data: data, encoding: .utf8) else { conn.cancel(); return }
            self.route(request: text, conn: conn)
        }
    }

    // MARK: - Router

    private func route(request: String, conn: NWConnection) {
        let lines    = request.components(separatedBy: "\r\n")
        let firstLine = lines.first ?? ""
        let parts    = firstLine.components(separatedBy: " ")
        guard parts.count >= 2 else { respond(400, "text/plain", "Bad Request", conn); return }

        let method = parts[0]
        let path   = parts[1].components(separatedBy: "?").first ?? parts[1]

        // ── POST /ingest ─────────────────────────────────────────────────────
        // Receives a CaptureEvent from Outlook LaunchEvent handlers that run in
        // a JavaScript-only runtime where WebSocket is not available.
        if method == "POST" && path == "/ingest" {
            // Extract body after the blank line that separates headers from body.
            let bodyString: String
            if let blankIdx = lines.firstIndex(of: "") {
                bodyString = lines[lines.index(after: blankIdx)...].joined(separator: "\r\n")
            } else {
                bodyString = ""
            }

            guard let bodyData = bodyString.data(using: .utf8),
                  let msg = try? JSONDecoder().decode(WSMessage.self, from: bodyData),
                  msg.type == .ingest,
                  let event = msg.event
            else {
                respond(400, "application/json", #"{"error":"bad body"}"#, conn)
                return
            }

            onEvent?(event)
            // Fan-out is handled by onEvent; acknowledge to the add-in immediately.
            respond(200, "application/json", #"{"type":"ack"}"#, conn)
            return
        }

        // ── GET routes ───────────────────────────────────────────────────────
        guard method == "GET" else {
            respond(405, "text/plain", "Method Not Allowed", conn)
            return
        }

        switch path {
        case "/", "/index.html":
            respond(200, "text/html; charset=utf-8", dashboardHTML, conn)

        case "/api/events":
            Task { [weak self] in
                guard let self, let store = self.eventStore else {
                    self?.respond(503, "application/json", "[]", conn); return
                }
                let events = await store.recent(limit: 500)
                let encoder = JSONEncoder()
                encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
                let body = (try? encoder.encode(events)).flatMap { String(data: $0, encoding: .utf8) } ?? "[]"
                self.respond(200, "application/json", body, conn)
            }

        case _ where path.hasPrefix("/screenshots/"):
            let filename = String(path.dropFirst("/screenshots/".count))
            serveScreenshot(filename: filename, conn: conn)

        default:
            respond(404, "text/plain", "Not Found", conn)
        }
    }

    // MARK: - Helpers

    private func serveScreenshot(filename: String, conn: NWConnection) {
        Task { [weak self] in
            guard let self, let store = self.eventStore else {
                self?.respond(404, "text/plain", "Not Found", conn); return
            }
            let dir = await store.screenshotsDirectory
            let fileURL = dir.appending(path: filename)
            guard let data = try? Data(contentsOf: fileURL) else {
                self.respond(404, "text/plain", "Screenshot not found", conn)
                return
            }
            self.respondRaw(200, "image/png", data, conn)
        }
    }

    private func respond(_ status: Int, _ ct: String, _ body: String, _ conn: NWConnection) {
        respondRaw(status, ct, Data(body.utf8), conn)
    }

    private func respondRaw(_ status: Int, _ ct: String, _ body: Data, _ conn: NWConnection) {
        let phrase = status == 200 ? "OK" : status == 404 ? "Not Found" : "Error"
        let header = "HTTP/1.1 \(status) \(phrase)\r\n" +
                     "Content-Type: \(ct)\r\n" +
                     "Content-Length: \(body.count)\r\n" +
                     "Access-Control-Allow-Origin: *\r\n" +
                     "Connection: close\r\n\r\n"
        var response = Data(header.utf8)
        response.append(body)
        conn.send(content: response, completion: .contentProcessed { _ in conn.cancel() })
    }
}
