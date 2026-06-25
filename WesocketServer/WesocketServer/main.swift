import Foundation

// Suppress Network.framework's low-level socket diagnostics that appear in the
// console whenever a WebSocket client disconnects (normal, expected behaviour):
//   nw_read_request_report [C…] Receive failed with error "Socket is not connected"
// These messages go through the OS unified-log activity system (os_activity).
// Setting OS_ACTIVITY_MODE=disable mutes them without touching our print() output.
// Must be called before the Network stack makes its first log call.
setenv("OS_ACTIVITY_MODE", "disable", 1)

// MARK: - Configuration from environment / args

let wsPort  = UInt16(ProcessInfo.processInfo.environment["CAPTURE_WS_PORT"]   ?? "") ?? 8765
let httpPort = UInt16(ProcessInfo.processInfo.environment["CAPTURE_HTTP_PORT"] ?? "") ?? 8080

let dataDir: URL = {
    if let custom = ProcessInfo.processInfo.environment["CAPTURE_DATA_DIR"] {
        return URL(fileURLWithPath: custom)
    }
    return FileManager.default.urls(for: .applicationSupportDirectory, in: .userDomainMask)[0]
        .appending(path: "OutlookCapture")
}()

// MARK: - Bootstrap

print("[Server] Data directory : \(dataDir.path(percentEncoded: false))")
print("[Server] WebSocket port  : \(wsPort)  (ingest + dashboard live feed)")
print("[Server] HTTP port       : \(httpPort) (dashboard UI + /api/events + /screenshots)")

let store: EventStore
do {
    store = try EventStore(dataDir: dataDir)
} catch {
    fputs("[Server] Failed to initialise EventStore: \(error)\n", stderr)
    exit(1)
}

// Load TLS identity from ~/.office-addin-dev-certs/ (generated automatically via openssl).
let tlsIdentity = WebSocketServer.loadIdentity()
if tlsIdentity != nil {
    print("[Server] TLS enabled  → wss://localhost:\(wsPort)")
} else {
    print("[Server] TLS disabled → ws://localhost:\(wsPort)  ⚠️  add-in will be blocked by mixed-content")
}

let wsServer  = WebSocketServer(port: wsPort)
let httpServer = HTTPServer(port: httpPort)
httpServer.eventStore = store

wsServer.onEvent = { event in
    Task {
        do {
            try await store.append(event)

            // ── Diagnostic log line ──────────────────────────────────────────
            let source: String
            switch event.appBundleId ?? "" {
            case let b where b.contains("Outlook"): source = "\u{001B}[34m[Outlook]\u{001B}[0m"
            case let b where b.contains("Excel"):   source = "\u{001B}[32m[Excel]\u{001B}[0m"
            default:                                 source = "\u{001B}[35m[Add-in]\u{001B}[0m"
            }

            let detail: String
            if let subject = event.metadata?.subject, !subject.isEmpty {
                let toStr = event.metadata?.to.prefix(2).joined(separator: ", ") ?? ""
                let attStr = event.metadata?.attachments.isEmpty == false
                    ? " 📎\(event.metadata!.attachments.count)"
                    : ""
                detail = "\"\(subject)\"\(toStr.isEmpty ? "" : " → \(toStr)")\(attStr)"
            } else if let range = event.excelMetadata?.selectedRange {
                detail = "\(event.excelMetadata?.workbookName ?? "")/\(event.excelMetadata?.sheetName ?? "") \(range)"
            } else {
                detail = "(no detail)"
            }

            let emoji: String
            switch event.action {
            case .send:             emoji = "📤"
            case .reply:            emoji = "↩ "
            case .replyAll:         emoji = "↩↩"
            case .forward:          emoji = "↪ "
            case .save:             emoji = "💾"
            case .compose:          emoji = "✏️ "
            case .select:           emoji = "👁 "
            case .close:            emoji = "🚪"
            case .recipientsChange: emoji = "👥"
            case .attachmentChange: emoji = "📎"
            case .fromChange:       emoji = "🔀"
            case .timeChange:       emoji = "🕐"
            case .recurrenceChange: emoji = "🔁"
            case .locationChange:   emoji = "📍"
            case .meetingAccept:    emoji = "✅"
            case .meetingTentative: emoji = "🤔"
            case .meetingDecline:   emoji = "❌"
            case .shortcutKey:      emoji = "⌨ "
            case .selectionChange:  emoji = "🖱 "
            case .edit:             emoji = "✏️ "
            case .open:             emoji = "📂"
            default:                emoji = "• "
            }

            let ts = String(event.timestamp.prefix(19)).replacingOccurrences(of: "T", with: " ")
            print("\u{001B}[1m[Capture]\u{001B}[0m \(source) \(emoji) \(event.action.rawValue) [\(event.itemKind.rawValue)]  \(detail)  \u{001B}[2m\(ts)\u{001B}[0m")
        } catch {
            print("[Store] Write error: \(error)")
        }
    }
}

do {
    try wsServer.start(identity: tlsIdentity)
    try httpServer.start()
} catch {
    fputs("[Server] Failed to start: \(error)\n", stderr)
    exit(1)
}

print("[Server] Dashboard → http://127.0.0.1:\(httpPort)")

// Keep the process alive.
RunLoop.main.run()
