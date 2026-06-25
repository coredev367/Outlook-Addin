import Foundation

// ── Silence Network.framework noise ──────────────────────────────────────────
//
// Apple's Network.framework prints low-level socket diagnostics such as:
//   nw_read_request_report [C8] Receive failed with error "Socket is not connected"
// whenever a WebSocket client disconnects — perfectly normal, entirely unactionable.
//
// Two channels deliver these messages; we kill both:
//   1. Unified-log / os_activity  → OS_ACTIVITY_MODE=disable
//   2. Direct stderr writes       → pipe-based line filter below

// Channel 1 – must be set before the network stack initialises.
setenv("OS_ACTIVITY_MODE", "disable", 1)

// Channel 2 – intercept stderr, drop lines that match known NW noise patterns,
// forward everything else to the real stderr (Xcode console or terminal tty).
func installStderrFilter() {
    // Patterns that identify Network.framework internals.
    // None of our own [Server]/[WSServer]/[Capture] messages match these.
    let dropPatterns: [StaticString] = [
        "nw_read_request_report",
        "Socket is not connected",
        "nw_socket_handle_socket_event",
        "nw_connection_copy_connected_path",
        "boringssl_",
        "TIC Read Status",
    ]

    let pipe     = Pipe()
    let realFd   = dup(STDERR_FILENO)              // save original stderr
    dup2(pipe.fileHandleForWriting.fileDescriptor, STDERR_FILENO)
    close(pipe.fileHandleForWriting.fileDescriptor) // STDERR_FILENO is now the sole writer

    let realOut  = FileHandle(fileDescriptor: realFd, closeOnDealloc: true)

    Thread.detachNewThread {
        var buf = Data()
        while true {
            let chunk = pipe.fileHandleForReading.availableData
            guard !chunk.isEmpty else { break }
            buf.append(chunk)

            // Emit complete lines only — never split a message across chunks.
            while let nl = buf.firstIndex(of: UInt8(ascii: "\n")) {
                let line = Data(buf[buf.startIndex ... nl])
                buf.removeSubrange(buf.startIndex ... nl)

                guard let text = String(data: line, encoding: .utf8) else {
                    realOut.write(line); continue
                }
                let suppress = dropPatterns.contains { text.contains("\($0)") }
                if !suppress { realOut.write(line) }
            }
        }
        if !buf.isEmpty { realOut.write(buf) }   // flush any trailing partial line
    }
}

installStderrFilter()

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
    print("[Server] Failed to initialise EventStore: \(error)")
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
    print("[Server] Failed to start: \(error)")
    exit(1)
}

print("[Server] Dashboard → http://127.0.0.1:\(httpPort)")

// Keep the process alive.
RunLoop.main.run()
