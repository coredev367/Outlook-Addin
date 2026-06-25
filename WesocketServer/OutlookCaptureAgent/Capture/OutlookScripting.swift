import AppKit
import ApplicationServices

/// Reads the currently selected / active Outlook item via AppleScript (classic Outlook for Mac)
/// and falls back to Accessibility text-scraping when AppleScript is unavailable
/// (new "One Outlook" / Microsoft 365 Outlook for Mac).
final class OutlookScripting: @unchecked Sendable {

    static let shared = OutlookScripting()
    private init() {}

    /// Whether to include the message body in captured metadata (off by default — bodies can be large).
    var captureBody = false

    private let scriptQueue = DispatchQueue(label: "oc.applescript", qos: .userInitiated)

    // MARK: - Public API

    /// Returns the current Outlook item's metadata synchronously on the calling thread.
    /// Call from a background / capture context, never from the main thread.
    func currentMetadata() -> MailMetadata? {
        if let meta = metadataViaAppleScript() { return meta }
        return metadataViaAccessibility()
    }

    // MARK: - AppleScript path (classic Outlook)

    private func metadataViaAppleScript() -> MailMetadata? {
        // Use osascript to avoid NSAppleScript's main-thread requirement
        let script = buildScript()
        let result = runOsascript(script)
        return parseResult(result)
    }

    private func buildScript() -> String {
        // Returns a tab-delimited string: subject\tto\tcc\tattachments[\tbody]
        let bodyLine = captureBody
            ? "set bodyText to content of theMsg"
            : "set bodyText to \"\""
        return """
        tell application "Microsoft Outlook"
            try
                set theMsg to item 1 of (get selection)
                set sub to subject of theMsg
                set toList to ""
                repeat with r in to recipients of theMsg
                    set toList to toList & (email address of address of r) & ","
                end repeat
                set ccList to ""
                repeat with r in cc recipients of theMsg
                    set ccList to ccList & (email address of address of r) & ","
                end repeat
                set attList to ""
                repeat with a in attachments of theMsg
                    set attList to attList & (name of a) & ","
                end repeat
                \(bodyLine)
                return sub & "\t" & toList & "\t" & ccList & "\t" & attList & "\t" & bodyText
            on error
                return ""
            end try
        end tell
        """
    }

    private func runOsascript(_ script: String) -> String? {
        let task = Process()
        task.executableURL = URL(fileURLWithPath: "/usr/bin/osascript")
        task.arguments = ["-e", script]

        let pipe = Pipe()
        task.standardOutput = pipe
        task.standardError = Pipe()  // suppress stderr

        do {
            try task.run()
            task.waitUntilExit()
            let data = pipe.fileHandleForReading.readDataToEndOfFile()
            let output = String(data: data, encoding: .utf8)?.trimmingCharacters(in: .whitespacesAndNewlines)
            return output?.isEmpty == false ? output : nil
        } catch {
            return nil
        }
    }

    private func parseResult(_ raw: String?) -> MailMetadata? {
        guard let raw, !raw.isEmpty else { return nil }
        let parts = raw.components(separatedBy: "\t")
        guard parts.count >= 4 else { return nil }

        func split(_ s: String) -> [String] {
            s.components(separatedBy: ",").map { $0.trimmingCharacters(in: .whitespaces) }.filter { !$0.isEmpty }
        }

        return MailMetadata(
            subject:     parts[0].isEmpty ? nil : parts[0],
            to:          split(parts[1]),
            cc:          split(parts[2]),
            attachments: split(parts[3]),
            body:        parts.count > 4 && !parts[4].isEmpty ? parts[4] : nil
        )
    }

    // MARK: - Accessibility fallback (new Outlook)

    private func metadataViaAccessibility() -> MailMetadata? {
        guard let app = NSWorkspace.shared.runningApplications.first(where: {
            ($0.bundleIdentifier ?? "").lowercased().contains("microsoft.outlook")
        }) else { return nil }

        let pid = app.processIdentifier
        let appEl = AXUIElementCreateApplication(pid)

        // Try to find the focused window's title (usually contains the subject)
        var windowRef: CFTypeRef?
        AXUIElementCopyAttributeValue(appEl, kAXFocusedWindowAttribute as CFString, &windowRef)
        guard let window = windowRef else { return nil }

        var titleRef: CFTypeRef?
        AXUIElementCopyAttributeValue(window as! AXUIElement, kAXTitleAttribute as CFString, &titleRef) // swiftlint:disable:this force_cast
        let subject = titleRef as? String

        return MailMetadata(subject: subject)
    }
}
