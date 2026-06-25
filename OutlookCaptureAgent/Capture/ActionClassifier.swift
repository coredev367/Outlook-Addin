import Foundation
import ApplicationServices

/// Maps an AX widget path or AX notification name to a `(CaptureAction, ItemKind)` pair.
enum ActionClassifier {

    // MARK: - From widget path (click events)

    static func classify(widgetPath: [WidgetNode]) -> (CaptureAction, ItemKind) {
        for node in widgetPath {
            if let result = classify(node: node) { return result }
        }
        // Fell through: a click on a list row / cell = selection
        let firstRole = widgetPath.first?.role ?? ""
        if firstRole == "AXRow" || firstRole == "AXCell" || firstRole == "AXStaticText" {
            return (.select, .mail)
        }
        return (.unknown, .unknown)
    }

    private static func classify(node: WidgetNode) -> (CaptureAction, ItemKind)? {
        let title = node.title?.trimmingCharacters(in: .whitespaces).lowercased() ?? ""
        let role  = node.role

        // --- Meeting responses ---
        if title.contains("accept") && !title.contains("decline") && !title.contains("tentative") {
            return (.meetingAccept, .meeting)
        }
        if title.contains("tentative") { return (.meetingTentative, .meeting) }
        if title.contains("decline")   { return (.meetingDecline, .meeting) }

        // --- Reply / Forward ---
        if title == "reply all" || title == "reply-all" || title.hasPrefix("reply all ") {
            return (.replyAll, .mail)
        }
        if title == "reply" || title.hasPrefix("reply ") { return (.reply, .mail) }
        if title.contains("forward")                     { return (.forward, .mail) }

        // --- Send ---
        if title == "send" || title == "send message" || title == "send email" {
            return (.send, .mail)
        }
        // Send button in compose windows has the paper-plane icon; title may be empty but role = AXButton
        if role == "AXButton", let id = node.identifier, id.lowercased().contains("send") {
            return (.send, .mail)
        }

        // --- Save ---
        if title.contains("save draft") || title == "save" { return (.save, .mail) }

        // --- Compose ---
        if title.contains("new email") || title.contains("new message") || title.contains("compose") {
            return (.compose, .mail)
        }
        if title.contains("new appointment") || title.contains("new event") {
            return (.compose, .appointment)
        }
        if title.contains("new meeting") || title.contains("schedule meeting") {
            return (.compose, .meeting)
        }

        // --- Close button ---
        if role == "AXButton" && node.subrole == "AXCloseButton" { return (.close, .unknown) }
        if role == "AXButton" && title == "close"                { return (.close, .unknown) }

        return nil
    }

    // MARK: - From AX notification name (non-click events)

    static func classify(notification: String, element: AXUIElement?) -> (CaptureAction, ItemKind)? {
        // AX notification name constants are CFString; extract the String value explicitly.
        let windowCreated   = kAXWindowCreatedNotification as String
        let elementDestroyed = kAXUIElementDestroyedNotification as String
        let rowsChanged     = kAXSelectedRowsChangedNotification as String
        let focusChanged    = kAXFocusedUIElementChangedNotification as String
        let menuSelected    = kAXMenuItemSelectedNotification as String

        switch notification {
        case windowCreated:
            return (.compose, .unknown)
        case elementDestroyed:
            return (.close, .unknown)
        case rowsChanged, focusChanged:
            return (.select, .mail)
        case menuSelected:
            if let el = element {
                let path = AccessibilityPath.build(from: el)
                if let result = path.first.flatMap({ classify(node: $0) }) { return result }
            }
            return nil
        default:
            return nil
        }
    }
}
