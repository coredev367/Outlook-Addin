import AppKit
import ApplicationServices

/// Surfaces the three TCC permissions the agent needs and prompts the user when they are missing.
@MainActor
final class PermissionsManager {

    static let shared = PermissionsManager()
    private init() {}

    // MARK: - Query

    var hasAccessibility: Bool {
        AXIsProcessTrustedWithOptions(nil)
    }

    var hasScreenRecording: Bool {
        CGPreflightScreenCaptureAccess()
    }

    // Automation (AppleEvents to Outlook) can only be checked by attempting to send;
    // we surface it as a user reminder rather than a binary flag.

    // MARK: - Request all

    func requestAll() {
        if !hasAccessibility {
            promptForAccessibility()
        }
        if !hasScreenRecording {
            promptForScreenRecording()
        }
        // Automation consent is requested automatically by macOS on first NSAppleScript execution.
    }

    // MARK: - Status summary

    func statusSummary() -> String {
        let ax  = hasAccessibility   ? "✓" : "✗"
        let scr = hasScreenRecording ? "✓" : "✗"
        return "Accessibility: \(ax)  Screen Recording: \(scr)"
    }

    // MARK: - Prompts

    private func promptForAccessibility() {
        let opts: NSDictionary = [kAXTrustedCheckOptionPrompt.takeUnretainedValue() as String: true]
        AXIsProcessTrustedWithOptions(opts)
    }

    private func promptForScreenRecording() {
        CGRequestScreenCaptureAccess()
    }

    // MARK: - Alert for missing permissions

    func showMissingPermissionsAlert() {
        let alert = NSAlert()
        alert.messageText = "OutlookCaptureAgent needs additional permissions"
        alert.informativeText = """
            Please grant the following in System Settings → Privacy & Security:

            • Accessibility  — required to capture clicked UI elements
            • Screen Recording — required to take screenshots

            After granting, relaunch OutlookCaptureAgent.
            """
        alert.addButton(withTitle: "Open System Settings")
        alert.addButton(withTitle: "Later")
        if alert.runModal() == .alertFirstButtonReturn {
            NSWorkspace.shared.open(
                URL(string: "x-apple.systempreferences:com.apple.preference.security?Privacy")!
            )
        }
    }
}
