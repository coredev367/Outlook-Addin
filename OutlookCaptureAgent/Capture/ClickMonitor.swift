import AppKit
import ApplicationServices

/// Installs a passive CGEventTap that fires on every left-mouse-down.
/// Events are ignored unless Microsoft Outlook is the frontmost application.
final class ClickMonitor: @unchecked Sendable {

    /// Called on the main queue with the click position and the AX element hit at that position
    /// in the Outlook process.
    var onOutlookClick: ((CGPoint, AXUIElement?) -> Void)?

    private var tapRef: CFMachPort?
    private var runLoopSource: CFRunLoopSource?

    // MARK: - Lifecycle

    func start() {
        guard tapRef == nil else { return }

        let mask = CGEventMask(1 << CGEventType.leftMouseDown.rawValue)
        let selfPtr = Unmanaged.passUnretained(self).toOpaque()

        guard let tap = CGEvent.tapCreate(
            tap: .cgSessionEventTap,
            place: .tailAppendEventTap,
            options: .listenOnly,
            eventsOfInterest: mask,
            callback: clickCallback,
            userInfo: selfPtr
        ) else {
            print("[ClickMonitor] Could not create event tap — Accessibility permission missing?")
            return
        }

        tapRef = tap
        runLoopSource = CFMachPortCreateRunLoopSource(nil, tap, 0)
        CFRunLoopAddSource(CFRunLoopGetMain(), runLoopSource, .commonModes)
        CGEvent.tapEnable(tap: tap, enable: true)
        print("[ClickMonitor] Event tap active")
    }

    func stop() {
        if let tap = tapRef {
            CGEvent.tapEnable(tap: tap, enable: false)
            if let src = runLoopSource {
                CFRunLoopRemoveSource(CFRunLoopGetMain(), src, .commonModes)
            }
        }
        tapRef = nil
        runLoopSource = nil
    }

    // MARK: - Internal handler

    fileprivate func handleCGEvent(_ event: CGEvent) {
        // Only care about Outlook
        let app = NSWorkspace.shared.frontmostApplication
        guard let bundle = app?.bundleIdentifier,
              bundle.lowercased().contains("microsoft.outlook") ||
              bundle.lowercased().contains("com.microsoft.outlook") else { return }

        let point = event.location
        let pid = app!.processIdentifier
        let appElement = AXUIElementCreateApplication(pid)

        // Hit-test at the click position
        var element: AXUIElement?
        AXUIElementCopyElementAtPosition(appElement, Float(point.x), Float(point.y), &element)

        let captured = element
        let capturedPoint = point
        DispatchQueue.main.async { [weak self] in
            self?.onOutlookClick?(capturedPoint, captured)
        }
    }
}

// MARK: - C callback (must be a free function / @convention(c) closure)

private let clickCallback: CGEventTapCallBack = { _, _, event, refcon in
    guard let refcon else { return Unmanaged.passUnretained(event) }
    let monitor = Unmanaged<ClickMonitor>.fromOpaque(refcon).takeUnretainedValue()
    monitor.handleCGEvent(event)
    return Unmanaged.passUnretained(event)
}
