import AppKit
import ApplicationServices

/// Watches Microsoft Outlook for AX notifications that signal user actions
/// which are NOT covered by click events (keyboard shortcuts, menu selections,
/// window lifecycle events).
final class AXObserverManager: @unchecked Sendable {

    /// Called on the main thread with a classified action and its context element (may be nil).
    var onAction: ((CaptureAction, ItemKind, AXUIElement?) -> Void)?

    private var axObserver: AXObserver?
    private var outlookPid: pid_t = 0
    private var watchTimer: Timer?

    private let notifications: [String] = [
        kAXWindowCreatedNotification as String,
        kAXUIElementDestroyedNotification as String,
        kAXSelectedRowsChangedNotification as String,
        kAXFocusedWindowChangedNotification as String,
        kAXMenuItemSelectedNotification as String,
    ]

    // MARK: - Lifecycle

    func start() {
        // Observe workspace app-activation changes
        NSWorkspace.shared.notificationCenter.addObserver(
            self,
            selector: #selector(appActivated(_:)),
            name: NSWorkspace.didActivateApplicationNotification,
            object: nil
        )
        // Poll every 5 s in case Outlook was already running at launch
        watchTimer = Timer.scheduledTimer(withTimeInterval: 5, repeats: true) { [weak self] _ in
            self?.refreshOutlookObserver()
        }
        refreshOutlookObserver()
    }

    func stop() {
        watchTimer?.invalidate()
        watchTimer = nil
        teardownObserver()
        NSWorkspace.shared.notificationCenter.removeObserver(self)
    }

    // MARK: - Observer management

    @objc private func appActivated(_ note: Notification) {
        refreshOutlookObserver()
    }

    private func refreshOutlookObserver() {
        let apps = NSWorkspace.shared.runningApplications
        guard let outlook = apps.first(where: {
            ($0.bundleIdentifier ?? "").lowercased().contains("microsoft.outlook")
        }) else { return }

        let pid = outlook.processIdentifier
        guard pid != outlookPid else { return }  // already watching this pid
        teardownObserver()
        setupObserver(pid: pid)
    }

    private func setupObserver(pid: pid_t) {
        guard AXIsProcessTrustedWithOptions(nil) else { return }

        var obs: AXObserver?
        let selfPtr = Unmanaged.passUnretained(self).toOpaque()

        let err = AXObserverCreate(pid, axNotificationCallback, &obs)
        guard err == .success, let observer = obs else {
            print("[AXObserver] Failed to create observer for pid \(pid): \(err.rawValue)")
            return
        }

        let appElement = AXUIElementCreateApplication(pid)
        for note in notifications {
            AXObserverAddNotification(observer, appElement, note as CFString, selfPtr)
        }

        CFRunLoopAddSource(
            CFRunLoopGetMain(),
            AXObserverGetRunLoopSource(observer),
            .commonModes
        )
        axObserver = observer
        outlookPid = pid
        print("[AXObserver] Watching Outlook pid \(pid)")
    }

    private func teardownObserver() {
        guard let obs = axObserver else { return }
        CFRunLoopRemoveSource(
            CFRunLoopGetMain(),
            AXObserverGetRunLoopSource(obs),
            .commonModes
        )
        axObserver = nil
        outlookPid = 0
    }

    // MARK: - Notification handler

    fileprivate func handleNotification(_ note: String, element: AXUIElement) {
        guard let (action, itemKind) = ActionClassifier.classify(
            notification: note,
            element: element
        ) else { return }

        // kAXFocusedWindowChangedNotification fires very frequently; only emit if it looks
        // like a new compose window (title contains expected keywords).
        if note == kAXFocusedWindowChangedNotification as String {
            var titleRef: CFTypeRef?
            AXUIElementCopyAttributeValue(element, kAXTitleAttribute as CFString, &titleRef)
            let title = (titleRef as? String ?? "").lowercased()
            guard title.contains("untitled") || title.contains("new message") ||
                  title.contains("forward") || title.contains("re:") else { return }
        }

        let captured = element
        DispatchQueue.main.async { [weak self] in
            self?.onAction?(action, itemKind, captured)
        }
    }
}

// MARK: - C callback

private let axNotificationCallback: AXObserverCallback = {
    _, element, notification, refcon in
    guard let refcon else { return }
    let manager = Unmanaged<AXObserverManager>.fromOpaque(refcon).takeUnretainedValue()
    manager.handleNotification(notification as String, element: element)
}
