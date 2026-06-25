import AppKit

@NSApplicationMain
@MainActor
class AppDelegate: NSObject, NSApplicationDelegate {

    private var statusItem: NSStatusItem!
    private var captureController: CaptureController?
    private var wsClient: WebSocketClient!

    // MARK: - Launch

    func applicationDidFinishLaunching(_ notification: Notification) {
        buildStatusMenu()
        checkAndRequestPermissions()
        startCapture()
    }

    func applicationWillTerminate(_ notification: Notification) {
        captureController?.stop()
        wsClient.stop()
    }

    // MARK: - Menu bar

    private func buildStatusMenu() {
        statusItem = NSStatusBar.system.statusItem(withLength: NSStatusItem.variableLength)

        if let button = statusItem.button {
            button.image = NSImage(systemSymbolName: "eye.circle.fill",
                                   accessibilityDescription: "Outlook Capture")
            button.toolTip = "OutlookCaptureAgent"
        }

        updateMenu()
    }

    private func updateMenu() {
        let menu = NSMenu()

        let titleItem = NSMenuItem(title: "OutlookCaptureAgent", action: nil, keyEquivalent: "")
        titleItem.isEnabled = false
        menu.addItem(titleItem)

        let permItem = NSMenuItem(title: PermissionsManager.shared.statusSummary(),
                                  action: nil, keyEquivalent: "")
        permItem.isEnabled = false
        menu.addItem(permItem)

        menu.addItem(.separator())

        let capturing = captureController?.isRunning ?? false
        let toggleTitle = capturing ? "Pause Capture" : "Resume Capture"
        menu.addItem(NSMenuItem(title: toggleTitle,
                                action: #selector(toggleCapture),
                                keyEquivalent: ""))

        menu.addItem(.separator())

        menu.addItem(NSMenuItem(title: "Open Dashboard",
                                action: #selector(openDashboard),
                                keyEquivalent: "d"))

        menu.addItem(.separator())

        menu.addItem(NSMenuItem(title: "Check Permissions",
                                action: #selector(checkAndRequestPermissions),
                                keyEquivalent: ""))

        menu.addItem(.separator())

        menu.addItem(NSMenuItem(title: "Quit OutlookCaptureAgent",
                                action: #selector(NSApplication.terminate(_:)),
                                keyEquivalent: "q"))

        statusItem.menu = menu
    }

    // MARK: - Actions

    @objc private func toggleCapture() {
        guard let ctrl = captureController else { return }
        if ctrl.isRunning {
            ctrl.stop()
        } else {
            ctrl.start()
        }
        updateMenu()
    }

    @objc private func openDashboard() {
        NSWorkspace.shared.open(URL(string: "http://127.0.0.1:8080")!)
    }

    @objc private func checkAndRequestPermissions() {
        PermissionsManager.shared.requestAll()
        let hasAll = PermissionsManager.shared.hasAccessibility &&
                     PermissionsManager.shared.hasScreenRecording
        if !hasAll {
            PermissionsManager.shared.showMissingPermissionsAlert()
        }
        updateMenu()
    }

    // MARK: - Capture bootstrap

    private func startCapture() {
        wsClient = WebSocketClient(host: "127.0.0.1", port: 8765)
        wsClient.start()

        captureController = CaptureController(client: wsClient)
        captureController?.start()
    }
}
