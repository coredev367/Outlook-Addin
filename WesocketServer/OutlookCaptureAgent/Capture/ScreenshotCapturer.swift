import AppKit
import ScreenCaptureKit
import CoreGraphics
import ImageIO
import UniformTypeIdentifiers

/// Captures a PNG screenshot of the Outlook window (or the whole screen as fallback)
/// and writes it to the shared OutlookCapture/screenshots/ directory.
actor ScreenshotCapturer {

    static let shared = ScreenshotCapturer()
    private init() {}

    private let saveDir: URL = {
        let base = FileManager.default.urls(for: .applicationSupportDirectory, in: .userDomainMask)[0]
            .appending(path: "OutlookCapture/screenshots")
        try? FileManager.default.createDirectory(at: base, withIntermediateDirectories: true)
        return base
    }()

    // MARK: - Public

    /// Captures a screenshot and returns the server-relative path "screenshots/<id>.png",
    /// or nil if Screen Recording permission is not granted or capture fails.
    func capture(correlationId: String) async -> String? {
        guard CGPreflightScreenCaptureAccess() else {
            print("[Screenshot] Screen Recording permission not granted")
            return nil
        }

        return await captureViaSCKit(correlationId: correlationId)
    }

    // MARK: - ScreenCaptureKit path (macOS 14+)

    private func captureViaSCKit(correlationId: String) async -> String? {
        do {
            let content = try await SCShareableContent.excludingDesktopWindows(
                false, onScreenWindowsOnly: true
            )

            // Prefer the frontmost Outlook window
            let outlookWindow = content.windows.first { window in
                (window.owningApplication?.bundleIdentifier ?? "")
                    .lowercased().contains("microsoft.outlook")
            }

            let filter: SCContentFilter
            if let win = outlookWindow {
                filter = SCContentFilter(desktopIndependentWindow: win)
            } else if let display = content.displays.first {
                filter = SCContentFilter(display: display,
                                         excludingApplications: [],
                                         exceptingWindows: [])
            } else {
                return nil
            }

            let config = SCStreamConfiguration()
            // Use the window's own pixel dimensions when available
            if let win = outlookWindow {
                let scale = Int(NSScreen.main?.backingScaleFactor ?? 2)
                config.width  = Int(win.frame.width)  * scale
                config.height = Int(win.frame.height) * scale
            }
            config.showsCursor = false

            let image = try await SCScreenshotManager.captureImage(
                contentFilter: filter,
                configuration: config
            )
            return write(image: image, name: correlationId)
        } catch {
            print("[Screenshot] SCKit error: \(error)")
            return nil
        }
    }

    // Note: CGWindowListCreateImage was removed in macOS 26; SCKit is the only supported path.

    // MARK: - Write helper

    private func write(image: CGImage, name: String) -> String? {
        let filename = "\(name).png"
        let fileURL  = saveDir.appending(path: filename)
        guard let dest = CGImageDestinationCreateWithURL(
            fileURL as CFURL, UTType.png.identifier as CFString, 1, nil
        ) else { return nil }
        CGImageDestinationAddImage(dest, image, nil)
        guard CGImageDestinationFinalize(dest) else { return nil }
        return "screenshots/\(filename)"
    }
}
