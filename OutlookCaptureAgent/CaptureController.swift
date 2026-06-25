import AppKit
import ApplicationServices

/// Orchestrates all capture subsystems: receives callbacks from ClickMonitor and
/// AXObserverManager, assembles a full CaptureEvent (widget path + metadata +
/// screenshot), and sends it via WebSocketClient.
@MainActor
final class CaptureController {

    private let clickMonitor  = ClickMonitor()
    private let axObserver    = AXObserverManager()
    private let wsClient: WebSocketClient

    // Forwarded from AppDelegate
    var isRunning: Bool = false

    init(client: WebSocketClient) {
        self.wsClient = client
        setupCallbacks()
    }

    // MARK: - Lifecycle

    func start() {
        guard !isRunning else { return }
        isRunning = true
        clickMonitor.start()
        axObserver.start()
        print("[CaptureController] Started")
    }

    func stop() {
        guard isRunning else { return }
        isRunning = false
        clickMonitor.stop()
        axObserver.stop()
        print("[CaptureController] Stopped")
    }

    // MARK: - Callbacks

    private func setupCallbacks() {
        // Click events in Outlook
        clickMonitor.onOutlookClick = { [weak self] point, element in
            self?.handleClick(at: point, element: element)
        }

        // AX-observer events (keyboard shortcuts, menu items, window lifecycle)
        axObserver.onAction = { [weak self] action, itemKind, element in
            self?.handleAXAction(action, itemKind: itemKind, element: element)
        }
    }

    // MARK: - Event assembly

    private func handleClick(at point: CGPoint, element: AXUIElement?) {
        let widgetPath = element.map { AccessibilityPath.build(from: $0) } ?? []
        let (action, itemKind) = ActionClassifier.classify(widgetPath: widgetPath)
        let correlationId = UUID().uuidString
        let clickPoint = PointCodable(x: point.x, y: point.y)

        Task {
            let metadata = OutlookScripting.shared.currentMetadata()
            let shotPath = await ScreenshotCapturer.shared.capture(correlationId: correlationId)
            let event = CaptureEvent(
                correlationId:  correlationId,
                action:         action,
                itemKind:       itemKind,
                metadata:       metadata,
                widgetPath:     widgetPath,
                screenshotPath: shotPath,
                clickPoint:     clickPoint
            )
            wsClient.send(event)
            print("[Capture] \(action.rawValue) [\(itemKind.rawValue)] \(metadata?.subject ?? "")")
        }
    }

    private func handleAXAction(_ action: CaptureAction, itemKind: ItemKind, element: AXUIElement?) {
        let widgetPath = element.map { AccessibilityPath.build(from: $0) } ?? []
        let correlationId = UUID().uuidString

        Task {
            let metadata = OutlookScripting.shared.currentMetadata()
            let shotPath = await ScreenshotCapturer.shared.capture(correlationId: correlationId)
            let event = CaptureEvent(
                correlationId:  correlationId,
                action:         action,
                itemKind:       itemKind,
                metadata:       metadata,
                widgetPath:     widgetPath,
                screenshotPath: shotPath
            )
            wsClient.send(event)
            print("[Capture/AX] \(action.rawValue) [\(itemKind.rawValue)] \(metadata?.subject ?? "")")
        }
    }
}
