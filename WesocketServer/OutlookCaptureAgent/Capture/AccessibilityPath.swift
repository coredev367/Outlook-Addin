import AppKit
import ApplicationServices

/// Walks the AX tree from a leaf element up to the window and returns
/// the ordered path as an array of WidgetNodes (leaf first, window last).
enum AccessibilityPath {

    static func build(from element: AXUIElement) -> [WidgetNode] {
        var path: [WidgetNode] = []
        var current: AXUIElement? = element

        while let el = current {
            let node = makeNode(from: el)
            path.append(node)

            // Stop at window or application root
            if node.role == kAXWindowRole as String ||
               node.role == kAXApplicationRole as String {
                break
            }

            // Prevent runaway loops
            if path.count > 20 { break }

            // Walk up to parent
            var parentRef: CFTypeRef?
            let result = AXUIElementCopyAttributeValue(el, kAXParentAttribute as CFString, &parentRef)
            guard result == .success, let parent = parentRef else { break }
            guard CFGetTypeID(parent) == AXUIElementGetTypeID() else { break }
            current = (parent as! AXUIElement) // swiftlint:disable:this force_cast
        }

        return path
    }

    // MARK: - Private helpers

    private static func makeNode(from el: AXUIElement) -> WidgetNode {
        WidgetNode(
            role:       stringAttr(el, kAXRoleAttribute)       ?? "unknown",
            subrole:    stringAttr(el, kAXSubroleAttribute),
            title:      stringAttr(el, kAXTitleAttribute),
            identifier: stringAttr(el, kAXIdentifierAttribute),
            frame:      rectAttr(el)
        )
    }

    private static func stringAttr(_ el: AXUIElement, _ attr: String) -> String? {
        var value: CFTypeRef?
        guard AXUIElementCopyAttributeValue(el, attr as CFString, &value) == .success,
              let s = value as? String else { return nil }
        return s
    }

    private static func rectAttr(_ el: AXUIElement) -> RectCodable? {
        var posRef: CFTypeRef?
        var sizeRef: CFTypeRef?
        guard AXUIElementCopyAttributeValue(el, kAXPositionAttribute as CFString, &posRef) == .success,
              AXUIElementCopyAttributeValue(el, kAXSizeAttribute as CFString, &sizeRef) == .success,
              let pv = posRef, CFGetTypeID(pv) == AXValueGetTypeID(),
              let sv = sizeRef, CFGetTypeID(sv) == AXValueGetTypeID()
        else { return nil }
        var pos  = CGPoint.zero
        var size = CGSize.zero
        AXValueGetValue(pv as! AXValue, .cgPoint, &pos)   // swiftlint:disable:this force_cast
        AXValueGetValue(sv as! AXValue, .cgSize,  &size)  // swiftlint:disable:this force_cast
        return RectCodable(x: pos.x, y: pos.y, width: size.width, height: size.height)
    }
}
