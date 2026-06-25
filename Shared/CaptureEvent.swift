import Foundation

// MARK: - Action enum

public enum CaptureAction: String, Codable, Sendable {
    // ── Outlook — navigation ──────────────────────────────────────────────────
    case select           // message/item selected in the list

    // ── Outlook — compose actions ─────────────────────────────────────────────
    case reply            // Reply button
    case replyAll         // Reply All button
    case forward          // Forward button
    case send             // Send button / Cmd+Return
    case save             // Save / Save Draft
    case compose          // New Email, New Appointment, New Meeting opened
    case close            // compose/inspector window closed

    // ── Outlook — compose field changes (LaunchEvents / item handlers) ────────
    case recipientsChange  // To / Cc / Bcc / Attendees added or removed
    case attachmentChange  // Attachment added or removed
    case fromChange        // Sender account (From) changed
    case timeChange        // Meeting start or end time changed
    case recurrenceChange  // Recurrence pattern changed
    case locationChange    // Location added or removed

    // ── Outlook — meeting responses ───────────────────────────────────────────
    case meetingAccept    // meeting response: Accept
    case meetingTentative // meeting response: Tentative
    case meetingDecline   // meeting response: Decline

    // ── Outlook — input monitoring (task pane) ────────────────────────────────
    case shortcutKey      // keyboard shortcut fired inside task pane

    // ── Excel ──────────────────────────────────────────────────────────────────
    case open             // workbook opened
    case activate         // workbook/window activated (user switched to it)
    case sheetActivate    // worksheet tab switched
    case selectionChange  // cell / range selection changed
    case edit             // cell value changed
    case doubleClick      // cell double-clicked (enter edit mode heuristic)
    case rightClick       // right-click context menu (best-effort)
    case macroClick       // form-control / macro button clicked (via COM bridge)

    case unknown
}

// MARK: - Item kind

public enum ItemKind: String, Codable, Sendable {
    // Outlook
    case mail
    case appointment
    case meeting

    // Excel
    case workbook
    case sheet
    case cell
    case range
    case table

    case unknown
}

// MARK: - Flexible JSON value (used for cell values / formulas)

/// Represents any JSON scalar or container. Used for Excel cell values which
/// can be strings, numbers, booleans, or null.
public indirect enum JSONValue: Codable, Sendable {
    case string(String)
    case number(Double)
    case bool(Bool)
    case null
    case array([JSONValue])
    case object([String: JSONValue])

    public init(from decoder: Decoder) throws {
        let c = try decoder.singleValueContainer()
        if c.decodeNil()                         { self = .null;              return }
        if let v = try? c.decode(Bool.self)      { self = .bool(v);           return }
        if let v = try? c.decode(Double.self)    { self = .number(v);         return }
        if let v = try? c.decode(String.self)    { self = .string(v);         return }
        if let v = try? c.decode([JSONValue].self)          { self = .array(v);  return }
        if let v = try? c.decode([String: JSONValue].self)  { self = .object(v); return }
        throw DecodingError.dataCorrupted(
            .init(codingPath: decoder.codingPath,
                  debugDescription: "Unknown JSON value type"))
    }

    public func encode(to encoder: Encoder) throws {
        var c = encoder.singleValueContainer()
        switch self {
        case .string(let v): try c.encode(v)
        case .number(let v): try c.encode(v)
        case .bool(let v):   try c.encode(v)
        case .null:          try c.encodeNil()
        case .array(let v):  try c.encode(v)
        case .object(let v): try c.encode(v)
        }
    }
}

// MARK: - Excel sub-structures

public struct ExcelCellFont: Codable, Sendable {
    public let name: String?
    public let size: Double?
    public let bold: Bool?
    public let italic: Bool?
    public let underline: String?   // "None" | "Single" | "Double" | ...
    public let color: String?       // hex, e.g. "#FF0000"

    public init(
        name: String? = nil, size: Double? = nil,
        bold: Bool? = nil, italic: Bool? = nil,
        underline: String? = nil, color: String? = nil
    ) {
        self.name = name; self.size = size
        self.bold = bold; self.italic = italic
        self.underline = underline; self.color = color
    }
}

public struct ExcelTableColumn: Codable, Sendable {
    public let name: String
    public let index: Int

    public init(name: String, index: Int) {
        self.name = name; self.index = index
    }
}

public struct ExcelTableMetadata: Codable, Sendable {
    public let name: String
    public let address: String?
    public let sheetName: String?
    public let columns: [ExcelTableColumn]
    public let rowCount: Int
    public let headerRowCount: Int

    public init(
        name: String,
        address: String? = nil,
        sheetName: String? = nil,
        columns: [ExcelTableColumn] = [],
        rowCount: Int = 0,
        headerRowCount: Int = 1
    ) {
        self.name = name; self.address = address; self.sheetName = sheetName
        self.columns = columns; self.rowCount = rowCount
        self.headerRowCount = headerRowCount
    }
}

/// Metadata captured from Excel workbook / sheet / cell interactions.
public struct ExcelMetadata: Codable, Sendable {
    public let workbookName: String?
    public let sheetName: String?
    public let selectedRange: String?           // e.g. "Sheet1!A1:B5"
    public let cellValues: [[JSONValue?]]?       // row-major; null = empty cell
    public let cellFormulas: [[String?]]?        // row-major formula strings
    public let cellFonts: [[ExcelCellFont?]]?    // row-major font info
    public let tables: [ExcelTableMetadata]?     // all tables in workbook
    public let visibleRange: String?             // address of the visible used range
    public let screenshotBase64: String?         // base64 PNG of selected range (transient)
    public let shortcutKey: String?              // e.g. "Ctrl+Shift+C"
    public let macroName: String?               // macro / form-control name

    public init(
        workbookName: String? = nil,
        sheetName: String? = nil,
        selectedRange: String? = nil,
        cellValues: [[JSONValue?]]? = nil,
        cellFormulas: [[String?]]? = nil,
        cellFonts: [[ExcelCellFont?]]? = nil,
        tables: [ExcelTableMetadata]? = nil,
        visibleRange: String? = nil,
        screenshotBase64: String? = nil,
        shortcutKey: String? = nil,
        macroName: String? = nil
    ) {
        self.workbookName = workbookName; self.sheetName = sheetName
        self.selectedRange = selectedRange; self.cellValues = cellValues
        self.cellFormulas = cellFormulas; self.cellFonts = cellFonts
        self.tables = tables; self.visibleRange = visibleRange
        self.screenshotBase64 = screenshotBase64
        self.shortcutKey = shortcutKey; self.macroName = macroName
    }
}

// MARK: - Shared sub-structures (Outlook)

public struct RectCodable: Codable, Sendable, Equatable {
    public let x: Double
    public let y: Double
    public let width: Double
    public let height: Double

    public init(x: Double, y: Double, width: Double, height: Double) {
        self.x = x; self.y = y; self.width = width; self.height = height
    }
}

public struct PointCodable: Codable, Sendable, Equatable {
    public let x: Double
    public let y: Double

    public init(x: Double, y: Double) { self.x = x; self.y = y }
}

/// One node in the accessibility widget path (leaf → window).
public struct WidgetNode: Codable, Sendable {
    public let role: String
    public let subrole: String?
    public let title: String?
    public let identifier: String?
    public let frame: RectCodable?

    public init(
        role: String,
        subrole: String? = nil,
        title: String? = nil,
        identifier: String? = nil,
        frame: RectCodable? = nil
    ) {
        self.role = role
        self.subrole = subrole
        self.title = title
        self.identifier = identifier
        self.frame = frame
    }
}

/// Email / appointment / meeting metadata (Outlook).
public struct MailMetadata: Codable, Sendable {
    public let subject: String?
    public let to: [String]
    public let cc: [String]
    public let bcc: [String]
    public let attachments: [String]   // filenames only
    public let body: String?           // nil unless captureBody is enabled

    public init(
        subject: String? = nil,
        to: [String] = [],
        cc: [String] = [],
        bcc: [String] = [],
        attachments: [String] = [],
        body: String? = nil
    ) {
        self.subject = subject
        self.to = to; self.cc = cc; self.bcc = bcc
        self.attachments = attachments; self.body = body
    }

    public static let empty = MailMetadata()
}

// MARK: - CaptureEvent (wire format, all platforms)

/// Universal record emitted for every captured interaction across all platforms.
public struct CaptureEvent: Codable, Sendable {
    public let id: String             // UUID string
    public let timestamp: String      // ISO 8601
    public let platform: String       // "macos" | "windows" | "web"
    public let appBundleId: String?
    public let correlationId: String?
    public let action: CaptureAction
    public let itemKind: ItemKind
    // ── Outlook ──────────────────────────────────────────────────────────────
    public let metadata: MailMetadata?
    public let widgetPath: [WidgetNode]
    public let screenshotPath: String?  // server-relative: "screenshots/<file>.png"
    public let clickPoint: PointCodable?
    // ── Excel ─────────────────────────────────────────────────────────────────
    public let excelMetadata: ExcelMetadata?

    public init(
        id: String = UUID().uuidString,
        timestamp: String = ISO8601DateFormatter().string(from: Date()),
        platform: String = "macos",
        appBundleId: String? = nil,
        correlationId: String? = nil,
        action: CaptureAction,
        itemKind: ItemKind = .unknown,
        metadata: MailMetadata? = nil,
        widgetPath: [WidgetNode] = [],
        screenshotPath: String? = nil,
        clickPoint: PointCodable? = nil,
        excelMetadata: ExcelMetadata? = nil
    ) {
        self.id = id
        self.timestamp = timestamp
        self.platform = platform
        self.appBundleId = appBundleId
        self.correlationId = correlationId
        self.action = action
        self.itemKind = itemKind
        self.metadata = metadata
        self.widgetPath = widgetPath
        self.screenshotPath = screenshotPath
        self.clickPoint = clickPoint
        self.excelMetadata = excelMetadata
    }
}

// MARK: - WebSocket message envelope

public enum WSMessageType: String, Codable, Sendable {
    case ingest      // agent  → server: new event to store
    case subscribe   // browser → server: begin live feed
    case broadcast   // server → browser: new event available
    case ack         // server → agent: event received OK
}

public struct WSMessage: Codable, Sendable {
    public let type: WSMessageType
    public let event: CaptureEvent?

    public init(type: WSMessageType, event: CaptureEvent? = nil) {
        self.type = type
        self.event = event
    }
}
