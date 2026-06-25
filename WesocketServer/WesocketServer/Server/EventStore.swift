import Foundation

/// Thread-safe, append-only JSONL store for CaptureEvents.
/// Keeps the last `maxMemory` events in RAM; appends every event to disk.
actor EventStore {
    private let storeURL: URL
    private let _screenshotsDirectory: URL
    private var recentEvents: [CaptureEvent] = []
    private let maxMemory = 500

    init(dataDir: URL) throws {
        storeURL = dataDir.appending(path: "events.jsonl")
        _screenshotsDirectory = dataDir.appending(path: "screenshots")
        let fm = FileManager.default
        try fm.createDirectory(at: dataDir, withIntermediateDirectories: true)
        try fm.createDirectory(at: _screenshotsDirectory, withIntermediateDirectories: true)
        loadRecent()
    }

    var screenshotsDirectory: URL { _screenshotsDirectory }

    func append(_ event: CaptureEvent) throws {
        let lineData = try JSONEncoder().encode(event) + Data("\n".utf8)
        let path = storeURL.path(percentEncoded: false)
        if FileManager.default.fileExists(atPath: path) {
            let fh = try FileHandle(forWritingTo: storeURL)
            defer { try? fh.close() }
            try fh.seekToEnd()
            try fh.write(contentsOf: lineData)
        } else {
            try lineData.write(to: storeURL)
        }
        recentEvents.append(event)
        if recentEvents.count > maxMemory {
            recentEvents.removeFirst(recentEvents.count - maxMemory)
        }
    }

    func recent(limit: Int = 200) -> [CaptureEvent] {
        Array(recentEvents.suffix(limit))
    }

    private func loadRecent() {
        guard let data = try? Data(contentsOf: storeURL) else { return }
        let text = String(data: data, encoding: .utf8) ?? ""
        let decoder = JSONDecoder()
        for line in text.components(separatedBy: "\n").suffix(maxMemory) where !line.isEmpty {
            if let event = try? decoder.decode(CaptureEvent.self, from: Data(line.utf8)) {
                recentEvents.append(event)
            }
        }
    }
}
