# OutlookCapture Ingest Protocol

This document is the normative contract between any capture agent
(macOS, Windows, Web) and the `WesocketServer` ingest server.

---

## Transport

| Parameter | Value |
|-----------|-------|
| Protocol  | WebSocket (RFC 6455) |
| Default address | `ws://127.0.0.1:8765` |
| Encoding | UTF-8 JSON, text frames only |

---

## Connection roles

A client identifies itself by the first message it sends after the WebSocket
handshake completes:

| First `type` | Role | Direction |
|---|---|---|
| `"ingest"` | Capture agent sending a new event | client → server |
| `"subscribe"` | Dashboard browser subscribing to live updates | client → server |

---

## Messages

### `ingest` — agent → server

```json
{
  "type": "ingest",
  "event": { ... }   // CaptureEvent object, see schema below
}
```

### `ack` — server → agent

```json
{ "type": "ack" }
```

Sent immediately after a valid `ingest` is processed.

### `subscribe` — browser → server

```json
{ "type": "subscribe" }
```

No `event` field; no response from server.

### `broadcast` — server → subscribed browsers

```json
{
  "type": "broadcast",
  "event": { ... }   // CaptureEvent object
}
```

---

## CaptureEvent schema

```jsonc
{
  // --- Identity ---
  "id":            "550e8400-e29b-41d4-a716-446655440000",  // UUID
  "timestamp":     "2026-06-21T20:00:00Z",                 // ISO 8601
  "platform":      "macos",    // "macos" | "windows" | "web"
  "appBundleId":   "com.microsoft.Outlook",
  "correlationId": "abc-123",  // optional; ties a click to its screenshot

  // --- Action & item ---
  "action": "reply",        // see Action enum below
  "itemKind": "mail",       // see ItemKind enum below

  // --- Email / meeting metadata ---
  "metadata": {
    "subject":     "Re: Project update",
    "to":          ["alice@example.com"],
    "cc":          ["bob@example.com"],
    "bcc":         [],
    "attachments": ["report.pdf"],
    "body":        null       // null unless server-side captureBody is enabled
  },

  // --- UI widget path (leaf → window) ---
  "widgetPath": [
    {
      "role":       "AXButton",
      "subrole":    null,
      "title":      "Reply",
      "identifier": "replyButton",
      "frame":      { "x": 120, "y": 45, "width": 80, "height": 28 }
    },
    {
      "role":  "AXToolbar",
      "title": "Message Toolbar"
    }
  ],

  // --- Screenshot ---
  "screenshotPath": "screenshots/abc-123.png",   // null if capture failed
  "clickPoint":     { "x": 160.0, "y": 59.0 }   // null for non-click events
}
```

### Action enum values

| Value | Meaning |
|---|---|
| `select` | Item selected in the message list |
| `reply` | Reply button / Cmd+R |
| `replyAll` | Reply All / Cmd+Shift+R |
| `forward` | Forward / Cmd+J |
| `send` | Send button / Cmd+Return |
| `save` | Save Draft |
| `compose` | New compose / appointment / meeting window opened |
| `meetingAccept` | Meeting invitation: Accept |
| `meetingTentative` | Meeting invitation: Tentative |
| `meetingDecline` | Meeting invitation: Decline |
| `close` | Compose/inspector window closed |
| `unknown` | Could not classify |

### ItemKind enum values

`mail` | `appointment` | `meeting` | `unknown`

---

## HTTP API (same host, port 8080)

| Method | Path | Response |
|--------|------|----------|
| `GET` | `/` | Dashboard HTML (live feed) |
| `GET` | `/api/events` | JSON array of the last 500 `CaptureEvent` objects |
| `GET` | `/screenshots/<file>.png` | PNG screenshot file |

All responses include `Access-Control-Allow-Origin: *`.

---

## Phase-2 extensions (not implemented in v1)

- **Windows agent** — UI Automation (IUIAutomation) + Win32 screenshot, same JSON schema, same WebSocket endpoint.
- **Web / new Outlook** — Office.js add-in posts `action`, `itemKind`, `metadata` (no screenshots); `platform = "web"`, `widgetPath = []`.
