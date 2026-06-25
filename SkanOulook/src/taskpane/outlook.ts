/*
 * outlook.ts — Office.js event capture for SkanOulook add-in.
 *
 * Captures Outlook interactions and forwards them to WesocketServer
 * (ws://localhost:8765) using the ingest protocol.
 *
 * To monitor events in the terminal, run:
 *   node bridge.js
 */

/* global Office, document, WebSocket, setTimeout */

// ── Inline WebSocket bridge ───────────────────────────────────────────────────

const _WS_URL = "wss://localhost:8766"; // bridge.js proxy — wss avoids mixed-content block from https
const _RECONNECT_MS = 3000;

let _ws: WebSocket | null = null;
let _queue: string[] = [];
let _logHandler: ((type: string, detail: Record<string, unknown>) => void) | null = null;

function genUuid(): string {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
  });
}

function bridgeEmit(type: string, detail: Record<string, unknown> = {}): void {
  if (_logHandler) {
    try { _logHandler(type, detail); } catch (_) {}
  }
}

function bridgeConnect(): void {
  bridgeEmit("connecting");
  _ws = new WebSocket(_WS_URL);

  _ws.onopen = () => {
    bridgeEmit("connected", { url: _WS_URL });
    _queue.forEach((msg) => _ws!.send(msg));
    _queue = [];
  };

  _ws.onmessage = (ev) => {
    try {
      const msg = JSON.parse(ev.data);
      if (msg.type === "ack") {
        bridgeEmit("ack", { id: msg.event?.id });
      }
    } catch (_) {}
  };

  _ws.onerror = () => bridgeEmit("error");

  _ws.onclose = () => {
    bridgeEmit("disconnected", { retryMs: _RECONNECT_MS });
    _ws = null;
    setTimeout(bridgeConnect, _RECONNECT_MS);
  };
}

function bridgeSend(partial: Record<string, unknown>): void {
  const event: Record<string, unknown> = {
    id: genUuid(),
    timestamp: new Date().toISOString(),
    platform: "web",
    appBundleId: null,
    correlationId: null,
    widgetPath: [],
    screenshotPath: null,
    clickPoint: null,
    excelMetadata: null,
    ...partial,
  };
  const payload = JSON.stringify({ type: "ingest", event });

  if (_ws && _ws.readyState === WebSocket.OPEN) {
    _ws.send(payload);
    bridgeEmit("sent", { action: event.action, id: event.id, event });
  } else {
    _queue.push(payload);
    bridgeEmit("queued", { action: event.action, id: event.id, event });
  }
}

function setLogHandler(fn: typeof _logHandler): void {
  _logHandler = fn;
}

// ── In-pane logger ────────────────────────────────────────────────────────────

const ACTION_COLORS: Record<string, string> = {
  select:           "#0078d7",
  recipientsChange: "#ff8c00",
  attachmentChange: "#ff8c00",
  timeChange:       "#ff8c00",
  recurrenceChange: "#ff8c00",
  locationChange:   "#ff8c00",
  send:             "#107c10",
  compose:          "#5c2d91",
};

function now(): string {
  return new Date().toLocaleTimeString("en-US", { hour12: false });
}

function updateStatus(text: string, color: string): void {
  const dot   = document.getElementById("bridge-status-dot");
  const label = document.getElementById("bridge-status-text");
  if (dot)   dot.style.color    = color;
  if (label) label.textContent  = text;
}

function addLogEntry(badge: string, badgeColor: string, message: string, detail?: string): void {
  const log = document.getElementById("event-log");
  if (!log) return;

  const entry   = document.createElement("div");
  entry.className = "log-entry";

  const time    = document.createElement("span");
  time.className = "log-time";
  time.textContent = now();

  const badgeEl = document.createElement("span");
  badgeEl.className = "log-badge";
  badgeEl.textContent = badge;
  badgeEl.style.background = badgeColor;

  const msg = document.createElement("span");
  msg.className = "log-msg";
  msg.textContent = message;

  entry.appendChild(time);
  entry.appendChild(badgeEl);
  entry.appendChild(msg);

  if (detail) {
    const detailEl = document.createElement("div");
    detailEl.className = "log-detail";
    detailEl.textContent = detail;
    entry.appendChild(detailEl);
  }

  log.insertBefore(entry, log.firstChild);
}

function handleBridgeLog(type: string, detail: Record<string, unknown>): void {
  switch (type) {
    case "connecting":
      updateStatus(`⬤ Connecting to ${_WS_URL}…`, "#888");
      addLogEntry("WS", "#888", `Connecting to ${_WS_URL}…`);
      break;
    case "connected":
      updateStatus(`⬤ Connected  (${_WS_URL})`, "#107c10");
      addLogEntry("WS", "#107c10", `Connected to ${_WS_URL}`);
      break;
    case "disconnected":
      updateStatus("⬤ Disconnected – reconnecting…", "#a80000");
      addLogEntry("WS", "#a80000", `Disconnected from ${_WS_URL} – retrying in ${detail.retryMs}ms`);
      break;
    case "error":
      updateStatus(`⬤ Cannot reach ${_WS_URL}`, "#a80000");
      addLogEntry("WS", "#a80000", `Connection error → ${_WS_URL}`);
      break;
    case "sent": {
      const ev   = detail.event as Record<string, unknown>;
      const meta = ev?.metadata as Record<string, unknown> | undefined;
      addLogEntry(
        String(detail.action),
        ACTION_COLORS[String(detail.action)] || "#333",
        String(meta?.subject || "(no subject)"),
        `id: ${detail.id}`
      );
      break;
    }
    case "queued": {
      const ev   = detail.event as Record<string, unknown>;
      const meta = ev?.metadata as Record<string, unknown> | undefined;
      addLogEntry(
        String(detail.action),
        "#ff8c00",
        `[queued] ${meta?.subject || "(no subject)"}`,
        `id: ${detail.id}`
      );
      break;
    }
    case "ack":
      addLogEntry("ACK", "#00b7c3", `Server acknowledged ${detail.id}`);
      break;
  }
}

// ── Office initialisation ─────────────────────────────────────────────────────

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  const sideloadMsg = document.getElementById("sideload-msg");
  const appBody     = document.getElementById("app-body");
  if (sideloadMsg) sideloadMsg.style.display = "none";
  if (appBody)     appBody.style.display     = "flex";

  // Wire log → pane, then start WebSocket
  setLogHandler(handleBridgeLog);
  bridgeConnect();

  // Clear button
  const clearBtn = document.getElementById("clear-log");
  if (clearBtn) {
    clearBtn.onclick = () => {
      const log = document.getElementById("event-log");
      if (log) log.innerHTML = "";
    };
  }

  // Capture the item that is currently selected when the pane first opens.
  // This is valid: the task pane only loads because Outlook is showing an
  // email, so an email IS selected — it is a real Outlook context.
  captureCurrentItem("select");

  // ItemChanged: fires when the user clicks a DIFFERENT email in the list.
  // This is the primary Outlook event captured by the task pane.
  if (Office.context.requirements.isSetSupported("Mailbox", "1.5")) {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      handleItemChanged,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          addLogEntry("ERR", "#a80000", "ItemChanged registration failed: " + result.error.message);
        } else {
          addLogEntry("OK", "#107c10", "Listening for ItemChanged events");
        }
      }
    );
  } else {
    addLogEntry("WARN", "#ff8c00", "Mailbox 1.5 not supported – ItemChanged unavailable");
  }

  // Compose-mode field changes (Mailbox 1.7+)
  if (Office.context.requirements.isSetSupported("Mailbox", "1.7")) {
    registerComposeHandlers();
  }
});

// ── Event: item changed ───────────────────────────────────────────────────────

function handleItemChanged(_ev: object): void {
  captureCurrentItem("select");
  if (Office.context.requirements.isSetSupported("Mailbox", "1.7")) {
    registerComposeHandlers();
  }
}

// ── Compose-mode field handlers ───────────────────────────────────────────────

function registerComposeHandlers(): void {
  const item = Office.context.mailbox.item;
  if (!item) return;

  const add = (eventType: Office.EventType, action: string) => {
    item.addHandlerAsync(eventType, () => captureCurrentItem(action), () => {});
  };

  add(Office.EventType.RecipientsChanged,       "recipientsChange");
  add(Office.EventType.AttachmentsChanged,      "attachmentChange");
  add(Office.EventType.AppointmentTimeChanged,  "timeChange");
  add(Office.EventType.RecurrenceChanged,       "recurrenceChange");
  add(Office.EventType.EnhancedLocationsChanged,"locationChange");
}

// ── Core capture logic ────────────────────────────────────────────────────────

function captureCurrentItem(action: string): void {
  const item = Office.context.mailbox.item as any;
  if (!item) return;

  const isAppointment = item.itemType === Office.MailboxEnums.ItemType.Appointment;
  const itemKind = isAppointment ? "appointment" : "mail";

  if (typeof item.subject === "string") {
    // Read mode – properties are synchronous
    bridgeSend({ action, itemKind, metadata: buildReadMetadata(item) });
  } else {
    // Compose mode – properties are async
    readComposeMetadata(item).then((metadata) => {
      bridgeSend({ action, itemKind, metadata });
    });
  }
}

function buildReadMetadata(item: any): object {
  return {
    subject:     item.subject ?? null,
    to:          (item.to          || []).map((r: any) => r.emailAddress),
    cc:          (item.cc          || []).map((r: any) => r.emailAddress),
    bcc:         (item.bcc         || []).map((r: any) => r.emailAddress),
    attachments: (item.attachments || []).map((a: any) => a.name),
    body:        null,
  };
}

function asyncGet<T>(fn: (cb: (r: Office.AsyncResult<T>) => void) => void): Promise<T | undefined> {
  return new Promise((resolve) => {
    fn((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : undefined));
  });
}

async function readComposeMetadata(item: any): Promise<object> {
  const [subject, toR, ccR, bccR] = await Promise.all([
    asyncGet<string>((cb) => item.subject.getAsync(cb)),
    asyncGet<Office.EmailAddressDetails[]>((cb) => item.to.getAsync(cb)),
    asyncGet<Office.EmailAddressDetails[]>((cb) => item.cc.getAsync(cb)),
    item.bcc
      ? asyncGet<Office.EmailAddressDetails[]>((cb) => item.bcc.getAsync(cb))
      : Promise.resolve([] as Office.EmailAddressDetails[]),
  ]);
  return {
    subject:     subject ?? null,
    to:          (toR ?? []).map((r) => r.emailAddress),
    cc:          (ccR ?? []).map((r) => r.emailAddress),
    bcc:         (bccR ?? []).map((r) => r.emailAddress),
    attachments: (item.attachments || []).map((a: any) => a.name),
    body:        null,
  };
}
