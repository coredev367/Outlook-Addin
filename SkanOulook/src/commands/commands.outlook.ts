/*
 * commands.outlook.ts — Ribbon button handler for SkanOulook.
 *
 * "Capture Snapshot" button (ExecuteFunction):
 *   - Reads the currently selected item in read mode
 *   - Sends a capture event directly to WesocketServer via wss://localhost:8765
 *   - Shows a brief notification confirming the capture
 *
 * Uses an ephemeral WebSocket (connect → send → ack → close) with a 4 s
 * safety timeout, the same pattern used by launchevents.ts.
 *
 * NOTE: LaunchEvents in launchevents.ts handle automatic Outlook events
 * (compose, send, recipients changed, etc.).
 * This button is for on-demand / manual snapshot capture only.
 */

/* global Office, WebSocket */

const CMD_SERVER_WSS  = "wss://localhost:8765";
const CMD_TIMEOUT_MS  = 4000;

// ── UUID ──────────────────────────────────────────────────────────────────────

function cmdUuid(): string {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
  });
}

// ── Ephemeral WebSocket send ──────────────────────────────────────────────────

function sendToServer(payload: object): Promise<void> {
  return new Promise((resolve) => {
    const timer = setTimeout(resolve, CMD_TIMEOUT_MS);
    const done  = () => { clearTimeout(timer); resolve(); };

    let ws: WebSocket;
    try {
      ws = new WebSocket(CMD_SERVER_WSS);
    } catch (_) {
      done();
      return;
    }

    ws.onopen    = () => ws.send(JSON.stringify(payload));
    ws.onmessage = (ev: MessageEvent) => {
      try {
        const msg = JSON.parse(ev.data as string);
        if (msg.type === "ack") { ws.close(); done(); }
      } catch (_) {}
    };
    ws.onerror   = () => done();
    ws.onclose   = () => done();
  });
}

// ── Capture current read-mode item ───────────────────────────────────────────

async function captureSnapshot(): Promise<void> {
  const item = Office.context.mailbox.item as any;
  if (!item) return;

  const isAppt = item.itemType === Office.MailboxEnums.ItemType.Appointment;

  const event = {
    id:            cmdUuid(),
    timestamp:     new Date().toISOString(),
    platform:      "web",
    appBundleId:   null,
    correlationId: null,
    action:        "select",
    itemKind:      isAppt ? "appointment" : "mail",
    metadata: {
      subject:     typeof item.subject === "string" ? item.subject : null,
      to:          (item.to  || []).map((r: any) => r.emailAddress),
      cc:          (item.cc  || []).map((r: any) => r.emailAddress),
      bcc:         (item.bcc || []).map((r: any) => r.emailAddress),
      attachments: (item.attachments || []).map((a: any) => a.name),
      body:        null,
    },
    widgetPath:     [],
    screenshotPath: null,
    clickPoint:     null,
    excelMetadata:  null,
  };

  await sendToServer({ type: "ingest", event });
}

// ── Ribbon button handler ─────────────────────────────────────────────────────

export function captureSnapshotCommand(event: Office.AddinCommands.Event): void {
  captureSnapshot().then(() => {
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "SkanOulookSnapshot",
      {
        type:       Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message:    "SkanOulook: snapshot sent to WesocketServer.",
        icon:       "Icon.80x80",
        persistent: false,
      }
    );
    event.completed();
  });
}

// ── Register with Office ──────────────────────────────────────────────────────

Office.onReady(async () => {
  Office.actions.associate("action", captureSnapshotCommand);
});
