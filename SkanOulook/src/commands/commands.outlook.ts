/*
 * commands.outlook.ts — Ribbon button handler for SkanOulook.
 *
 * "Capture Snapshot" button (ExecuteFunction):
 *   - Reads the currently selected item in read mode
 *   - POSTs a capture event to bridge.js as a manual snapshot
 *   - Shows a brief notification confirming the capture
 *
 * NOTE: LaunchEvents in launchevents.ts handle the real Outlook actions
 * (compose, send, recipients changed, etc.) automatically.
 * This button is for on-demand / manual capture only.
 */

/* global Office */

const CMD_BRIDGE_URL = "https://localhost:8766/event"; // HTTP POST (wss for task pane)

// ── UUID ──────────────────────────────────────────────────────────────────────

function cmdUuid(): string {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
  });
}

// ── Capture current read-mode item via HTTP POST ──────────────────────────────

function captureSnapshot(): void {
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

  // Fire-and-forget HTTP POST — bridge.js handles it
  fetch(CMD_BRIDGE_URL, {
    method:  "POST",
    headers: { "Content-Type": "application/json" },
    body:    JSON.stringify({ type: "ingest", event }),
  }).catch(() => { /* bridge.js not running — silent */ });
}

// ── Ribbon button handler ─────────────────────────────────────────────────────

export function captureSnapshotCommand(event: Office.AddinCommands.Event): void {
  captureSnapshot();

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "SkanOulookSnapshot",
    {
      type:       Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message:    "SkanOulook: snapshot sent to bridge.",
      icon:       "Icon.80x80",
      persistent: false,
    }
  );

  event.completed();
}

// ── Register with Office ──────────────────────────────────────────────────────

Office.onReady(async () => {
  Office.actions.associate("action", captureSnapshotCommand);
});
