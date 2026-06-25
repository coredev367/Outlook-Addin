/*
 * launchevents.ts — Event-based activation handlers for SkanOulook.
 *
 * Connects directly to WesocketServer via wss://localhost:8765.
 * (No bridge.js required.)
 *
 * Protocol:
 *   send → { "type": "ingest", "event": { ...CaptureEvent } }
 *   recv ← { "type": "ack",   "event": { "id": "..."     } }
 *
 * Each handler reads item properties, opens an ephemeral WebSocket,
 * sends the event, waits for ack, then calls event.completed().
 * A 4 s safety timeout ensures completed() is always called.
 */

/* global Office, WebSocket */

export {}; // module scope — prevents symbol clashes

/** Minimal shape of the event object Outlook passes to LaunchEvent handlers. */
interface LaunchEvent {
  completed(options?: { allowEvent?: boolean }): void;
}

const SERVER_WSS   = "wss://localhost:8765"; // WesocketServer direct
const SEND_TIMEOUT = 4000;                   // safety margin under Outlook's ~5 s limit

// ── UUID ──────────────────────────────────────────────────────────────────────

function uuid(): string {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
  });
}

// ── Async property helper ─────────────────────────────────────────────────────

function getAsync<T>(fn: (cb: (r: Office.AsyncResult<T>) => void) => void): Promise<T | undefined> {
  return new Promise((resolve) => {
    fn((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : undefined));
  });
}

// ── WebSocket send-and-ack ────────────────────────────────────────────────────

/**
 * Open an ephemeral WebSocket to WesocketServer, send one ingest message,
 * wait for ack, close. Resolves (never rejects) so Outlook is never blocked.
 */
function sendToServer(action: string, itemKind: string, metadata: object): Promise<void> {
  return new Promise((resolve) => {
    const event = {
      id:            uuid(),
      timestamp:     new Date().toISOString(),
      platform:      "web",
      appBundleId:   null,
      correlationId: null,
      action,
      itemKind,
      metadata,
      widgetPath:     [],
      screenshotPath: null,
      clickPoint:     null,
      excelMetadata:  null,
    };

    const payload = JSON.stringify({ type: "ingest", event });

    const timer = setTimeout(resolve, SEND_TIMEOUT);
    const done  = () => { clearTimeout(timer); resolve(); };

    let ws: WebSocket;
    try {
      ws = new WebSocket(SERVER_WSS);
    } catch (_) {
      done();
      return;
    }

    ws.onopen    = () => ws.send(payload);
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

// ── Item metadata readers ─────────────────────────────────────────────────────

async function readMailMeta(item: any): Promise<object> {
  const [subject, toR, ccR, bccR] = await Promise.all([
    getAsync<string>((cb)                       => item.subject.getAsync(cb)),
    getAsync<Office.EmailAddressDetails[]>((cb) => item.to.getAsync(cb)),
    getAsync<Office.EmailAddressDetails[]>((cb) => item.cc.getAsync(cb)),
    item.bcc
      ? getAsync<Office.EmailAddressDetails[]>((cb) => item.bcc.getAsync(cb))
      : Promise.resolve([] as Office.EmailAddressDetails[]),
  ]);
  return {
    subject:     subject     ?? null,
    to:          (toR  ?? []).map((r) => r.emailAddress),
    cc:          (ccR  ?? []).map((r) => r.emailAddress),
    bcc:         (bccR ?? []).map((r) => r.emailAddress),
    attachments: (item.attachments || []).map((a: any) => a.name),
    body:        null,
  };
}

async function readApptMeta(item: any): Promise<object> {
  const [subject, required, optional] = await Promise.all([
    getAsync<string>((cb) => item.subject.getAsync(cb)),
    item.requiredAttendees
      ? getAsync<Office.EmailAddressDetails[]>((cb) => item.requiredAttendees.getAsync(cb))
      : Promise.resolve([] as Office.EmailAddressDetails[]),
    item.optionalAttendees
      ? getAsync<Office.EmailAddressDetails[]>((cb) => item.optionalAttendees.getAsync(cb))
      : Promise.resolve([] as Office.EmailAddressDetails[]),
  ]);
  return {
    subject:     subject   ?? null,
    to:          (required ?? []).map((r) => r.emailAddress),
    cc:          (optional ?? []).map((r) => r.emailAddress),
    bcc:         [],
    attachments: (item.attachments || []).map((a: any) => a.name),
    body:        null,
  };
}

// ── LaunchEvent handlers ──────────────────────────────────────────────────────

async function onNewMessageCompose(ev: LaunchEvent): Promise<void> {
  const meta = await readMailMeta(Office.context.mailbox.item);
  await sendToServer("compose", "mail", meta);
  ev.completed();
}

async function onNewAppointmentOrganizer(ev: LaunchEvent): Promise<void> {
  const meta = await readApptMeta(Office.context.mailbox.item);
  await sendToServer("compose", "appointment", meta);
  ev.completed();
}

async function onMessageSend(ev: LaunchEvent): Promise<void> {
  const meta = await readMailMeta(Office.context.mailbox.item);
  await sendToServer("send", "mail", meta);
  ev.completed({ allowEvent: true });
}

async function onAppointmentSend(ev: LaunchEvent): Promise<void> {
  const meta = await readApptMeta(Office.context.mailbox.item);
  await sendToServer("send", "appointment", meta);
  ev.completed({ allowEvent: true });
}

async function onMessageRecipientsChanged(ev: LaunchEvent): Promise<void> {
  const meta = await readMailMeta(Office.context.mailbox.item);
  await sendToServer("recipientsChange", "mail", meta);
  ev.completed();
}

async function onMessageAttachmentsChanged(ev: LaunchEvent): Promise<void> {
  const meta = await readMailMeta(Office.context.mailbox.item);
  await sendToServer("attachmentChange", "mail", meta);
  ev.completed();
}

async function onAppointmentAttendeesChanged(ev: LaunchEvent): Promise<void> {
  const meta = await readApptMeta(Office.context.mailbox.item);
  await sendToServer("recipientsChange", "appointment", meta);
  ev.completed();
}

async function onAppointmentTimeChanged(ev: LaunchEvent): Promise<void> {
  const meta = await readApptMeta(Office.context.mailbox.item);
  await sendToServer("timeChange", "appointment", meta);
  ev.completed();
}

async function onAppointmentRecurrenceChanged(ev: LaunchEvent): Promise<void> {
  const meta = await readApptMeta(Office.context.mailbox.item);
  await sendToServer("recurrenceChange", "appointment", meta);
  ev.completed();
}

async function onEnhancedLocationsChanged(ev: LaunchEvent): Promise<void> {
  const meta = await readApptMeta(Office.context.mailbox.item);
  await sendToServer("locationChange", "appointment", meta);
  ev.completed();
}

// ── Register all handlers with Office ────────────────────────────────────────

Office.onReady(() => {
  Office.actions.associate("onNewMessageCompose",           onNewMessageCompose);
  Office.actions.associate("onNewAppointmentOrganizer",     onNewAppointmentOrganizer);
  Office.actions.associate("onMessageSend",                 onMessageSend);
  Office.actions.associate("onAppointmentSend",             onAppointmentSend);
  Office.actions.associate("onMessageRecipientsChanged",    onMessageRecipientsChanged);
  Office.actions.associate("onMessageAttachmentsChanged",   onMessageAttachmentsChanged);
  Office.actions.associate("onAppointmentAttendeesChanged", onAppointmentAttendeesChanged);
  Office.actions.associate("onAppointmentTimeChanged",      onAppointmentTimeChanged);
  Office.actions.associate("onAppointmentRecurrenceChanged",onAppointmentRecurrenceChanged);
  Office.actions.associate("onEnhancedLocationsChanged",    onEnhancedLocationsChanged);
});
