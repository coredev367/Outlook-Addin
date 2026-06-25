#!/usr/bin/env node
/**
 * bridge.js — Secure WebSocket Event Proxy for SkanOulook.
 *
 * Serves wss://localhost:8766 (TLS) so the add-in — loaded from
 * https://localhost:3000 — can connect without mixed-content errors.
 * Uses the same dev-certs that the webpack dev-server uses.
 *
 *   Outlook add-in ──wss://localhost:8766──► bridge.js ──ws://localhost:8765──► WesocketServer
 *                                               │
 *                                        logs to terminal
 *
 * Usage:
 *   node bridge.js
 *   node bridge.js 8766 ws://localhost:9000   (custom listen port + upstream URL)
 */

"use strict";

const https          = require("https");
const fs             = require("fs");
const path           = require("path");
const { WebSocketServer } = require("ws");

const LISTEN_PORT  = parseInt(process.argv[2], 10) || 8766;
const UPSTREAM_URL = process.argv[3] || "ws://localhost:8765";
const RECONNECT_MS = 3000;

// Dev-cert paths — same files the webpack dev-server uses
const CERT_DIR  = path.join(process.env.HOME, ".office-addin-dev-certs");
const TLS_CERT  = path.join(CERT_DIR, "localhost.crt");
const TLS_KEY   = path.join(CERT_DIR, "localhost.key");
const TLS_CA    = path.join(CERT_DIR, "ca.crt");   // may not exist; optional

// ── ANSI colours ──────────────────────────────────────────────────────────────

const C = {
  reset:   "\x1b[0m",
  bold:    "\x1b[1m",
  red:     "\x1b[31m",
  green:   "\x1b[32m",
  yellow:  "\x1b[33m",
  blue:    "\x1b[34m",
  magenta: "\x1b[35m",
  cyan:    "\x1b[36m",
  white:   "\x1b[37m",
  gray:    "\x1b[90m",
};

const ACTION_COLOR = {
  select:           C.blue,
  reply:            C.cyan,
  replyAll:         C.cyan,
  forward:          C.cyan,
  send:             C.green,
  save:             C.green,
  compose:          C.magenta,
  close:            C.gray,
  recipientsChange: C.yellow,
  attachmentChange: C.yellow,
  timeChange:       C.yellow,
  recurrenceChange: C.yellow,
  locationChange:   C.yellow,
  fromChange:       C.yellow,
  meetingAccept:    C.green,
  meetingTentative: C.yellow,
  meetingDecline:   C.red,
};

// ── Helpers ───────────────────────────────────────────────────────────────────

function ts() {
  return new Date().toLocaleTimeString("en-US", { hour12: false });
}

function hr() {
  return C.gray + "─".repeat(68) + C.reset;
}

function log(tag, color, message) {
  console.log(C.gray + ts() + C.reset + "  " + color + "[" + tag + "]" + C.reset + " " + message);
}

function printEvent(event) {
  try {
    const color = ACTION_COLOR[event.action] || C.white;
    const meta  = event.metadata || {};

    const subject = meta.subject      || "(no subject)";
    const toStr   = (meta.to          || []).join(", ") || "–";
    const ccStr   = (meta.cc          || []).join(", ");
    const bccStr  = (meta.bcc         || []).join(", ");
    const attStr  = (meta.attachments || []).join(", ");

    console.log("\n" + hr());
    console.log(
      C.gray + ts() + C.reset + "  " +
      color + C.bold + "[" + (event.action || "?").toUpperCase().padEnd(18) + "]" + C.reset + "  " +
      C.bold + subject + C.reset
    );
    console.log("  " + C.gray + "id       " + C.reset + (event.id       || "–"));
    console.log("  " + C.gray + "kind     " + C.reset + (event.itemKind || "–"));
    console.log("  " + C.gray + "platform " + C.reset + (event.platform || "–"));
    console.log("  " + C.gray + "to       " + C.reset + toStr);
    if (ccStr)  console.log("  " + C.gray + "cc       " + C.reset + ccStr);
    if (bccStr) console.log("  " + C.gray + "bcc      " + C.reset + bccStr);
    if (attStr) console.log("  " + C.gray + "attach   " + C.reset + attStr);
    if (meta.body) {
      var preview = String(meta.body).replace(/\s+/g, " ").slice(0, 120);
      console.log("  " + C.gray + "body     " + C.reset + preview + (meta.body.length > 120 ? "…" : ""));
    }
    console.log("  " + C.gray + "at       " + C.reset + (event.timestamp || "–"));
    console.log("");
  } catch (_) {}
}

// ── Upstream: connection to WesocketServer ────────────────────────────────────

var _upstream  = null;
var _upQueue   = [];
var _upClients = new Set();

function connectUpstream() {
  log("UP", C.gray, "Connecting to WesocketServer at " + UPSTREAM_URL + "…");

  var ws = new WebSocket(UPSTREAM_URL);   // Node.js v21+ built-in client
  _upstream = ws;

  ws.addEventListener("open", function () {
    log("UP", C.green, "Connected to WesocketServer");
    _upQueue.forEach(function (msg) { ws.send(msg); });
    _upQueue = [];
  });

  ws.addEventListener("message", function (ev) {
    try {
      var msg = JSON.parse(ev.data);
      if (msg.type === "ack") {
        log("UP", C.cyan, "WesocketServer ack  id=" + (msg.event && msg.event.id || "?"));
        var ackJson = ev.data;
        _upClients.forEach(function (client) {
          try {
            if (client.readyState === 1) { client.send(ackJson); }
          } catch (_) {}
        });
      }
    } catch (_) {}
  });

  ws.addEventListener("error", function () {
    log("UP", C.red, "WesocketServer connection error");
  });

  ws.addEventListener("close", function () {
    log("UP", C.yellow, "Disconnected from WesocketServer – retrying in " + RECONNECT_MS + "ms…");
    _upstream = null;
    setTimeout(connectUpstream, RECONNECT_MS);
  });
}

function forwardUpstream(payload) {
  var raw = typeof payload === "string" ? payload : JSON.stringify(payload);
  if (_upstream && _upstream.readyState === 1) {
    _upstream.send(raw);
  } else {
    _upQueue.push(raw);
    log("UP", C.yellow, "Queued – WesocketServer not connected yet");
  }
}

// ── TLS setup ─────────────────────────────────────────────────────────────────

function loadTls() {
  if (!fs.existsSync(TLS_CERT) || !fs.existsSync(TLS_KEY)) {
    console.error(
      C.red + "[ERR] Dev certs not found at " + CERT_DIR + C.reset + "\n" +
      "      Run:  npx office-addin-dev-certs install\n" +
      "      or:   npm run start   (which installs them automatically)"
    );
    process.exit(1);
  }
  var opts = {
    cert: fs.readFileSync(TLS_CERT),
    key:  fs.readFileSync(TLS_KEY),
  };
  if (fs.existsSync(TLS_CA)) {
    opts.ca = fs.readFileSync(TLS_CA);
  }
  return opts;
}

// ── Downstream: secure WebSocket server for Outlook add-in ───────────────────

var _clientCount = 0;

function startServer(tlsOpts) {
  var httpsServer = https.createServer(tlsOpts, function (req, res) {
    // CORS for fetch() calls from https://localhost:3000 (LaunchEvent handlers)
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");

    // Pre-flight
    if (req.method === "OPTIONS") {
      res.writeHead(204);
      res.end();
      return;
    }

    // POST /event — receives events from LaunchEvent handlers via fetch()
    if (req.method === "POST" && req.url === "/event") {
      var body = "";
      req.on("data", function (chunk) { body += chunk.toString(); });
      req.on("end", function () {
        try {
          var msg = JSON.parse(body);
          if (msg.type === "ingest" && msg.event) {
            printEvent(msg.event);
            forwardUpstream(msg);
            log("IN", C.blue, "HTTP event received  action=" + (msg.event.action || "?"));
          }
          res.writeHead(200, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ type: "ack", event: { id: msg && msg.event && msg.event.id } }));
        } catch (err) {
          log("IN", C.red, "Bad POST /event body: " + err.message);
          res.writeHead(400);
          res.end("Bad request");
        }
      });
      return;
    }

    // GET / — health-check
    res.writeHead(200, { "Content-Type": "text/plain" });
    res.end("SkanOulook bridge running  ws-clients=" + _clientCount + "\n");
  });

  var wss = new WebSocketServer({ server: httpsServer });

  wss.on("connection", function (clientWs) {
    _clientCount++;
    _upClients.add(clientWs);
    log("IN", C.cyan, "Add-in connected  (" + _clientCount + " active)");

    clientWs.on("message", function (data) {
      var msg;
      try {
        msg = JSON.parse(data.toString());
      } catch (_) {
        log("IN", C.red, "Could not parse message from add-in");
        return;
      }

      if (msg.type === "ingest" && msg.event) {
        printEvent(msg.event);
        forwardUpstream(msg);

        // Immediate local ack back to the add-in
        try {
          if (clientWs.readyState === 1) {
            clientWs.send(JSON.stringify({ type: "ack", event: { id: msg.event.id } }));
          }
        } catch (sendErr) {
          log("IN", C.red, "Failed to send ack: " + sendErr.message);
        }
      }
    });

    clientWs.on("close", function () {
      _clientCount--;
      _upClients.delete(clientWs);
      log("IN", C.yellow, "Add-in disconnected  (" + _clientCount + " active)");
    });

    clientWs.on("error", function (err) {
      _upClients.delete(clientWs);
      log("IN", C.red, "Add-in connection error: " + err.message);
    });
  });

  wss.on("error", function (err) {
    console.error(C.red + "[ERR] wss server error: " + err.message + C.reset);
    process.exit(1);
  });

  httpsServer.listen(LISTEN_PORT, function () {
    log("IN", C.green, "Listening (wss) on wss://localhost:" + LISTEN_PORT);
    console.log("");
  });
}

// ── Banner + start ────────────────────────────────────────────────────────────

console.log("");
console.log(C.bold + "  SkanOulook Bridge — Event Proxy  (wss)" + C.reset);
console.log(C.gray + "  Listening : " + C.reset + "wss://localhost:" + LISTEN_PORT +
            C.gray + "   ← Outlook add-in connects here" + C.reset);
console.log(C.gray + "  Upstream  : " + C.reset + UPSTREAM_URL +
            C.gray + "            ← WesocketServer" + C.reset);
console.log(C.gray + "  Certs     : " + C.reset + CERT_DIR);
console.log(C.gray + "  Stop      : Ctrl+C" + C.reset);
console.log("");

var tlsOpts = loadTls();
startServer(tlsOpts);
connectUpstream();
