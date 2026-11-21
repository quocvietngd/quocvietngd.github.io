#!/usr/bin/env bash
set -euo pipefail

# Auto-push restored local state when production backend is reachable.
# Usage:
#   bash scripts/push-backfill-when-live.sh [BACKEND_URL] [MAX_MINUTES]

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
STATE_FILE="${REPO_DIR}/server/telegram-bridge-state.json"
BACKEND_URL="${1:-https://nora-sync-quocvietngd-2026.onrender.com}"
MAX_MINUTES="${2:-30}"

if [[ ! -f "$STATE_FILE" ]]; then
  echo "[ERROR] Missing state file: $STATE_FILE"
  exit 1
fi

export BACKEND_URL STATE_FILE MAX_MINUTES

node - <<'NODE'
const fs = require('fs');

const base = String(process.env.BACKEND_URL || '').replace(/\/+$/, '');
const stateFile = String(process.env.STATE_FILE || '');
const maxMinutes = Math.max(1, Number(process.env.MAX_MINUTES || 30));

const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function dateOfRow(raw = {}) {
  return String(raw.registrationDate || raw.appointmentDate || raw.date || '').slice(0, 10);
}

function summarizeRows(reports = []) {
  const byDate = new Map();
  for (const item of reports) {
    const raw = item && typeof item.raw === 'object' ? item.raw : {};
    const d = dateOfRow(raw);
    if (!d) continue;
    byDate.set(d, (byDate.get(d) || 0) + 1);
  }
  return Object.fromEntries([...byDate.entries()].sort(([a], [b]) => a.localeCompare(b)));
}

async function fetchText(path, options = {}) {
  try {
    const res = await fetch(`${base}${path}`, options);
    const txt = await res.text();
    return { ok: res.ok, status: res.status, text: txt, headers: res.headers };
  } catch (error) {
    return { ok: false, status: 0, text: String(error && error.message ? error.message : error) };
  }
}

(async () => {
  if (!base) {
    console.error('[ERROR] Missing BACKEND_URL');
    process.exit(1);
  }

  const state = JSON.parse(fs.readFileSync(stateFile, 'utf8'));
  const secret = String(state.webhookSecret || '').trim();
  if (!secret) {
    console.error('[ERROR] webhookSecret is empty in local state file');
    process.exit(1);
  }

  const localReports = Array.isArray(state.reports) ? state.reports : [];
  const localSummary = summarizeRows(localReports);
  const localTotal = localReports.length;
  console.log('[INFO] Local reports:', localTotal);
  const interesting = Object.entries(localSummary).filter(([d]) => d >= '2026-05-29');
  console.log('[INFO] Local dates >= 2026-05-29:', interesting.length ? JSON.stringify(Object.fromEntries(interesting)) : '{}');

  const started = Date.now();
  const maxMs = maxMinutes * 60 * 1000;
  let lastStatus = -1;

  while (Date.now() - started < maxMs) {
    const health = await fetchText('/api/telegram/health');
    if (health.status !== lastStatus) {
      const preview = String(health.text || '').replace(/\s+/g, ' ').slice(0, 120);
      console.log(`[INFO] Health status=${health.status} body=${preview}`);
      lastStatus = health.status;
    }

    if (health.ok) {
      console.log('[INFO] Backend reachable. Starting import-full-state...');
      const payload = {
        token: state.token || '',
        chatId: state.chatId || '',
        lastUpdateId: Number(state.lastUpdateId || 0),
        webhookBaseUrl: state.webhookBaseUrl || '',
        webhookSecret: state.webhookSecret || '',
        webhookPath: state.webhookPath || '',
        updatedAt: Date.now(),
        reports: Array.isArray(state.reports) ? state.reports : [],
        users: Array.isArray(state.users) ? state.users : [],
        kpiReports: Array.isArray(state.kpiReports) ? state.kpiReports : [],
        appState: state.appState && typeof state.appState === 'object' ? state.appState : {}
      };

      const imported = await fetchText('/api/admin/import-full-state', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${secret}`
        },
        body: JSON.stringify(payload)
      });

      if (!imported.ok) {
        const preview = String(imported.text || '').replace(/\s+/g, ' ').slice(0, 180);
        console.error(`[ERROR] Import failed status=${imported.status} body=${preview}`);
        process.exit(2);
      }

      let importJson = {};
      try { importJson = JSON.parse(imported.text || '{}'); } catch {}
      console.log('[INFO] Import response:', JSON.stringify(importJson));

      const pending = await fetchText('/api/telegram/pending?all=1&sync=0');
      if (!pending.ok) {
        const preview = String(pending.text || '').replace(/\s+/g, ' ').slice(0, 180);
        console.error(`[ERROR] Verify pending failed status=${pending.status} body=${preview}`);
        process.exit(3);
      }

      let pendingJson = {};
      try { pendingJson = JSON.parse(pending.text || '{}'); } catch {}
      const rows = Array.isArray(pendingJson.rows) ? pendingJson.rows : [];
      const remoteSummary = {};
      for (const row of rows) {
        const d = dateOfRow(row);
        if (!d) continue;
        remoteSummary[d] = (remoteSummary[d] || 0) + 1;
      }
      const remoteInteresting = Object.fromEntries(Object.entries(remoteSummary).filter(([d]) => d >= '2026-05-29').sort(([a], [b]) => a.localeCompare(b)));

      console.log('[SUCCESS] Restored state to production backend.');
      console.log('[SUCCESS] Remote pendingCount:', Number(pendingJson.pendingCount || rows.length || 0));
      console.log('[SUCCESS] Remote dates >= 2026-05-29:', JSON.stringify(remoteInteresting));
      process.exit(0);
    }

    await wait(15000);
  }

  console.error(`[ERROR] Timeout after ${maxMinutes} minutes waiting backend availability.`);
  process.exit(4);
})();
NODE
