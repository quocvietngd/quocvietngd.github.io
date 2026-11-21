#!/usr/bin/env bash
# restore-state.sh — Đẩy dữ liệu local lên Render server sau khi redeploy
# Cách dùng:
#   bash scripts/restore-state.sh [BACKEND_URL]
#
# Nếu không có BACKEND_URL, sẽ dùng endpoint mặc định từ public/runtime-config.json

set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
STATE_FILE="${REPO_DIR}/server/telegram-bridge-state.json"
CONFIG_FILE="${REPO_DIR}/public/runtime-config.json"

RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; NC='\033[0m'
info() { echo -e "${GREEN}[INFO]${NC} $*"; }
warn() { echo -e "${YELLOW}[WARN]${NC} $*"; }
err()  { echo -e "${RED}[ERROR]${NC} $*"; exit 1; }

# ─── 1. Xác định backend URL ──────────────────────────────────────────────
BACKEND_URL="${1:-}"
if [[ -z "$BACKEND_URL" ]]; then
  if [[ -f "$CONFIG_FILE" ]]; then
    USERS_ENDPOINT=$(python3 -c "import json; print(json.load(open('${CONFIG_FILE}')).get('usersSyncEndpoint',''))" 2>/dev/null || true)
    if [[ -n "$USERS_ENDPOINT" ]]; then
      BACKEND_URL=$(python3 -c "
url = '${USERS_ENDPOINT}'.strip().rstrip('/')
# Normalize to base URL
import re
m = re.match(r'(https?://[^/]+)', url)
print(m.group(1) if m else url)
" 2>/dev/null || true)
    fi
  fi
fi

if [[ -z "$BACKEND_URL" ]]; then
  BACKEND_URL="https://nora-sync-quocvietngd-2026-2.onrender.com"
fi

info "Backend URL: ${BACKEND_URL}"

# ─── 2. Kiểm tra local state file ────────────────────────────────────────
if [[ ! -f "$STATE_FILE" ]]; then
  err "Không tìm thấy state file: ${STATE_FILE}"
fi

STATE_SIZE=$(wc -c < "$STATE_FILE" | tr -d ' ')
info "Local state file: ${STATE_SIZE} bytes"

# ─── 3. Lấy webhook secret từ local state ────────────────────────────────
WEBHOOK_SECRET=$(python3 -c "
import json
with open('${STATE_FILE}') as f:
    d = json.load(f)
print(d.get('webhookSecret',''))
" 2>/dev/null || true)

if [[ -z "$WEBHOOK_SECRET" ]]; then
  # Lấy từ server đang chạy
  info "Đang lấy webhook secret từ server..."
  SERVER_SECRET=$(curl -sf "${BACKEND_URL}/api/storage/status" 2>/dev/null | python3 -c "
import sys, json
d = json.load(sys.stdin)
" 2>/dev/null || true)
  err "Không lấy được webhook secret. Hãy kiểm tra local state file."
fi

info "Webhook secret đã sẵn sàng."

# ─── 4. Chờ server sẵn sàng ─────────────────────────────────────────────
info "Đang kiểm tra server..."
MAX_RETRIES=20
for i in $(seq 1 $MAX_RETRIES); do
  STATUS=$(curl -sf "${BACKEND_URL}/api/storage/status" 2>/dev/null || true)
  if [[ -n "$STATUS" ]]; then
    MODE=$(echo "$STATUS" | python3 -c "import sys,json; print(json.load(sys.stdin).get('mode',''))" 2>/dev/null || true)
    info "Server đang chạy ở chế độ: ${MODE}"
    if [[ "$MODE" == "postgres" ]]; then
      info "✓ Server đã chuyển sang PostgreSQL mode!"
    elif [[ "$MODE" == "json-file" ]]; then
      warn "Server vẫn ở json-file mode. Kiểm tra lại DATABASE_URL trên Render."
      warn "Vẫn tiếp tục restore (dữ liệu sẽ ở trong json-file)..."
    fi
    break
  fi
  echo "  Thử lại ${i}/${MAX_RETRIES}..."
  sleep 5
done

if [[ -z "${STATUS:-}" ]]; then
  err "Server không phản hồi sau ${MAX_RETRIES} lần thử. Hãy chờ server deploy xong rồi thử lại."
fi

# ─── 5. Push state lên server ─────────────────────────────────────────────
info "Đang đẩy dữ liệu lên server..."
RESPONSE=$(curl -sf -X POST "${BACKEND_URL}/api/admin/import-full-state" \
  -H "Authorization: Bearer ${WEBHOOK_SECRET}" \
  -H "Content-Type: application/json" \
  --data-binary "@${STATE_FILE}" 2>/dev/null) || {
  err "Kết nối thất bại. Kiểm tra lại backend URL và webhook secret."
}

OK=$(echo "$RESPONSE" | python3 -c "import sys,json; print(json.load(sys.stdin).get('ok','false'))" 2>/dev/null || echo "false")

if [[ "$OK" == "True" || "$OK" == "true" ]]; then
  echo ""
  echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
  echo -e "${GREEN}✓ KHÔI PHỤC THÀNH CÔNG${NC}"
  echo ""
  echo "$RESPONSE" | python3 -c "
import sys, json
d = json.load(sys.stdin)
counts = d.get('counts', {})
print('  Dữ liệu đã đẩy lên:')
for k, v in counts.items():
    print(f'    • {k}: {v}')
"
  echo ""
  echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
else
  warn "Phản hồi từ server:"
  echo "$RESPONSE"
  err "Import thất bại. Kiểm tra lại."
fi
