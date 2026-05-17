#!/usr/bin/env bash
# setup-render.sh — Tự động cấu hình Render backend để dùng storage bền vững
# Cách dùng:
#   bash scripts/setup-render.sh <RENDER_API_KEY> [GITHUB_TOKEN hoặc DATABASE_URL]
#
# OPTION 1 (ĐƠN GIẢN NHẤT — 30 giây):
#   - GITHUB_TOKEN: GitHub Personal Access Token với scope "gist"
#   - Tạo tại: https://github.com/settings/tokens/new?scopes=gist&description=nora-bridge
#   - Ví dụ: bash scripts/setup-render.sh rnd_xxxx ghp_yyyy
#
# OPTION 2 (PostgreSQL):
#   - DATABASE_URL: chuỗi kết nối PostgreSQL (Neon free tier)
#   - Tạo tại: https://neon.tech
#   - Ví dụ: bash scripts/setup-render.sh rnd_xxxx postgresql://...

set -euo pipefail

RENDER_API_KEY="${1:-}"
SECOND_PARAM="${2:-}"
SERVICE_NAME="nora-sync-quocvietngd-2026"
API_BASE="https://api.render.com/v1"

RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; NC='\033[0m'

info()    { echo -e "${GREEN}[INFO]${NC} $*"; }
warn()    { echo -e "${YELLOW}[WARN]${NC} $*"; }
err()     { echo -e "${RED}[ERROR]${NC} $*"; exit 1; }

# Determine storage type
GITHUB_TOKEN=""
DATABASE_URL=""
if [[ "$SECOND_PARAM" == ghp_* || "$SECOND_PARAM" == github_pat_* || "$SECOND_PARAM" == gho_* ]]; then
  GITHUB_TOKEN="$SECOND_PARAM"
  info "Sẽ cấu hình GitHub Gist storage"
elif [[ "$SECOND_PARAM" == postgresql://* || "$SECOND_PARAM" == postgres://* ]]; then
  DATABASE_URL="$SECOND_PARAM"
  info "Sẽ cấu hình PostgreSQL storage"
fi

# ─── 1. Kiểm tra API key ───────────────────────────────────────────────────
if [[ -z "$RENDER_API_KEY" ]]; then
  echo ""
  echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
  echo "  NORA Dashboard — Cấu hình bền vững cho Render Backend"
  echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
  echo ""
  warn "Bạn chưa cung cấp Render API Key."
  echo ""
  echo "  Bước 1: Lấy Render API Key (30 giây)"
  echo "  → Mở: https://dashboard.render.com/u/settings → 'API Keys' → 'Create API Key'"
  echo ""
  echo "  Bước 2: Lấy GitHub Token cho Gist storage (30 giây, KHÔNG cần credit card)"
  echo "  → Mở: https://github.com/settings/tokens/new?scopes=gist&description=nora-bridge"
  echo "  → Nhấn 'Generate token' → Copy token (ghp_...)"
  echo ""
  echo "  Bước 3: Chạy:"
  echo "     bash scripts/setup-render.sh <RENDER_API_KEY> <GITHUB_TOKEN>"
  echo ""
  open "https://dashboard.render.com/u/settings" 2>/dev/null || true
  exit 1
fi

# ─── 2. Tìm service ID ────────────────────────────────────────────────────
info "Đang tìm service '${SERVICE_NAME}' trên Render..."
SERVICES_JSON=$(curl -sf "${API_BASE}/services?limit=100" \
  -H "Authorization: Bearer ${RENDER_API_KEY}" \
  -H "Accept: application/json") || err "Không kết nối được Render API. Kiểm tra API key."

SERVICE_ID=$(echo "$SERVICES_JSON" | python3 -c "
import sys, json
data = json.load(sys.stdin)
items = data if isinstance(data, list) else data.get('services', [])
for item in items:
    svc = item.get('service', item)
    if svc.get('name','') == '${SERVICE_NAME}':
        print(svc['id'])
        break
" 2>/dev/null || true)

if [[ -z "$SERVICE_ID" ]]; then
  warn "Không tìm thấy service '${SERVICE_NAME}'."
  echo "$SERVICES_JSON" | python3 -c "
import sys, json
data = json.load(sys.stdin)
items = data if isinstance(data, list) else data.get('services', [])
for item in items:
    svc = item.get('service', item)
    print(f\"  - {svc.get('name','')} ({svc.get('id','')})\")" 2>/dev/null
  err "Không tìm thấy service."
fi

info "Tìm thấy service: ${SERVICE_ID}"

# ─── 3. Prompt nếu chưa có storage config ────────────────────────────────────
if [[ -z "$GITHUB_TOKEN" && -z "$DATABASE_URL" ]]; then
  echo ""
  echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
  warn "Chưa có storage config. Chọn 1 trong 2:"
  echo ""
  echo "  OPTION 1 (ĐƠN GIẢN — GitHub Gist, miễn phí):"
  echo "  → Tạo token: https://github.com/settings/tokens/new?scopes=gist&description=nora-bridge"
  echo "  → Chạy: bash scripts/setup-render.sh ${RENDER_API_KEY} <ghp_...>"
  echo ""
  echo "  OPTION 2 (PostgreSQL — Neon free tier):"
  echo "  → Tạo DB: https://console.neon.tech/signup"
  echo "  → Chạy: bash scripts/setup-render.sh ${RENDER_API_KEY} postgresql://..."
  echo ""
  open "https://github.com/settings/tokens/new?scopes=gist&description=nora-bridge" 2>/dev/null || true
  exit 1
fi

# ─── 4. Cập nhật env vars trên Render ─────────────────────────────────────
info "Đang cập nhật env vars cho service ${SERVICE_ID}..."

ENV_CURRENT=$(curl -sf "${API_BASE}/services/${SERVICE_ID}/env-vars" \
  -H "Authorization: Bearer ${RENDER_API_KEY}" \
  -H "Accept: application/json" 2>/dev/null || echo "[]")

NEW_ENVS=$(echo "$ENV_CURRENT" | python3 -c "
import sys, json
existing = json.load(sys.stdin)
if isinstance(existing, list):
    items = [e.get('envVar', e) for e in existing]
else:
    items = existing.get('envVars', [])

envmap = {e['key']: e['value'] for e in items if 'key' in e and 'value' in e}

github_token = '''${GITHUB_TOKEN}'''
database_url = '''${DATABASE_URL}'''

if github_token:
    envmap['GITHUB_TOKEN'] = github_token
    envmap.pop('DATABASE_URL', None)
    envmap.pop('PERSISTENCE_MODE', None)
elif database_url:
    envmap['DATABASE_URL'] = database_url
    envmap['PERSISTENCE_MODE'] = 'postgres-only'
    envmap.pop('GITHUB_TOKEN', None)

result = [{'key': k, 'value': v} for k, v in envmap.items()]
print(json.dumps(result))
" 2>/dev/null)

UPDATE_RESULT=$(curl -sf -X PUT "${API_BASE}/services/${SERVICE_ID}/env-vars" \
  -H "Authorization: Bearer ${RENDER_API_KEY}" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json" \
  -d "$NEW_ENVS") || err "Cập nhật env vars thất bại."

info "Env vars đã được cập nhật."

# ─── 5. Trigger redeploy ──────────────────────────────────────────────────
info "Đang trigger redeploy..."
DEPLOY_RESULT=$(curl -sf -X POST "${API_BASE}/services/${SERVICE_ID}/deploys" \
  -H "Authorization: Bearer ${RENDER_API_KEY}" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json" \
  -d '{"clearCache":"do_not_clear"}') || warn "Trigger redeploy thất bại."

DEPLOY_ID=$(echo "$DEPLOY_RESULT" | python3 -c "
import sys, json
d = json.load(sys.stdin)
dep = d.get('deploy', d)
print(dep.get('id', 'unknown'))
" 2>/dev/null || echo "unknown")

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo -e "${GREEN}✓ HOÀN TẤT${NC}"
echo ""
if [[ -n "$GITHUB_TOKEN" ]]; then
  echo "  Storage mode: GitHub Gist (durable)"
else
  echo "  Storage mode: PostgreSQL (durable)"
fi
echo "  Deploy ID: ${DEPLOY_ID}"
echo ""
warn "Sau khi deploy xong (2–3 phút), mở dashboard và dữ liệu sẽ tự sync từ browser."
echo "  Hoặc chạy: bash scripts/restore-state.sh"
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# setup-render.sh — Tự động cấu hình Render backend để dùng PostgreSQL bền vững
# Cách dùng:
#   bash scripts/setup-render.sh <RENDER_API_KEY> [DATABASE_URL]
#
# - RENDER_API_KEY: lấy tại https://dashboard.render.com/u/settings → API Keys
# - DATABASE_URL (tuỳ chọn): chuỗi kết nối PostgreSQL (Neon, Supabase, v.v.)
#   Nếu không cung cấp, script sẽ hỏi hoặc hướng dẫn tạo Neon miễn phí.

set -euo pipefail

RENDER_API_KEY="${1:-}"
DATABASE_URL="${2:-}"
SERVICE_NAME="nora-sync-quocvietngd-2026"
API_BASE="https://api.render.com/v1"

RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; NC='\033[0m'

info()    { echo -e "${GREEN}[INFO]${NC} $*"; }
warn()    { echo -e "${YELLOW}[WARN]${NC} $*"; }
err()     { echo -e "${RED}[ERROR]${NC} $*"; exit 1; }

# ─── 1. Kiểm tra API key ───────────────────────────────────────────────────
if [[ -z "$RENDER_API_KEY" ]]; then
  echo ""
  echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
  echo "  NORA Dashboard — Cấu hình bền vững cho Render Backend"
  echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
  echo ""
  warn "Bạn chưa cung cấp Render API Key."
  echo ""
  echo "  1. Mở: https://dashboard.render.com/u/settings"
  echo "  2. Cuộn xuống phần 'API Keys' → nhấn 'Create API Key'"
  echo "  3. Copy key và chạy lại:"
  echo ""
  echo "     bash scripts/setup-render.sh <YOUR_RENDER_API_KEY>"
  echo ""
  open "https://dashboard.render.com/u/settings" 2>/dev/null || true
  exit 1
fi

# ─── 2. Tìm service ID ────────────────────────────────────────────────────
info "Đang tìm service '${SERVICE_NAME}' trên Render..."
SERVICES_JSON=$(curl -sf "${API_BASE}/services?limit=100" \
  -H "Authorization: Bearer ${RENDER_API_KEY}" \
  -H "Accept: application/json") || err "Không kết nối được Render API. Kiểm tra API key."

SERVICE_ID=$(echo "$SERVICES_JSON" | python3 -c "
import sys, json
data = json.load(sys.stdin)
items = data if isinstance(data, list) else data.get('services', [])
for item in items:
    svc = item.get('service', item)
    if svc.get('name','') == '${SERVICE_NAME}':
        print(svc['id'])
        break
" 2>/dev/null || true)

if [[ -z "$SERVICE_ID" ]]; then
  warn "Không tìm thấy service '${SERVICE_NAME}'."
  echo "Các service hiện có:"
  echo "$SERVICES_JSON" | python3 -c "
import sys, json
data = json.load(sys.stdin)
items = data if isinstance(data, list) else data.get('services', [])
for item in items:
    svc = item.get('service', item)
    print(f\"  - {svc.get('name','')} ({svc.get('id','')})\")
"
  err "Không tìm thấy service. Kiểm tra lại tên service trong Render."
fi

info "Tìm thấy service: ${SERVICE_ID}"

# ─── 3. Lấy DATABASE_URL ──────────────────────────────────────────────────
if [[ -z "$DATABASE_URL" ]]; then
  # Kiểm tra xem đã có postgres nào trong account chưa
  info "Đang kiểm tra danh sách PostgreSQL trong account..."
  PG_JSON=$(curl -sf "${API_BASE}/postgres?limit=50" \
    -H "Authorization: Bearer ${RENDER_API_KEY}" \
    -H "Accept: application/json" 2>/dev/null || echo "[]")

  PG_COUNT=$(echo "$PG_JSON" | python3 -c "
import sys, json
data = json.load(sys.stdin)
items = data if isinstance(data, list) else data.get('postgresInstances', [])
print(len(items))
" 2>/dev/null || echo "0")

  if [[ "$PG_COUNT" -gt "0" ]]; then
    info "Tìm thấy ${PG_COUNT} PostgreSQL database. Đang lấy connection string..."
    DB_INFO=$(echo "$PG_JSON" | python3 -c "
import sys, json
data = json.load(sys.stdin)
items = data if isinstance(data, list) else data.get('postgresInstances', [])
if items:
    pg = items[0].get('postgres', items[0])
    print(pg.get('name','?'), pg.get('id','?'))
" 2>/dev/null || true)
    info "Database: ${DB_INFO}"

    PG_ID=$(echo "$PG_JSON" | python3 -c "
import sys, json
data = json.load(sys.stdin)
items = data if isinstance(data, list) else data.get('postgresInstances', [])
if items:
    pg = items[0].get('postgres', items[0])
    print(pg.get('id',''))
" 2>/dev/null || true)

    if [[ -n "$PG_ID" ]]; then
      DB_CONN=$(curl -sf "${API_BASE}/postgres/${PG_ID}/connection-string" \
        -H "Authorization: Bearer ${RENDER_API_KEY}" \
        -H "Accept: application/json" 2>/dev/null || true)
      DATABASE_URL=$(echo "$DB_CONN" | python3 -c "
import sys, json
d = json.load(sys.stdin)
print(d.get('connectionString', d.get('internalConnectionString', d.get('externalConnectionString', ''))))
" 2>/dev/null || true)
    fi
  fi

  if [[ -z "$DATABASE_URL" ]]; then
    echo ""
    echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    warn "Không tìm thấy PostgreSQL. Cần tạo database miễn phí:"
    echo ""
    echo "  CÁCH NHANH NHẤT (Neon PostgreSQL — miễn phí mãi mãi):"
    echo ""
    echo "  1. Mở: https://console.neon.tech/signup"
    echo "  2. Đăng ký (có thể dùng Google/GitHub)"
    echo "  3. Tạo project mới → copy 'Connection string'"
    echo "     (dạng: postgresql://user:pass@ep-xxx.neon.tech/neondb)"
    echo "  4. Chạy lại:"
    echo ""
    echo "     bash scripts/setup-render.sh ${RENDER_API_KEY} 'postgresql://...'"
    echo ""
    open "https://console.neon.tech/signup" 2>/dev/null || true
    exit 1
  fi
fi

info "DATABASE_URL đã sẵn sàng."

# ─── 4. Cập nhật env vars trên Render ─────────────────────────────────────
info "Đang cập nhật env vars cho service ${SERVICE_ID}..."

# Lấy env vars hiện tại
ENV_CURRENT=$(curl -sf "${API_BASE}/services/${SERVICE_ID}/env-vars" \
  -H "Authorization: Bearer ${RENDER_API_KEY}" \
  -H "Accept: application/json" 2>/dev/null || echo "[]")

# Build danh sách env vars mới — giữ lại tất cả và update/thêm các key cần thiết
NEW_ENVS=$(echo "$ENV_CURRENT" | python3 -c "
import sys, json
existing = json.load(sys.stdin)
if isinstance(existing, list):
    items = [e.get('envVar', e) for e in existing]
else:
    items = existing.get('envVars', [])

envmap = {e['key']: e['value'] for e in items if 'key' in e and 'value' in e}
envmap['DATABASE_URL'] = '''${DATABASE_URL}'''
envmap['PERSISTENCE_MODE'] = 'postgres-only'

result = [{'key': k, 'value': v} for k, v in envmap.items()]
print(json.dumps(result))
" 2>/dev/null)

UPDATE_RESULT=$(curl -sf -X PUT "${API_BASE}/services/${SERVICE_ID}/env-vars" \
  -H "Authorization: Bearer ${RENDER_API_KEY}" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json" \
  -d "$NEW_ENVS") || err "Cập nhật env vars thất bại."

info "Env vars đã được cập nhật."

# ─── 5. Trigger redeploy ──────────────────────────────────────────────────
info "Đang trigger redeploy..."
DEPLOY_RESULT=$(curl -sf -X POST "${API_BASE}/services/${SERVICE_ID}/deploys" \
  -H "Authorization: Bearer ${RENDER_API_KEY}" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json" \
  -d '{"clearCache":"do_not_clear"}') || warn "Trigger redeploy thất bại, hãy deploy thủ công."

DEPLOY_ID=$(echo "$DEPLOY_RESULT" | python3 -c "
import sys, json
d = json.load(sys.stdin)
dep = d.get('deploy', d)
print(dep.get('id', 'unknown'))
" 2>/dev/null || echo "unknown")

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo -e "${GREEN}✓ HOÀN TẤT${NC}"
echo ""
echo "  Deploy ID: ${DEPLOY_ID}"
echo "  Theo dõi: https://dashboard.render.com/web/${SERVICE_ID}/deploys/${DEPLOY_ID}"
echo ""
warn "Sau khi deploy xong (khoảng 2–3 phút), chạy lệnh sau để khôi phục dữ liệu:"
echo ""
echo "     bash scripts/restore-state.sh"
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
