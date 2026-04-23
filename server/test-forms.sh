#!/bin/bash
# Test tất cả 4 form template sau khi Render deploy mới

BASE="${SERVICE_URL:-https://nora-sync-quocvietngd-2026-2.onrender.com}"
SECRET=""

echo "=== Kiểm tra Render build version ==="
curl -s "$BASE/api/telegram/health" | python3 -c "import json,sys;d=json.load(sys.stdin);print('buildTs:', d.get('buildTs','OLD - chưa deploy!'), '| configured:', d.get('configured','?'))"
echo ""

echo "=== Cấu hình lại Telegram Webhook ==="
CONFIG_JSON=$(curl -s -X POST "$BASE/api/telegram/config" \
  -H 'Content-Type: application/json' \
  -d '{
    "token": "8609837486:AAHU03XBr83izxnOl8wAoJ3iC7UgWOoaN4A",
    "chatId": "-5296647339",
    "webhookBaseUrl": "'"$BASE"'"
  }')
echo "$CONFIG_JSON" | python3 -c "import json,sys;d=json.load(sys.stdin);print('configured:', d.get('configured'), '| webhookUrl:', d.get('webhookUrl',''))"
SECRET=$(echo "$CONFIG_JSON" | python3 -c "import json,sys;d=json.load(sys.stdin);u=d.get('webhookUrl','');print(u.rsplit('/',1)[-1] if '/' in u else '')")
if [ -z "$SECRET" ]; then
  echo "❌ Không lấy được webhook secret từ cấu hình"
  exit 1
fi
echo ""

BEFORE=$(curl -s "$BASE/api/telegram/pending?all=1" | python3 -c "import json,sys;d=json.load(sys.stdin);print(d.get('pendingCount',0))")
echo "=== Records trước khi test: $BEFORE ==="
echo ""

send_test() {
  local label="$1" text="$2" uid="$3"
  result=$(curl -s -X POST "$BASE/api/telegram/webhook/$SECRET" \
    -H 'Content-Type: application/json' \
    -d "{\"update_id\":$uid,\"message\":{\"message_id\":$uid,\"date\":1714000000,\"chat\":{\"id\":-5296647339},\"text\":\"$text\"}}")
  accepted=$(echo "$result" | python3 -c "import json,sys;d=json.load(sys.stdin);print('✅ ACCEPTED' if d.get('accepted') else '❌ IGNORED')" 2>/dev/null)
  echo "[$label] $accepted"
}

echo "=== Test 4 form templates ==="
BASE_ID=$(date +%s)
send_test "#dieuduong" "#dieuduong\nTên DD: Hoa Nguyễn\nDịch vụ: Chăm sóc bầu\nTên KH: Chị Lan\nMã HĐ: HD-001\nSố buổi: 5\nKhoảng cách: 3km" $((BASE_ID + 1))
send_test "#marketing" "#marketing\nTên NV: Anh Phúc\nChi phí: 5000000\nMess: 120\nSĐT: 45\nLịch: 8\nHợp đồng: 3\nDoanh số: 50000000" $((BASE_ID + 2))
send_test "#tuvan" "#tuvan\nTên TV: Cô Mai\nTên KH: Anh Dũng\nKết quả: Đồng ý\nMã HĐ: HD-205\nSố tiền: 20000000\nPTTT: Chuyển khoản" $((BASE_ID + 3))
send_test "#telesale" "#telesale\nTên NV: Bảo\nMess: 80\nSĐT: 30\nLịch: 5\nCa hoãn/huỷ: 2\nHợp đồng: 2\nDoanh số: 30000000" $((BASE_ID + 4))
echo ""

AFTER=$(curl -s "$BASE/api/telegram/pending?all=1" | python3 -c "import json,sys;d=json.load(sys.stdin);print(d.get('pendingCount',0))")
echo "=== Records sau khi test: $AFTER (tăng thêm $((AFTER - BEFORE))) ==="
