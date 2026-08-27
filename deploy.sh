#!/bin/bash
set -e
cd "$(dirname "$0")"

# 支持两种位置：
#   /home/.../nnroad-office-convert/deploy.sh   （与 Office 一样，旁边是 repo/）
#   仓库根目录 ./deploy.sh                      （convert_api.py 在同级）
if [ -d repo ] && [ -f repo/convert_api.py ]; then
  ROOT="$(pwd)"
  REPO_DIR="$ROOT/repo"
elif [ -f convert_api.py ]; then
  REPO_DIR="$(pwd)"
  ROOT="$(cd .. && pwd)"
else
  echo "[错误] 找不到 convert_api.py，请把本脚本放在 nnroad-office-convert/ 或仓库根目录"
  exit 1
fi

cd "$REPO_DIR"
git fetch origin
git checkout main
git reset --hard origin/main

echo "===== LAST COMMIT ====="
git log -1 --pretty=format:"commit: %h%n信息: %s%n作者: %an%n时间: %ai%n"
echo "========================"

ENV_FILE="$ROOT/env/convert.env"
if [ ! -f "$ENV_FILE" ]; then
  echo "[错误] 找不到环境文件: $ENV_FILE"
  exit 1
fi
sed -i 's/\r$//' "$ENV_FILE"
set -a
. "$ENV_FILE"
set +a
echo "[环境] CONVERT_HOST=${CONVERT_HOST:-127.0.0.1} CONVERT_PORT=${CONVERT_PORT:-8765}"

VENV="$ROOT/venv"
if [ ! -x "$VENV/bin/python" ]; then
  echo "[错误] 找不到虚拟环境: $VENV （先 python3 -m venv venv）"
  exit 1
fi

echo "===== install dependencies ====="
"$VENV/bin/pip" install -q -r "$REPO_DIR/requirements.txt"

ps -ef | grep "[u]vicorn convert_api:app" | awk '{print $2}' | xargs -r kill -9
sleep 1

cd "$ROOT"
nohup "$VENV/bin/python" -m uvicorn convert_api:app \
  --app-dir "$REPO_DIR" \
  --host "${CONVERT_HOST:-127.0.0.1}" \
  --port "${CONVERT_PORT:-8765}" \
  > "$ROOT/convert.out" 2>&1 &

sleep 3
echo "===== health ====="
curl -sS "http://127.0.0.1:${CONVERT_PORT:-8765}/health" || true
echo
echo "===== convert.out ====="
tail -n 20 "$ROOT/convert.out"
echo "===== convert deploy finished ====="
