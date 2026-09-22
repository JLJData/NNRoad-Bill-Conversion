#!/bin/bash
# 拷到服务器 /home/nnroad-office/nnroad-office-convert/deploy.sh 再执行。
# Dockerfile 也放在这一层（与 repo 并列）。脚本会重置 repo。
set -e
cd "$(dirname "$0")"

if [ -d repo ] && [ -f repo/convert_api.py ]; then
    ROOT="$(pwd)"
    REPO_DIR="$ROOT/repo"
elif [ -f convert_api.py ]; then
    REPO_DIR="$(pwd)"
    ROOT="$(cd .. && pwd)"
else
    echo "[错误] 找不到 convert_api.py"
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
DOCKERFILE="$ROOT/Dockerfile"
if [ ! -f "$DOCKERFILE" ]; then
    DOCKERFILE="$REPO_DIR/Dockerfile"
fi
if [ ! -f "$DOCKERFILE" ]; then
    echo "[错误] 找不到 Dockerfile（放在 $ROOT/Dockerfile）"
    exit 1
fi

sed -i 's/\r$//' "$ENV_FILE"
set -a
# shellcheck disable=SC1090
. "$ENV_FILE"
set +a
echo "[环境] CONVERT_HOST=${CONVERT_HOST:-127.0.0.1} CONVERT_PORT=${CONVERT_PORT:-8765}"

DOCKER_ENV=$(mktemp)
chmod 600 "$DOCKER_ENV"
trap 'rm -f "$DOCKER_ENV"' EXIT
while IFS= read -r line || [ -n "$line" ]; do
    [[ "$line" =~ ^[[:space:]]*# ]] && continue
    [[ -z "${line//[[:space:]]/}" ]] && continue
    line="${line#"${line%%[![:space:]]*}"}"
    line="${line#export }"
    key="${line%%=*}"
    [[ "$key" =~ ^[A-Za-z_][A-Za-z0-9_]*$ ]] || continue
    printf '%s=%s\n' "$key" "${!key}" >> "$DOCKER_ENV"
done < "$ENV_FILE"

systemctl enable docker
if docker image inspect nnroad-office-convert:latest >/dev/null 2>&1; then
    docker tag nnroad-office-convert:latest "nnroad-office-convert:$(date +%Y%m%d%H%M)"
fi
docker build -f "$DOCKERFILE" -t nnroad-office-convert:latest "$REPO_DIR"

ps -ef | grep "[u]vicorn convert_api:app" | awk '{print $2}' | xargs -r kill -9 || true
sleep 1
docker rm -f nnroad-office-convert >/dev/null 2>&1 || true

docker run -d --name nnroad-office-convert \
    --network host \
    --restart unless-stopped \
    --env-file "$DOCKER_ENV" \
    nnroad-office-convert:latest

sleep 3
echo "===== health ====="
curl -sS "http://127.0.0.1:${CONVERT_PORT:-8765}/health" || true
echo
docker logs --tail 40 nnroad-office-convert
echo "===== convert deploy finished ====="
