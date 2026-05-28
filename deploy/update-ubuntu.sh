#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-/opt/excel-splitter}"

if [ "$(id -u)" -ne 0 ]; then
  echo "请用 sudo/root 运行更新脚本" >&2
  exit 1
fi

cd "$APP_DIR"
git pull --ff-only
docker compose up -d --build
curl -fsS http://127.0.0.1:8123/health >/dev/null
echo "更新完成: http://127.0.0.1:8123/"
