#!/usr/bin/env bash
set -euo pipefail

REPO_URL="${1:-}"
APP_DIR="${APP_DIR:-/opt/excel-splitter}"
APP_PORT="${APP_PORT:-8123}"

if [ -z "$REPO_URL" ]; then
  echo "用法: curl -fsSL <你的脚本地址> | sudo bash -s -- <GitHub仓库URL>" >&2
  echo "示例: sudo bash deploy/install-ubuntu.sh https://github.com/你的用户名/excel-splitter.git" >&2
  exit 1
fi

if [ "$(id -u)" -ne 0 ]; then
  echo "请用 sudo/root 运行部署脚本" >&2
  exit 1
fi

export DEBIAN_FRONTEND=noninteractive

apt-get update
apt-get install -y ca-certificates curl git ufw

install -m 0755 -d /etc/apt/keyrings
if [ ! -f /etc/apt/keyrings/docker.asc ]; then
  curl -fsSL https://download.docker.com/linux/ubuntu/gpg -o /etc/apt/keyrings/docker.asc
  chmod a+r /etc/apt/keyrings/docker.asc
fi

. /etc/os-release
ARCH="$(dpkg --print-architecture)"
CODENAME="${VERSION_CODENAME:-jammy}"
echo "deb [arch=${ARCH} signed-by=/etc/apt/keyrings/docker.asc] https://download.docker.com/linux/ubuntu ${CODENAME} stable" > /etc/apt/sources.list.d/docker.list

apt-get update
apt-get install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin

mkdir -p "$(dirname "$APP_DIR")"
if [ -d "$APP_DIR/.git" ]; then
  git -C "$APP_DIR" fetch --all --prune
  git -C "$APP_DIR" reset --hard origin/main || git -C "$APP_DIR" reset --hard origin/master
else
  rm -rf "$APP_DIR"
  git clone "$REPO_URL" "$APP_DIR"
fi

cd "$APP_DIR"

if [ "$APP_PORT" != "8123" ]; then
  sed -i "s/8123:8123/${APP_PORT}:8123/g" docker-compose.yml
fi

docker compose up -d --build

if command -v ufw >/dev/null 2>&1; then
  ufw allow OpenSSH >/dev/null || true
  ufw allow "${APP_PORT}/tcp" >/dev/null || true
fi

for _ in $(seq 1 30); do
  if curl -fsS "http://127.0.0.1:${APP_PORT}/health" >/dev/null; then
    echo "部署成功"
    echo "本机访问: http://127.0.0.1:${APP_PORT}/"
    PUBLIC_IP="$(curl -fsS https://api.ipify.org 2>/dev/null || true)"
    if [ -n "$PUBLIC_IP" ]; then
      echo "公网访问: http://${PUBLIC_IP}:${APP_PORT}/"
    fi
    exit 0
  fi
  sleep 2
done

echo "服务启动超时，查看日志：" >&2
docker compose logs --tail=100 >&2
exit 1
