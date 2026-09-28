#!/bin/bash
# 文件检索管理系统 — 停止脚本
cd "$(dirname "$0")"
port="$(cat .port 2>/dev/null || echo 8765)"
lsof -ti tcp:"$port" | xargs -r kill 2>/dev/null
echo "已停止（端口 $port）"