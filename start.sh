#!/bin/bash
# 文件检索管理系统 — 启动脚本
cd "$(dirname "$0")"
port="${1:-8765}"
echo "$port" > .port
python3 workdoc_web.py "$port"