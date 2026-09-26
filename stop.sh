#!/bin/bash
# 文件检索管理系统 — 停止脚本
lsof -ti tcp:8765 | xargs -r kill 2>/dev/null
echo "已停止（若端口非8765，请改脚本中的端口）"