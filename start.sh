#!/bin/bash
# 文件检索管理系统 — 启动脚本
cd "$(dirname "$0")"
python3 workdoc_web.py "$@"