#!/bin/bash
echo ""
echo "============================================="
echo "  吉林省高考志愿规划系统 · 本地版 v2.0"
echo "============================================="
echo ""
command -v python3 &>/dev/null || { echo "[错误] 未找到 Python3"; exit 1; }
echo "安装依赖..."
pip3 install flask openpyxl pandas numpy -q
(sleep 2 && (open "http://localhost:5000" 2>/dev/null || xdg-open "http://localhost:5000" 2>/dev/null)) &
echo "启动服务: http://localhost:5000"
echo "Ctrl+C 停止"
echo ""
PYTHONIOENCODING=utf-8 python3 app.py
