#!/bin/bash
pkill -f "gunicorn.*app:app" 2>/dev/null
echo "✅ 已停止"
