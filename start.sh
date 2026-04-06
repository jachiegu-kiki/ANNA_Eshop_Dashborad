#!/bin/bash
cd /opt/ecom-dashboard
source venv/bin/activate
PORT=$(grep "^PORT=" .env 2>/dev/null | cut -d'=' -f2 || echo "8888")
pkill -f "gunicorn.*app:app" 2>/dev/null || true
sleep 1
nohup gunicorn --bind 0.0.0.0:$PORT --workers 2 --timeout 120 \
  --access-logfile logs/access.log --error-logfile logs/error.log \
  app:app > logs/startup.log 2>&1 &
sleep 2
if pgrep -f "gunicorn.*app:app" > /dev/null; then
    echo "✅ 已启动 端口:$PORT"
else
    echo "❌ 启动失败"; cat logs/error.log | tail -20
fi
