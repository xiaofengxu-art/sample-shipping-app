#!/bin/zsh

cd "$(dirname "$0")" || exit 1

URL="http://localhost:8501"

echo "正在检查依赖..."
if ! python3 -c "import streamlit, openpyxl, pandas" >/dev/null 2>&1; then
  echo "依赖缺失，开始安装（首次可能需要一两分钟）..."
  if command -v pip3 >/dev/null 2>&1; then
    pip3 install -r requirements.txt || exit 1
  else
    python3 -m pip install -r requirements.txt || exit 1
  fi
else
  echo "依赖已就绪，跳过安装。"
fi

# 已经在运行就别再起一个，直接开浏览器
if curl -fsS --max-time 2 "$URL/_stcore/health" >/dev/null 2>&1; then
  echo "检测到工具已在运行，直接打开浏览器。"
  open "$URL"
  exit 0
fi

echo "正在启动样品寄送工具..."
(
  for i in {1..30}; do
    if curl -fsS --max-time 1 "$URL/_stcore/health" >/dev/null 2>&1; then
      open "$URL"
      exit 0
    fi
    sleep 0.5
  done
) &

python3 -m streamlit run sample_shipping_app.py --server.headless true
