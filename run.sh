#!/bin/zsh
set -e

cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  echo "首次运行：正在创建 Python 虚拟环境..."
  python3 -m venv .venv
fi

echo "正在检查依赖..."
.venv/bin/python -m pip install --disable-pip-version-check -q -r requirements.txt

echo "启动行程单生成脚本..."
.venv/bin/python generate_itinerary.py "$@"
