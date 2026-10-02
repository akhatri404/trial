#!/usr/bin/env bash
# Install dependencies and start the Streamlit app.
# Usage: ./run.sh            (local, http://localhost:8501)
#        PORT=8080 ./run.sh  (custom port, e.g. on a VM)
set -euo pipefail
cd "$(dirname "$0")"

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

exec streamlit run app.py --server.port="${PORT:-8501}" --server.address="${HOST:-0.0.0.0}" --server.headless=true
