#!/usr/bin/env bash
# Start development server with venv activation and uvicorn --reload
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="${SCRIPT_DIR}/.."
cd "${REPO_ROOT}"
# activate virtualenv if present
if [ -f ".venv/bin/activate" ]; then
  # shellcheck source=/dev/null
  source .venv/bin/activate
fi
# Use python from venv if available
PYTHON=${VIRTUAL_ENV:+"${VIRTUAL_ENV}/bin/python"}
if [ -z "${PYTHON}" ]; then
  PYTHON=python
fi
# Run uvicorn with reload for dev
exec $PYTHON -m uvicorn src.shiftortools.api:app --reload --host 127.0.0.1 --port 8000
