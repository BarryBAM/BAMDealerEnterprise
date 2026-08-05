#!/usr/bin/env bash
set -euo pipefail
export BAM_DATA_ROOT="${BAM_DATA_ROOT:-/home/data}"
export BAM_SECURE_COOKIES="${BAM_SECURE_COOKIES:-1}"
mkdir -p "$BAM_DATA_ROOT/database" "$BAM_DATA_ROOT/uploads" "$BAM_DATA_ROOT/backups" "$BAM_DATA_ROOT/reports"
exec gunicorn --bind "0.0.0.0:${PORT:-8000}" --workers 1 --threads 4 --timeout 180 --access-logfile - --error-logfile - app.app:app
