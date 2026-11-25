#!/bin/sh
# Ekspa İş Takibi - Linux/macOS başlatıcı
cd "$(dirname "$0")"
exec python3 app.py "$@"
