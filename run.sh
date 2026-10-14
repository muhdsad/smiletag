#!/usr/bin/env bash
# Retail Price Tag Generator Launcher
cd "$(dirname "$0")"

echo "======================================================="
echo "🏷️  Starting Smile PriceTag Master..."
echo "======================================================="

# Check if python3 is available
if command -v python3 &>/dev/null; then
    echo "Starting local Python server on http://localhost:8080..."
    # Launch browser after a brief delay in background
    (sleep 1.5 && (xdg-open "http://localhost:8080" 2>/dev/null || open "http://localhost:8080" 2>/dev/null || true)) &
    python3 server.py
else
    echo "Python3 not found. Opening index.html directly in browser..."
    xdg-open index.html 2>/dev/null || open index.html 2>/dev/null || sensible-browser index.html
fi
