#!/bin/bash

# Script để chạy ứng dụng từ virtual environment
# Usage: ./run.sh [port]

PORT=${1:-8007}

echo "╔════════════════════════════════════════╗"
echo "║  Starting Application                  ║"
echo "║  Port: $PORT                            ║"
echo "╚════════════════════════════════════════╝"
echo ""

# Activate virtual environment
source venv/bin/activate

# Run application
PORT=$PORT python3 /home/vtst/s2/gsmnv.py

# Deactivate virtual environment when done
deactivate
