#!/bin/bash
# Startup script for Input Studio Web Server

# Activate virtual environment if it exists
if [ -d ".venv" ]; then
    source .venv/bin/activate
fi

# Set environment variables
export PORT=${PORT:-8000}
export PYTHONUNBUFFERED=1

# Run the server
python server/main.py
