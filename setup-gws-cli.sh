#!/usr/bin/env bash
# GWS CLI (Google Workspace CLI) Installation & Setup Script
# Repository: https://github.com/googleworkspace/cli

set -e

echo "=== GWS CLI Setup ==="

# Check if gws is already installed
if command -v gws &> /dev/null; then
    echo "GWS CLI is already installed: $(gws --version 2>&1 | head -1)"
    exit 0
fi

# Install via cargo (build from source)
if command -v cargo &> /dev/null; then
    echo "Installing GWS CLI via Cargo (building from source)..."
    cargo install --git https://github.com/googleworkspace/cli --locked
    echo "GWS CLI installed successfully: $(gws --version 2>&1 | head -1)"
    echo ""
    echo "Next steps:"
    echo "  1. Place your OAuth client_secret.json in ~/.config/gws/"
    echo "  2. Run: gws auth login"
    echo "  3. Or set: export GOOGLE_WORKSPACE_CLI_TOKEN=<your-token>"
    exit 0
fi

# Install via npm
if command -v npm &> /dev/null; then
    echo "Installing GWS CLI via npm..."
    npm install -g @googleworkspace/cli
    echo "GWS CLI installed successfully: $(gws --version 2>&1 | head -1)"
    exit 0
fi

echo "Error: Neither cargo nor npm found. Please install one of them first."
echo "  - Cargo: https://rustup.rs/"
echo "  - Node.js/npm: https://nodejs.org/"
exit 1
