#!/bin/bash
# Build standalone macOS executable
# Prerequisites: pip3 install pyinstaller

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR/.."

echo "=== Building macOS executable ==="
mkdir -p build/specs
pyinstaller --onefile --name extract_pptx_elements --specpath build/specs extract_pptx_elements.py

echo ""
echo "Done! Binary at: dist/extract_pptx_elements"
echo "Run: dist/extract_pptx_elements --help"
ls -lh dist/extract_pptx_elements
