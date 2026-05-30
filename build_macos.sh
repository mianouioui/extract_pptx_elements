#!/bin/bash
# Build standalone macOS executable
# Prerequisites: pip3 install pyinstaller

set -e

cd "$(dirname "$0")"

echo "=== Building macOS executable ==="
pyinstaller --onefile --name extract_pptx_elements extract_pptx_elements.py

echo ""
echo "Done! Binary at: dist/extract_pptx_elements"
echo "Run: dist/extract_pptx_elements --help"
ls -lh dist/extract_pptx_elements
