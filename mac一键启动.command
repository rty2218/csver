#!/bin/zsh

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "CSV batch converter"
echo "Folder: $SCRIPT_DIR"
echo
echo "Converting all CSV files in this folder to XLSX, TXT, and Markdown..."
echo "Output folder: $SCRIPT_DIR/converted"
echo

if command -v python3 >/dev/null 2>&1; then
  python3 csv_batch_convert.py --infer-types
elif command -v python >/dev/null 2>&1; then
  python csv_batch_convert.py --infer-types
else
  echo "Error: Python is not installed or not available in PATH."
  echo "Please install Python 3, then run this command again."
fi

echo
echo "Done. Press Enter to close this window."
read -r _
