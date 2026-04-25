#!/bin/zsh

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "CSV batch converter"
echo "Folder: $SCRIPT_DIR"
echo
echo "Opening the CSV batch converter window..."
echo

if [ -x /usr/bin/python3 ] && /usr/bin/python3 -c "import tkinter" >/dev/null 2>&1; then
  /usr/bin/python3 csv_batch_convert_gui.py
elif command -v python3 >/dev/null 2>&1; then
  python3 csv_batch_convert_gui.py
elif command -v python >/dev/null 2>&1; then
  python csv_batch_convert_gui.py
else
  echo "Error: Python is not installed or not available in PATH."
  echo "Please install Python 3, then run this command again."
fi

echo
echo "Window closed. Press Enter to close this terminal."
read -r _
