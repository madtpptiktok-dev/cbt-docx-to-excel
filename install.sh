#!/bin/bash

set -e

echo "=== CBT DOCX TO EXCEL INSTALLER ==="

sudo apt update
sudo apt install -y python3-pip python3-venv python3-tk

python3 -m venv .venv

source .venv/bin/activate

pip install --upgrade pip
pip install -r requirements.txt

echo
echo "=== INSTALL SELESAI ==="
echo "Jalankan aplikasi dengan:"
echo "source .venv/bin/activate && python cbt_docx_to_excel_tool.py"
