#!/usr/bin/env bash

echo "==> Installing system packages..."
apt-get update -y || true
apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-calc \
    tesseract-ocr \
    tesseract-ocr-eng \
    poppler-utils \
    ghostscript \
    fonts-liberation \
    fonts-dejavu \
    fonts-noto || true

echo "==> Installing Python packages..."
pip install --upgrade pip
pip install -r requirements.txt

echo "==> Build complete!"
