#!/bin/bash
# تثبيت Tesseract OCR على Render
apt-get update && apt-get install -y tesseract-ocr tesseract-ocr-ara
pip install pytesseract pillow
pip install -r requirements.txt
