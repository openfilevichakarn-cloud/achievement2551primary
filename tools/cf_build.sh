#!/usr/bin/env bash
# เว็บนี้เป็น static ไม่ต้อง build จริง แค่คัดลอกไฟล์ไปไว้ใน dist/
# Cloudflare Pages ตั้ง Build output directory เป็น dist (หรือ / ก็ใช้ได้)
set -euo pipefail
rm -rf dist
mkdir -p dist
cp -r index.html css js dist/
echo "built -> dist/"
