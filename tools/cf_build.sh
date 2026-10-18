#!/usr/bin/env bash
# คัดเฉพาะไฟล์ที่หน้าเว็บใช้ไปไว้ใน dist/ (ไม่เผยแพร่ sql/ tools/ supabase/ docs/ .github/)
#   Cloudflare Pages > Build command: bash tools/cf_build.sh · Build output directory: dist
set -euo pipefail
cd "$(dirname "$0")/.."
rm -rf dist && mkdir dist
for f in *.html; do cp "$f" dist/; done
for d in css js fonts img data; do [ -d "$d" ] && cp -r "$d" dist/; done
printf 'User-agent: *\nDisallow: /admin.html\n' > dist/robots.txt
echo "built -> dist/ ($(find dist -type f | wc -l) files)"
