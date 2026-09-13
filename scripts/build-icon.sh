#!/usr/bin/env bash
# Builds build/icon.icns from build/icon.svg using only macOS built-in tools.
# Usage: ./scripts/build-icon.sh
set -euo pipefail

cd "$(dirname "$0")/.."
SVG="build/icon.svg"
ICONSET="build/icon.iconset"
ICNS="build/icon.icns"
MASTER="build/icon-1024.png"

# 1. Render the SVG to a 1024px PNG with Quick Look.
rm -rf "$ICONSET" "$MASTER"
qlmanage -t -s 1024 -o build "$SVG" >/dev/null 2>&1
mv "build/icon.svg.png" "$MASTER"

# 2. Produce every size Apple expects in an .iconset.
mkdir -p "$ICONSET"
for size in 16 32 128 256 512; do
  double=$((size * 2))
  sips -z "$size" "$size"     "$MASTER" --out "$ICONSET/icon_${size}x${size}.png"    >/dev/null
  sips -z "$double" "$double" "$MASTER" --out "$ICONSET/icon_${size}x${size}@2x.png" >/dev/null
done

# 3. Convert the iconset to .icns.
iconutil -c icns "$ICONSET" -o "$ICNS"
rm -rf "$ICONSET"
echo "Wrote $ICNS"
