#!/usr/bin/env bash
# macOS/Linux 에서 Windows 64비트용 self-contained 단일 exe 게시
# 출력: Build/publish/ETA.exe (+ 의존 파일 일부)
# 이후 Windows 에서 Build/installer.iss 로 설치 파일 생성
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$ROOT/Build/publish"

echo "📦  ETA Windows 게시 (self-contained, win-x64)"
rm -rf "$OUT"

dotnet publish "$ROOT/ETA.sln" \
    -c Release \
    -r win-x64 \
    --self-contained true \
    -p:PublishSingleFile=true \
    -p:IncludeNativeLibrariesForSelfExtract=true \
    -p:DebugType=embedded \
    -o "$OUT"

# Data / Templates 폴더는 exe 옆으로 복사 (설치 시 Program Files 하위 배포)
echo "📂  Data/Templates 복사"
mkdir -p "$OUT/Data"
cp -R "$ROOT/Data/Templates"          "$OUT/Data/" 2>/dev/null || true
cp -R "$ROOT/Assets"                  "$OUT/Assets" 2>/dev/null || true
cp    "$ROOT/appsettings.json"        "$OUT/" 2>/dev/null || true

echo "✅  게시 완료: $OUT"
echo ""
echo "다음 단계:"
echo "  1) 이 폴더를 Windows PC 에 복사"
echo "  2) Windows 에서 Inno Setup 으로 Build/installer.iss 컴파일"
echo "  3) 생성된 ETA-Setup-x.y.z.exe 배포"
