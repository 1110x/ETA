#!/usr/bin/env bash
# Velopack 을 이용해 macOS 에서 Windows 설치 파일 생성
#   결과: Build/Releases/ETA-Setup.exe (+ 업데이트용 nupkg)
# 최초 1회:  dotnet tool install -g vpk
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
VERSION="${1:-1.4.5}"     # 첫 인자로 버전 지정. 기본 1.4.5
PUB="$ROOT/Build/publish"
REL="$ROOT/Build/Releases"

export PATH="$PATH:$HOME/.dotnet/tools"
# vpk 가 .NET 9 런타임을 요구할 때 설치된 상위 버전으로 롤포워드 허용
export DOTNET_ROLL_FORWARD="${DOTNET_ROLL_FORWARD:-LatestMajor}"

echo "📦  ETA v$VERSION Windows 설치 파일 생성 (Velopack)"

# 1) 기존 산출물 정리
rm -rf "$PUB" "$REL"
mkdir -p "$PUB"

# 2) .csproj 기준 publish (self-contained, single file)
echo "▶  dotnet publish ..."
dotnet publish "$ROOT/ETA.csproj" \
    -c Release -r win-x64 --self-contained \
    -p:PublishSingleFile=true \
    -p:IncludeNativeLibrariesForSelfExtract=true \
    -p:Version="$VERSION" \
    -o "$PUB"

# 3) 리소스 복사
echo "▶  리소스 복사 (Templates / Assets / appsettings)"
[ -d "$ROOT/Data/Templates" ] && { mkdir -p "$PUB/Data"; cp -R "$ROOT/Data/Templates" "$PUB/Data/"; }
[ -d "$ROOT/Assets"          ] && cp -R "$ROOT/Assets" "$PUB/Assets"
[ -f "$ROOT/appsettings.json" ] && cp "$ROOT/appsettings.json" "$PUB/"

# 4) vpk pack
if ! command -v vpk >/dev/null; then
    echo "❌  vpk 명령을 찾을 수 없습니다. 설치: dotnet tool install -g vpk"
    exit 1
fi

echo "▶  vpk pack (Windows 설치 파일 생성 중...)"
# 참고: macOS 호스트의 vpk 는 -i 옵션에 .icns 만 허용.
# 실행 파일·바로가기 아이콘은 ETA.csproj 의 <ApplicationIcon> 으로 이미 박혀있음.
vpk pack \
    -u ETA \
    -v "$VERSION" \
    -p "$PUB" \
    -e ETA.exe \
    -o "$REL"

echo ""
echo "✅  완료!  결과물:"
ls -lh "$REL" | tail -n +2

echo ""
echo "배포할 파일: $REL/ETA-win-Setup.exe"
