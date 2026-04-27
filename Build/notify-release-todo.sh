#!/usr/bin/env bash
# 릴리즈 빌드(GitHub Actions) 완료를 폴링하다가, 끝나면 Microsoft To Do
# "ETA 개발" 리스트에 작업을 등록한다. 작업 메모에는 인스톨러 직접 다운로드
# URL · 파일명 · 크기, GitHub Release / 워크플로 링크가 포함된다.
#
# 사용:
#   ./Build/notify-release-todo.sh v1.4.2          # 특정 태그
#   ./Build/notify-release-todo.sh                 # 최신 태그 자동 사용
#
# 의존:
#   - python3 (urllib 만 사용 — 별도 패키지 불필요)
#   - 같은 리포의 Services/SERVICE1/TodoService.cs 에서 client_id / refresh_token 추출
#     ※ ETA_TODO_REFRESH_TOKEN / ETA_TODO_CLIENT_ID 환경변수가 있으면 그걸 우선
#
# 백그라운드 실행 권장:
#   nohup ./Build/notify-release-todo.sh v1.4.3 > /tmp/eta-todo.log 2>&1 &
set -euo pipefail

REPO="1110x/ETA"
LIST_NAME="ETA 개발"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
TOKEN_SRC="$ROOT/Services/SERVICE1/TodoService.cs"

# 1) 태그 결정
TAG="${1:-}"
if [ -z "$TAG" ]; then
    TAG="$(git -C "$ROOT" describe --tags --abbrev=0 2>/dev/null || true)"
    [ -z "$TAG" ] && { echo "❌ 태그를 결정할 수 없습니다. 인자로 태그명을 넘기세요."; exit 1; }
fi
echo "▶  대상 태그: $TAG"

# 2) 워크플로 run id 조회 (해당 태그의 가장 최근 실행)
RUN_ID="$(curl -s "https://api.github.com/repos/$REPO/actions/runs?per_page=20" \
    | python3 -c "
import json, sys
runs = json.load(sys.stdin).get('workflow_runs', [])
tag = '$TAG'
for r in runs:
    if r.get('head_branch') == tag:
        print(r['id']); break")"
[ -z "$RUN_ID" ] && { echo "❌ $TAG 태그의 워크플로 run 을 찾지 못했습니다."; exit 1; }
echo "▶  워크플로 run id: $RUN_ID"

# 3) 완료 대기 (30초 간격 폴링)
echo "⏳  빌드 완료 대기..."
until [ "$(curl -s "https://api.github.com/repos/$REPO/actions/runs/$RUN_ID" \
    | python3 -c "import json,sys; print(json.load(sys.stdin).get('status','?'))")" = "completed" ]; do
    sleep 30
done

# 4) 결과 + 다운로드 URL + To Do 등록 (Python 한 번에)
TAG="$TAG" RUN_ID="$RUN_ID" REPO="$REPO" LIST_NAME="$LIST_NAME" \
TOKEN_SRC="$TOKEN_SRC" python3 <<'PYEOF'
import os, re, urllib.request, urllib.parse, json, sys

TAG       = os.environ["TAG"]
RUN_ID    = os.environ["RUN_ID"]
REPO      = os.environ["REPO"]
LIST_NAME = os.environ["LIST_NAME"]
TOKEN_SRC = os.environ["TOKEN_SRC"]

# ── 자격증명: 환경변수 우선, 없으면 TodoService.cs 에서 추출
client_id     = os.environ.get("ETA_TODO_CLIENT_ID", "")
refresh_token = os.environ.get("ETA_TODO_REFRESH_TOKEN", "")
if not (client_id and refresh_token):
    try:
        src = open(TOKEN_SRC, "r", encoding="utf-8").read()
        if not client_id:
            m = re.search(r'ClientId\s*=\s*"([^"]+)"', src);     client_id     = m.group(1) if m else ""
        if not refresh_token:
            m = re.search(r'RefreshToken\s*=\s*"([^"]+)"', src);  refresh_token = m.group(1) if m else ""
    except Exception as e:
        sys.exit(f"❌ 자격증명 로드 실패: {e}")
if not (client_id and refresh_token):
    sys.exit("❌ ClientId / RefreshToken 을 찾을 수 없습니다.")

# ── 워크플로 결과
run = json.loads(urllib.request.urlopen(
    f"https://api.github.com/repos/{REPO}/actions/runs/{RUN_ID}").read())
result = run.get("conclusion") or "?"

# ── 릴리즈 에셋 (성공 시)
asset_url = asset_name = ""; size_mb = 0.0
try:
    rel = json.loads(urllib.request.urlopen(
        f"https://api.github.com/repos/{REPO}/releases/tags/{TAG}").read())
    for a in rel.get("assets", []):
        if a["name"].lower().endswith(".exe"):
            asset_url  = a["browser_download_url"]
            asset_name = a["name"]
            size_mb    = a["size"] / 1024 / 1024
            break
except Exception:
    pass

# ── 액세스 토큰
body = urllib.parse.urlencode({
    "client_id": client_id, "grant_type": "refresh_token",
    "refresh_token": refresh_token,
    "scope": "Tasks.ReadWrite User.Read offline_access",
}).encode()
req = urllib.request.Request(
    "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
    data=body, method="POST",
    headers={"Content-Type": "application/x-www-form-urlencoded"})
access = json.loads(urllib.request.urlopen(req).read())["access_token"]

# ── 리스트 ID
req = urllib.request.Request("https://graph.microsoft.com/v1.0/me/todo/lists",
    headers={"Authorization": f"Bearer {access}"})
lists = json.loads(urllib.request.urlopen(req).read())
target_id = next((l["id"] for l in lists.get("value", [])
                  if l.get("displayName") == LIST_NAME), None)
if not target_id:
    sys.exit(f"❌ '{LIST_NAME}' 리스트가 없습니다.")

# ── 작업 본문
icon  = "✅" if result == "success" else "❌"
lines = [f"GitHub Actions: {result}", "", "📦 배포 파일 다운로드"]
if asset_url:
    lines += [f"   {asset_url}", f"   파일명: {asset_name}  ({size_mb:.1f} MB)"]
else:
    lines += ["   (릴리즈 에셋 미발견 — 워크플로 로그 확인 필요)"]
lines += [
    "",
    f"릴리즈 페이지: https://github.com/{REPO}/releases/tag/{TAG}",
    f"워크플로:     https://github.com/{REPO}/actions/runs/{RUN_ID}",
]

task = {
    "title": f"{icon} {TAG} 릴리즈 빌드 {result}",
    "body":  {"contentType": "text", "content": "\n".join(lines)},
}
req = urllib.request.Request(
    f"https://graph.microsoft.com/v1.0/me/todo/lists/{target_id}/tasks",
    data=json.dumps(task).encode(), method="POST",
    headers={"Authorization": f"Bearer {access}",
             "Content-Type": "application/json"})
urllib.request.urlopen(req).read()
print(f"{icon} {TAG} 빌드 {result} → '{LIST_NAME}' 리스트에 작업 등록 완료")
PYEOF
