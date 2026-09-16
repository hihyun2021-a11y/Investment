#!/usr/bin/env bash
# 신규 프로젝트 스캐폴딩
# 사용법: platform/scripts/new_project.sh <코드명> "<자산명>" [자산유형] [투자유형] [담당자]
#   예)   platform/scripts/new_project.sh GNG-TOWER "강남 G타워" 오피스 매입 홍길동
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
TEMPLATE="$ROOT/platform/templates/project"
REGISTRY="$ROOT/projects/REGISTRY.md"

CODE="${1:-}"
NAME="${2:-}"
ASSET_TYPE="${3:-미정}"
DEAL_TYPE="${4:-매입}"
OWNER="${5:-${PROJECT_OWNER:-운용역}}"
DATE="$(date +%Y-%m-%d)"

if [[ -z "$CODE" || -z "$NAME" ]]; then
  echo "사용법: $0 <코드명> \"<자산명>\" [자산유형] [투자유형] [담당자]" >&2
  exit 1
fi
if [[ ! "$CODE" =~ ^[A-Za-z0-9][A-Za-z0-9_-]*$ ]]; then
  echo "오류: 코드명은 영문·숫자·하이픈·밑줄만 사용 (예: GNG-TOWER)" >&2
  exit 1
fi

DEST="$ROOT/projects/$CODE"
if [[ -e "$DEST" ]]; then
  echo "오류: 이미 존재하는 프로젝트입니다: $DEST" >&2
  exit 1
fi

cp -R "$TEMPLATE" "$DEST"

# 플레이스홀더 치환
escape() { printf '%s' "$1" | sed -e 's/[\/&|]/\\&/g'; }
find "$DEST" -type f -name '*.md' -print0 | xargs -0 sed -i \
  -e "s|{{CODE}}|$(escape "$CODE")|g" \
  -e "s|{{NAME}}|$(escape "$NAME")|g" \
  -e "s|{{ASSET_TYPE}}|$(escape "$ASSET_TYPE")|g" \
  -e "s|{{DEAL_TYPE}}|$(escape "$DEAL_TYPE")|g" \
  -e "s|{{OWNER}}|$(escape "$OWNER")|g" \
  -e "s|{{DATE}}|$DATE|g"

# 레지스트리 등록
if [[ ! -f "$REGISTRY" ]]; then
  cat > "$REGISTRY" <<'HDR'
# 프로젝트 레지스트리

`platform/scripts/new_project.sh`가 자동 등록합니다. 상태·단계는 PROJECT.md와 동기화하세요.

| 코드명 | 자산명 | 자산유형 | 투자유형 | 담당 | 생성일 | 단계 | 상태 |
|---|---|---|---|---|---|---|---|
HDR
fi
printf '| [%s](%s/PROJECT.md) | %s | %s | %s | %s | %s | 1. 소싱/검토 | 진행중 |\n' \
  "$CODE" "$CODE" "$NAME" "$ASSET_TYPE" "$DEAL_TYPE" "$OWNER" "$DATE" >> "$REGISTRY"

cat <<MSG
✅ 프로젝트 생성 완료: projects/$CODE  ($NAME / $ASSET_TYPE / $DEAL_TYPE / 담당 $OWNER)

다음 단계 (킥오프 — Claude가 자동 수행):
  1. data-steward     : 00_원본자료/raw/ 자료 접수 → INDEX.md 등록, 정제
  2. pmo-scheduler    : TIMELINE.md 역산 일정 수립, ACTION_ITEMS 갱신
  3. finance-ir       : ASSUMPTIONS.md 초기 가정, QA_BANK 시드 질문
  4. legal-counsel    : 검토 대상 계약 목록화, ISSUE_LOG 초기화
  5. research-analyst : 1-pager 초안 골격
  6. decision-advisor : 첫 의사결정 아젠다 (진행 여부 Go/No-go 조건)

원본 자료를 projects/$CODE/00_원본자료/raw/ 에 넣고 "새 자료 접수 처리해줘"라고 요청하세요.
MSG
