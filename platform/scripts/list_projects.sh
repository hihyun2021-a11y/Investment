#!/usr/bin/env bash
# 세션 시작 시 활성 프로젝트 현황 출력 (SessionStart 훅에서 호출)
set -uo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
PROJ="$ROOT/projects"

echo "=== 부동산 투자 플랫폼: 프로젝트 현황 ($(date +%Y-%m-%d)) ==="
shopt -s nullglob
dirs=("$PROJ"/*/)
if [[ ${#dirs[@]} -eq 0 ]]; then
  echo "등록된 프로젝트 없음. 신규 생성: /new-project <코드명> \"<자산명>\" [자산유형] [투자유형]"
  exit 0
fi
for d in "${dirs[@]}"; do
  code="$(basename "$d")"
  pm="$d/PROJECT.md"
  [[ -f "$pm" ]] || continue
  name="$(grep -m1 '^| 자산명' "$pm" | awk -F'|' '{gsub(/^ +| +$/,"",$3); print $3}')"
  stage="$(grep -m1 '^| 현재 단계' "$pm" | awk -F'|' '{gsub(/^ +| +$/,"",$3); print $3}')"
  status="$(grep -m1 '^| 상태' "$pm" | awk -F'|' '{gsub(/^ +| +$/,"",$3); print $3}')"
  raw_n="$(find "$d/00_원본자료/raw" -type f ! -name .gitkeep 2>/dev/null | wc -l)"
  idx_n="$(grep -c '^| [0-9]' "$d/00_원본자료/INDEX.md" 2>/dev/null)"; idx_n="${idx_n:-0}"
  open_ai="$(grep -E '^\| A-[0-9]+' "$d/07_일정_회의록/ACTION_ITEMS.md" 2>/dev/null | grep -vc '완료')"; open_ai="${open_ai:-0}"
  echo "- $code ($name) | $stage | $status | 원본 ${raw_n}건(인덱스 ${idx_n}건) | 미완료 액션 ${open_ai}건"
  if [[ "$raw_n" -gt "$idx_n" ]]; then
    echo "    ⚠ 미등록 원본 자료 $((raw_n - idx_n))건 → data-steward 접수 처리 필요"
  fi
done
echo "명령: /new-project (신규) · /project-status <코드명> (현황) · /ic-package <코드명> (투심 패키지)"
