#!/usr/bin/env bash
# 프로젝트 폴더 구조 검증
# 사용법: platform/scripts/validate_project.sh <코드명>
set -uo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
CODE="${1:?코드명을 입력하세요}"
DEST="$ROOT/projects/$CODE"
[[ -d "$DEST" ]] || { echo "프로젝트 없음: $DEST" >&2; exit 1; }

missing=0
while IFS= read -r rel; do
  if [[ ! -e "$DEST/$rel" ]]; then
    echo "누락: $rel"; missing=$((missing+1))
  fi
done < <(cd "$ROOT/platform/templates/project" && find . -type d -o -type f -name '*.md' | sed 's|^\./||' | grep -v '^\.$')

if grep -rl '{{[A-Z_]*}}' "$DEST" --include='*.md' >/dev/null 2>&1; then
  echo "치환되지 않은 플레이스홀더 존재:"; grep -rn '{{[A-Z_]*}}' "$DEST" --include='*.md'
  missing=$((missing+1))
fi
[[ $missing -eq 0 ]] && echo "✅ $CODE 구조 정상" || { echo "❌ 문제 ${missing}건"; exit 1; }
