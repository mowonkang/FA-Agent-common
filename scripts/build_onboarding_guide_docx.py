"""온보딩_가이드.md → 배포용 Word(.docx) 빌더.

GitHub 접근 권한이 아직 없는 신규 구성원에게 메일/메신저로 배포하기 위한
사본을 만든다. 원본(source of truth)은 리포 루트의 `온보딩_가이드.md` 이며,
가이드 수정 후 본 스크립트를 다시 실행해 사본을 갱신한다.

    python3 scripts/build_onboarding_guide_docx.py

변환 엔진은 scripts/md_to_docx_agent_teams.py 의 build() 를 재사용한다
(빌더별 중복 정의 금지). docx 에서는 markdown 내부 앵커 링크(#...)가
의미 없으므로 변환 전에 링크 텍스트만 남기고 제거한다.
"""

from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from md_to_docx_agent_teams import build  # noqa: E402

SRC = ROOT / "온보딩_가이드.md"
CACHE = ROOT / "outputs" / ".cache"

INTERNAL_LINK = re.compile(r"\[([^\]]+)\]\(#[^)]*\)")


def main() -> int:
    if not SRC.exists():
        print(f"원본을 찾을 수 없습니다: {SRC}", file=sys.stderr)
        return 1

    body = SRC.read_text(encoding="utf-8")
    body = INTERNAL_LINK.sub(r"\1", body)

    CACHE.mkdir(parents=True, exist_ok=True)
    pre = CACHE / "온보딩_가이드.docx_input.md"
    pre.write_text(body, encoding="utf-8")

    out = ROOT / "outputs" / f"FA-Agent_온보딩_가이드_{date.today():%Y-%m-%d}.docx"
    try:
        build(pre, out)
    except Exception as exc:  # noqa: BLE001 — 배포 사본 실패는 원인과 함께 중단
        print(f"docx 변환 실패: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
