"""TaskCompleted hook — outputs/*.md 출처 인용 검증.

Agent Teams 가이드 7.7절의 `TaskCompleted` 패턴을 본 리포의 절대원칙
("추측 금지·출처 명시")과 결합한 품질 게이트.

검증 대상: 작업 완료 직전(최근 180초 이내) outputs/ 에 수정된 .md 파일.
검증 규칙: 본문이 충분한 길이(>= 400 bytes) 라면 다음 출처 인용 패턴이
하나라도 포함되어야 한다.

    - `(출처:` 또는 `(출처 :`
    - `[N행]` (정수 N)
    - `참조:` 또는 `참고:` + 파일 경로
    - `자료 미확인` (명시적 인정)

위반 시 exit code 2 로 종료하여 작업 완료를 차단하고 stderr 로 피드백을 보낸다.
판정 불가능한 상황(파일 없음, hook 페이로드 파싱 실패 등)은 통과시킨다.
"""

from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
OUTPUTS = REPO_ROOT / "outputs"
RECENT_WINDOW_SEC = 180
MIN_BODY_BYTES = 400

CITATION_PATTERNS = [
    re.compile(r"\(출처\s*[:：]"),
    re.compile(r"\[\d+\s*행\]"),
    re.compile(r"참조\s*[:：].*\.(md|pdf|pptx|docx|csv|json)", re.IGNORECASE),
    re.compile(r"참고\s*[:：].*\.(md|pdf|pptx|docx|csv|json)", re.IGNORECASE),
    re.compile(r"자료\s*미확인"),
]


def recently_modified_md_files() -> list[Path]:
    if not OUTPUTS.exists():
        return []
    cutoff = time.time() - RECENT_WINDOW_SEC
    return sorted(
        (
            p
            for p in OUTPUTS.rglob("*.md")
            if p.is_file() and p.stat().st_mtime >= cutoff and ".cache" not in p.parts
        ),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )


def has_citation(text: str) -> bool:
    return any(pat.search(text) for pat in CITATION_PATTERNS)


def main() -> int:
    try:
        sys.stdin.read()  # drain payload; we don't rely on its schema
    except Exception:
        pass

    targets = recently_modified_md_files()
    if not targets:
        return 0

    offenders: list[tuple[Path, int]] = []
    for path in targets:
        try:
            body = path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        if len(body.encode("utf-8")) < MIN_BODY_BYTES:
            continue
        if not has_citation(body):
            offenders.append((path, len(body)))

    if not offenders:
        return 0

    msg = {
        "block": "missing_citations",
        "rule": "outputs/*.md 는 (출처: ...) 또는 [N행] 인용을 최소 1회 포함하거나, "
        "정보가 없는 항목에는 '자료 미확인' 을 명시해야 합니다.",
        "files": [
            {"path": str(p.relative_to(REPO_ROOT)), "bytes": n} for p, n in offenders
        ],
        "guide": "CLAUDE.md '절대 원칙' + Agent Teams 가이드 7.7절 TaskCompleted hook",
    }
    print(json.dumps(msg, ensure_ascii=False), file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main())
