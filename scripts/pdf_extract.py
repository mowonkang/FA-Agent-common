"""PDF(.pdf) 파일에서 텍스트를 추출해 markdown 으로 변환한다.

각 페이지는 `## Page N` H2 섹션으로 출력된다. 표 추출은 지원하지 않으며,
표 위주 문서가 많다면 pdfplumber 기반 추출을 별도로 고려하라.

사용법:
    python3 scripts/pdf_extract.py <input.pdf> [--out <output.md>] [--pages <range>]

페이지 지정 예시:
    --pages 1-5        # 1쪽부터 5쪽까지
    --pages 1,3,5      # 1,3,5쪽만
    --pages 1-3,7,10-12  # 조합

기본 출력 위치는 입력 파일과 같은 폴더의 `<basename>.md` 이다.
페이지 번호는 1-based 이다.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

try:
    from pypdf import PdfReader
except ImportError:
    sys.stderr.write(
        "pypdf 가 설치되어 있지 않습니다. 먼저 다음을 실행하세요:\n"
        "    pip install -r requirements.txt\n"
    )
    sys.exit(2)


def parse_page_range(spec: str, total_pages: int) -> list[int]:
    """'1-3,7,10-12' 같은 문자열을 정렬된 1-based 페이지 번호 리스트로 변환."""
    pages: set[int] = set()
    for chunk in spec.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            start_str, end_str = chunk.split("-", 1)
            start, end = int(start_str), int(end_str)
            if start > end:
                start, end = end, start
            for p in range(start, end + 1):
                pages.add(p)
        else:
            pages.add(int(chunk))
    invalid = [p for p in pages if p < 1 or p > total_pages]
    if invalid:
        raise ValueError(
            f"잘못된 페이지 번호: {sorted(invalid)} (전체 {total_pages}쪽)"
        )
    return sorted(pages)


def _clean_text(text: str) -> str:
    """추출된 텍스트의 연속 공백·빈 줄을 정리."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def convert(input_path: Path, page_spec: str | None = None) -> str:
    """PDF 파일을 markdown 문자열로 변환."""
    reader = PdfReader(str(input_path))
    total = len(reader.pages)
    if page_spec:
        page_numbers = parse_page_range(page_spec, total)
    else:
        page_numbers = list(range(1, total + 1))

    parts: list[str] = [
        f"# {input_path.stem}",
        "",
        f"_총 {total}쪽 · 추출 페이지: {len(page_numbers)}쪽_",
        "",
    ]
    for pnum in page_numbers:
        page = reader.pages[pnum - 1]
        try:
            raw = page.extract_text() or ""
        except Exception as e:  # pypdf 가 일부 페이지에서 실패할 수 있음
            raw = f"_(추출 실패: {e})_"
        cleaned = _clean_text(raw)
        parts.append(f"## Page {pnum}")
        parts.append("")
        parts.append(cleaned if cleaned else "_(빈 페이지)_")
        parts.append("")
    return "\n".join(parts).rstrip() + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(
        description="PDF 파일에서 텍스트를 추출해 markdown 으로 변환합니다."
    )
    parser.add_argument("input", type=Path, help="입력 .pdf 파일")
    parser.add_argument(
        "--out", type=Path, default=None, help="출력 .md 경로 (기본: 입력과 같은 폴더)"
    )
    parser.add_argument(
        "--pages",
        default=None,
        help="추출할 페이지 (예: 1-5, 1,3,5, 1-3,7,10-12). 미지정 시 전체.",
    )
    args = parser.parse_args()

    if not args.input.exists():
        sys.stderr.write(f"입력 파일을 찾을 수 없습니다: {args.input}\n")
        return 1
    if args.input.suffix.lower() != ".pdf":
        sys.stderr.write(f"확장자가 .pdf 가 아닙니다: {args.input.suffix}\n")
        return 1

    output_path = args.out or args.input.with_suffix(".md")

    try:
        markdown = convert(args.input, page_spec=args.pages)
    except ValueError as e:
        sys.stderr.write(f"{e}\n")
        return 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(markdown, encoding="utf-8")
    print(f"변환 완료: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
