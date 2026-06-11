"""references/ 폴더의 자료를 스캔해 INDEX.md + manifest.json 을 생성.

각 .md 파일에서 메타데이터(title/category/date/summary/tags/owner/related)를 추출.
frontmatter가 있으면 그 값을 우선 사용하고, 없는 필드는 본문/파일명/mtime 으로 자동 추출.

비코딩 사용자가 .md 파일을 폴더에 그냥 두기만 해도 인덱스에 등록되도록 설계되어 있다.

표준 라이브러리만 사용 (PyYAML 없이 정규식 파서).

사용:
    python3 scripts/build_reference_index.py            # 일반 실행
    python3 scripts/build_reference_index.py --quiet    # 출력 억제 (SessionStart hook용)
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REF_DIR = ROOT / "references"
INDEX_MD = REF_DIR / "INDEX.md"
MANIFEST_JSON = REF_DIR / "manifest.json"

# 카테고리 표시 순서
CATEGORY_ORDER = ["roadmap", "policy", "과거회의록", "기술자료", "경쟁사", "용어집", "기타"]

FRONTMATTER_RE = re.compile(r"\A---\s*\n(.*?)\n---\s*\n", re.DOTALL)
H1_RE = re.compile(r"^#\s+(.+?)\s*$", re.MULTILINE)
DATE_RE = re.compile(r"(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})")
SUMMARY_MAX_LEN = 200


# ────────────────────────────────────────────────────────────────
# Frontmatter 파서 (기존 유지)
# ────────────────────────────────────────────────────────────────
def parse_frontmatter(text: str) -> dict | None:
    """간이 YAML frontmatter 파서. dict / list / scalar 만 지원."""
    m = FRONTMATTER_RE.match(text)
    if not m:
        return None
    body = m.group(1)
    data: dict = {}
    current_key: str | None = None
    for raw_line in body.splitlines():
        line = raw_line.rstrip()
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        if line.startswith("  - ") or line.startswith("- "):
            item = line.split("- ", 1)[1].strip()
            item = item.strip("\"'")
            if current_key and isinstance(data.get(current_key), list):
                data[current_key].append(item)
            continue
        if ":" in line and not line.startswith(" "):
            key, _, value = line.partition(":")
            key = key.strip()
            value = value.strip()
            current_key = key
            if value == "":
                data[key] = []
            elif value.startswith("[") and value.endswith("]"):
                inner = value[1:-1].strip()
                if not inner:
                    data[key] = []
                else:
                    data[key] = [s.strip().strip("\"'") for s in inner.split(",")]
            else:
                data[key] = value.strip("\"'")
    return data


# ────────────────────────────────────────────────────────────────
# 휴리스틱 추출 (frontmatter 없을 때)
# ────────────────────────────────────────────────────────────────
def strip_frontmatter(text: str) -> str:
    """본문에서 frontmatter 블록 제거."""
    return FRONTMATTER_RE.sub("", text, count=1)


def extract_title(text: str, path: Path) -> tuple[str, str]:
    """(title, source) 반환. source ∈ {h1, filename}."""
    body = strip_frontmatter(text)
    m = H1_RE.search(body)
    if m:
        return m.group(1).strip(), "h1"
    # 파일명 stem, 언더스코어·하이픈을 공백으로
    return path.stem.replace("_", " ").replace("-", " "), "filename"


def extract_category(path: Path) -> tuple[str, str]:
    """(category, source) 반환. source ∈ {folder, fallback}."""
    try:
        parts = path.relative_to(REF_DIR).parts
    except ValueError:
        return "기타", "fallback"
    if len(parts) >= 2:
        return parts[0], "folder"
    return "기타", "fallback"


def extract_date(path: Path) -> tuple[str, str]:
    """(date YYYY-MM-DD, source) 반환. source ∈ {filename, mtime}."""
    m = DATE_RE.search(path.stem)
    if m:
        y, mo, d = m.groups()
        return f"{y}-{mo}-{d}", "filename"
    return (
        datetime.fromtimestamp(path.stat().st_mtime).date().isoformat(),
        "mtime",
    )


def extract_summary(text: str) -> tuple[str, str]:
    """(summary, source) 반환. source ∈ {paragraph, first_line, empty}."""
    body = strip_frontmatter(text)
    # 첫 H1 라인 제거 (제목이 요약에 또 들어가지 않게)
    body = H1_RE.sub("", body, count=1)
    # 단락(빈 줄로 구분) 단위로 보고, 헤딩·인용·표 라인은 스킵
    for para in re.split(r"\n\s*\n", body):
        clean = para.strip()
        if not clean:
            continue
        if clean.startswith(("#", ">", "|", "---")):
            continue
        # 여러 줄이면 공백으로 합침
        flat = " ".join(line.strip() for line in clean.splitlines() if line.strip())
        if not flat:
            continue
        if len(flat) > SUMMARY_MAX_LEN:
            return flat[:SUMMARY_MAX_LEN].rstrip() + "…", "paragraph"
        return flat, "paragraph"
    # 폴백: 첫 non-empty 줄
    for line in body.splitlines():
        clean = line.strip()
        if clean:
            if len(clean) > SUMMARY_MAX_LEN:
                return clean[:SUMMARY_MAX_LEN].rstrip() + "…", "first_line"
            return clean, "first_line"
    return "", "empty"


# ────────────────────────────────────────────────────────────────
# 스캔
# ────────────────────────────────────────────────────────────────
def build_entry(md_path: Path) -> dict:
    rel = md_path.relative_to(ROOT).as_posix()
    try:
        text = md_path.read_text(encoding="utf-8")
    except Exception as e:
        return {"path": rel, "error": f"read failed: {e}"}

    fm = parse_frontmatter(text) or {}
    extracted_from: dict[str, str] = {}

    # title
    if "title" in fm:
        title = fm["title"]
        extracted_from["title"] = "frontmatter"
    else:
        title, src = extract_title(text, md_path)
        extracted_from["title"] = src

    # category
    if "category" in fm:
        category = fm["category"]
        extracted_from["category"] = "frontmatter"
    else:
        category, src = extract_category(md_path)
        extracted_from["category"] = src

    # date
    if "date" in fm:
        date = fm["date"]
        extracted_from["date"] = "frontmatter"
    else:
        date, src = extract_date(md_path)
        extracted_from["date"] = src

    # summary
    if "summary" in fm:
        summary = fm["summary"]
        extracted_from["summary"] = "frontmatter"
    else:
        summary, src = extract_summary(text)
        extracted_from["summary"] = src

    entry: dict = {
        "path": rel,
        "title": title,
        "category": category,
        "date": date,
        "summary": summary,
    }

    # tags / owner / related — frontmatter 만
    if "tags" in fm:
        entry["tags"] = fm["tags"] if isinstance(fm["tags"], list) else [fm["tags"]]
    if "owner" in fm:
        entry["owner"] = fm["owner"]
    if "related" in fm:
        entry["related"] = fm["related"] if isinstance(fm["related"], list) else [fm["related"]]

    entry["_extracted_from"] = extracted_from
    return entry


def scan_references() -> list[dict]:
    entries: list[dict] = []
    for md_path in sorted(REF_DIR.rglob("*.md")):
        if md_path.name in ("README.md", "INDEX.md"):
            continue
        entries.append(build_entry(md_path))
    return entries


def discover_existing_category_dirs() -> list[str]:
    """references/ 아래의 실제 하위 폴더명을 반환 (자료 0개여도 포함)."""
    if not REF_DIR.exists():
        return []
    return sorted(
        p.name for p in REF_DIR.iterdir()
        if p.is_dir() and not p.name.startswith(".")
    )


def group_by_category(entries: list[dict]) -> dict[str, list[dict]]:
    groups: dict[str, list[dict]] = {}
    # 자료가 0개인 폴더도 빈 카테고리로 미리 등록
    for cat in discover_existing_category_dirs():
        groups.setdefault(cat, [])
    for e in entries:
        cat = e.get("category", "기타")
        groups.setdefault(cat, []).append(e)
    return groups


# ────────────────────────────────────────────────────────────────
# 파일 쓰기 (멱등)
# ────────────────────────────────────────────────────────────────
def write_if_changed(path: Path, new_content: str) -> bool:
    if path.exists():
        try:
            if path.read_text(encoding="utf-8") == new_content:
                return False
        except Exception:
            pass
    path.write_text(new_content, encoding="utf-8")
    return True


def build_index_md(groups: dict[str, list[dict]]) -> str:
    lines: list[str] = []
    lines.append("# references/ — 자료 인덱스")
    lines.append("")
    lines.append("이 파일은 `scripts/build_reference_index.py` 가 자동 생성합니다. 수동 편집하지 마세요.")
    lines.append("(마지막 갱신 시각은 git 커밋 기록으로 확인하세요.)")
    lines.append("")
    total = sum(len(v) for v in groups.values())
    lines.append(f"**총 {total} 건**")
    lines.append("")

    ordered_cats = [c for c in CATEGORY_ORDER if c in groups]
    ordered_cats += [c for c in groups.keys() if c not in CATEGORY_ORDER]
    for cat in ordered_cats:
        items = groups[cat]
        lines.append(f"## {cat} ({len(items)})")
        lines.append("")
        if not items:
            lines.append("_아직 자료가 없습니다. `references/" + cat + "/` 에 .md 파일을 두면 자동 등록됩니다._")
            lines.append("")
            continue
        lines.append("| 제목 | 날짜 | 태그 | 요약 | 경로 |")
        lines.append("|---|---|---|---|---|")
        for it in items:
            title = it.get("title", it.get("path", ""))
            date = it.get("date", "")
            tags = it.get("tags", [])
            tags_str = ", ".join(tags) if tags else ""
            summary = it.get("summary", "")
            path = it.get("path", "")
            err = it.get("error")
            mark = f" ⚠️ {err}" if err else ""
            # Markdown 표 셀 안의 파이프 이스케이프
            def esc(s):
                return str(s).replace("|", "\\|").replace("\n", " ")
            lines.append(
                f"| {esc(title)}{mark} | {esc(date)} | {esc(tags_str)} | {esc(summary)} | `{esc(path)}` |"
            )
        lines.append("")
    return "\n".join(lines)


def build_manifest_json(groups: dict[str, list[dict]], total: int) -> str:
    payload = {
        "total": total,
        "categories": list(groups.keys()),
        "by_category": groups,
    }
    return json.dumps(payload, ensure_ascii=False, indent=2) + "\n"


# ────────────────────────────────────────────────────────────────
# main
# ────────────────────────────────────────────────────────────────
def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--quiet", action="store_true",
                        help="suppress stdout (for SessionStart hook)")
    args = parser.parse_args()

    if not REF_DIR.exists():
        if not args.quiet:
            print(f"[skip] references/ folder not found at {REF_DIR}",
                  file=sys.stderr)
        return 0

    entries = scan_references()
    groups = group_by_category(entries)
    total = len(entries)

    index_changed = write_if_changed(INDEX_MD, build_index_md(groups))
    manifest_changed = write_if_changed(
        MANIFEST_JSON, build_manifest_json(groups, total))

    if not args.quiet:
        print(f"[OK] {total} reference(s) indexed")
        status = lambda c: "updated" if c else "unchanged"
        print(f"  → {INDEX_MD.relative_to(ROOT)} ({status(index_changed)})")
        print(f"  → {MANIFEST_JSON.relative_to(ROOT)} ({status(manifest_changed)})")
        errs = [e for e in entries if e.get("error")]
        if errs:
            print(f"[err] {len(errs)} file(s) failed:")
            for e in errs:
                print(f"  - {e['path']}: {e['error']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
