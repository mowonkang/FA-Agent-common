"""임원 보고 양식 카탈로그 빌더.

templates/사내양식/(보고양식)YYMMDD_보고제목_v1.0.pptx (LGES 임원 보고용
12종 슬라이드 양식 모음집) 를 scripts/ppt_extract.py 로 파싱한 뒤,
ppt-writer 가 소비할 수 있는 슬림 카탈로그 JSON 으로 정리한다.

- 12 슬라이드를 8개 재사용 보고 블록(archetype)으로 매핑.
- 각 슬라이드: 블록 id / 제목 / 용도 / placeholder 단서 / 표 형태 / 페이지 표시.
- 산출: templates/사내양식/보고양식_카탈로그.json (idempotent — 내용 동일 시 미기록)

양식·톤·색·사이즈는 '참고 가이드'이며 가독성·스토리를 위해 자율 조정 가능.
단 산출물 텍스트는 박스를 벗어나선 안 된다. (보고양식_카탈로그.md 참조)
"""
from __future__ import annotations

import json
import os
import re
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ppt_extract import extract  # noqa: E402

TEMPLATE_DIR = Path("templates/사내양식")
OUT_JSON = TEMPLATE_DIR / "보고양식_카탈로그.json"
SCHEMA_VERSION = "1.0"

# 슬라이드(0-base) → 8개 재사용 보고 블록 매핑
# (background 에이전트의 12-슬라이드 인벤토리 분석 결과 기반)
ARCHETYPES = [
    ("A", "보고 프레임 표준", "분류·제목·세션메시지·내용요약·풋노트 — "
     "메인 한 장 요약 레이아웃 기준", [0]),
    ("G", "As-Is / To-Be 전환", "현재 vs 미래 상태 비교 (적용 현황·"
     "Head message)", [1, 10]),
    ("B", "Master Plan 일정·로드맵", "분기 Gantt, Kick Off→상세설계→"
     "검증, 자동화 과제 지속 발굴", [2]),
    ("C", "투자비·예산 (Capex/KPI)", "설비 투자비 As-is/To-be, 억원, "
     "증감·비고, 레거시 버전 비교", [3]),
    ("F", "기술 사양·구성 최적화", "설비 구성 방안, 기존안/변경안/Concept, "
     "환경 조건, 개발 과제 사양·프로토타입 개선", [4, 8, 9]),
    ("D", "추진체계·조직도", "ES/PRI/CNS 매트릭스, 총원·상근·비상근, "
     "리더·스폰서 체계", [5]),
    ("E", "전략 추진방향·컨셉", "지향점·중점과제, Task 1.0/2.0, "
     "설비 컨셉(이미지 중심)", [6, 7]),
    ("H", "향후계획·검증", "보조 과제, Concept PoC 검증 항목·결과 "
     "로드맵, 참고 내용", [11]),
]

SLIDE_PURPOSE = {
    0: "보고 프레임 가이드 — 분류 1~4 / 제목 / 세션 메시지 / 내용 요약 / "
       "표·일정·보조자료 / Foot note 의 표준 레이아웃 (메인 요약 슬라이드 기준)",
    1: "보조 레이아웃 — As-is 적용 현황 vs To-be 비교",
    2: "Master Plan — 분기 타임라인(Kick Off·Ideation·컨셉구체화·상세설계·"
       "내구성/신뢰성 품평회), Pre Task·F-Test·Demo·Pilot 검증",
    3: "투자비 분석 — 물류/전극조립활성화/건설/설계용역 As-is·To-be·증감·"
       "비고 (억원), 레거시 1.0/2.0 비교",
    4: "설비 구성 방안 — 월별 설계·제작·구축·검증, Rack·메자닌·UT·벽체 "
       "구성·수량 매트릭스",
    5: "조직도 — 총원/상근/비상근, ES·PRI·CNS 팀, Co-Work, 리더·그룹장·"
       "센터장·사업부장·스폰서",
    6: "Task 추진방향 — 지향점·중점과제, Task 1.0/2.0, Re-Usability·"
       "유연성·원가혁신, 개발 항목 As-is/To-be",
    7: "설비 컨셉 — 이미지 중심 컨셉 제시 (멀티 페이지 섹션 표지)",
    8: "구성 최적화 — 기존안/변경안/Concept, 정면도·필요면적·온습도 조건, "
       "Alt1/Alt2 대안 비교",
    9: "개발 과제 — 운반물·사이즈·중량 사양, 반송 구간 요건, 프로토타입 "
       "경량화·Loss 저감 정량 개선",
    10: "과제 AS-IS / TO-BE — 프로세스·시스템 전환 비교 (Head message)",
    11: "향후계획 — 보조 과제, Concept PoC 검증 항목/결과/주요 검증 항목 "
        "로드맵",
}

# 자연어 placeholder 휴리스틱 패턴
NL_PATTERNS = [
    r"YYMMDD", r"보고제목", r"\bxx+\b", r"[xX]{3,}", r"ㅇㅇㅇ+", r"OOO+",
    r"d{4,}", r"XX설비", r"ESXX", r"ESMI|ESGM|ESWA|ESAZ", r"'?\d?\d?년",
    r"20YY", r"분류\s*\d", r"요약 내용", r"Head message", r"Foot ?note",
    r"세션 메시지", r"이름", r"장소",
]
NL_RE = [re.compile(p) for p in NL_PATTERNS]


def find_template() -> Path:
    """NFD/NFC 정규화 무시하고 (보고양식)...v1.0.pptx 를 찾는다."""
    target = unicodedata.normalize("NFC", "(보고양식)")
    for p in TEMPLATE_DIR.iterdir():
        name = unicodedata.normalize("NFC", p.name)
        if name.startswith(target) and name.endswith("v1.0.pptx"):
            return p
    raise FileNotFoundError("(보고양식)…v1.0.pptx 를 찾지 못했습니다")


def collect_cues(slide: dict) -> tuple[list[str], list[str], dict | None,
                                       bool, str | None]:
    cues: set[str] = set()
    explicit: set[str] = set()
    table_info = None
    has_image = False
    page_ind = None
    for sh in slide["shapes"]:
        txt = (sh.get("text") or "").strip()
        if sh.get("is_image"):
            has_image = True
        for k in sh.get("explicit_placeholders", []):
            explicit.add(k)
        if sh.get("highlight_color_rgb"):
            cues.add(f"강조색({sh['highlight_color_rgb']})")
        if txt:
            for rx in NL_RE:
                for m in rx.findall(txt):
                    if isinstance(m, tuple):
                        m = next((x for x in m if x), "")
                    if m:
                        cues.add(m.strip())
            pm = re.search(r"\b(\d)\s*/\s*(\d)\b", txt)
            if pm and page_ind is None:
                page_ind = pm.group(0).replace(" ", "")
        if sh.get("is_table") and sh.get("table") and table_info is None:
            t = sh["table"]
            table_info = {
                "rows": t["rows"],
                "cols": t["cols"],
                "header": [h for h in t["header"] if h],
            }
    return (sorted(cues)[:14], sorted(explicit), table_info,
            has_image, page_ind)


def build_catalog(struct: dict, src_rel: str) -> dict:
    slide_to_arch = {}
    arch_out = []
    for aid, name, desc, slides in ARCHETYPES:
        arch_out.append({
            "id": aid, "name": name, "description": desc,
            "slides": slides,
        })
        for s in slides:
            slide_to_arch[s] = aid

    slides_out = []
    for sl in struct["slides"]:
        idx = sl["slide_index"]
        cues, explicit, table_info, has_image, page_ind = collect_cues(sl)
        slides_out.append({
            "index": idx,
            "archetype_id": slide_to_arch.get(idx),
            "title": _title(sl),
            "purpose": SLIDE_PURPOSE.get(idx, ""),
            "layout_name": sl.get("layout_name"),
            "placeholder_cues": cues,
            "explicit_placeholders": explicit,
            "table": table_info,
            "has_image": has_image,
            "page_indicator": page_ind,
        })

    sz = struct["slide_size"]
    return {
        "schema_version": SCHEMA_VERSION,
        "source_file": src_rel,
        "note": ("양식·톤·색·사이즈는 참고 가이드이며 가독성·스토리를 위해 "
                 "자율 조정 가능. 단 산출물 텍스트는 박스를 벗어나선 안 됨. "
                 "상세는 templates/사내양식/보고양식_카탈로그.md 참조."),
        "slide_size_emu": sz,
        "slide_size_cm": {
            "width": round(sz["width_emu"] / 360000, 2),
            "height": round(sz["height_emu"] / 360000, 2),
        },
        "archetypes": arch_out,
        "slides": slides_out,
        "summary": {
            "slide_count": len(slides_out),
            "archetype_count": len(arch_out),
        },
    }


def _title(slide: dict) -> str:
    for sh in slide["shapes"]:
        t = (sh.get("text") or "").strip()
        if t and not sh.get("is_table"):
            return t.split("\n")[0][:60]
    return ""


def main() -> int:
    src = find_template()
    struct = extract(src)
    src_rel = f"templates/사내양식/{unicodedata.normalize('NFC', src.name)}"
    catalog = build_catalog(struct, src_rel)
    new = json.dumps(catalog, ensure_ascii=False, indent=2,
                     sort_keys=False)
    old = OUT_JSON.read_text(encoding="utf-8") if OUT_JSON.exists() else None
    # extracted_at 류 휘발 필드가 없으므로 내용 비교로 idempotent 보장
    if old is not None and json.loads(old) == json.loads(new):
        print(f"unchanged: {OUT_JSON}")
        return 0
    OUT_JSON.write_text(new, encoding="utf-8")
    print(f"saved: {OUT_JSON}  (slides={catalog['summary']['slide_count']}, "
          f"archetypes={catalog['summary']['archetype_count']})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
