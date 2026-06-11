"""소개자료 PPT 채움용 values.json 생성기.

structure.json 을 스캔하여 각 placeholder 키에 대해 shape_id / table cell 을 자동 매칭하고,
content dict 의 값을 fill 항목으로 변환한다.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path


CONTENT: dict[str, str] = {
    # 표지
    "deck_title": "FA 기술혁신 Agent 팀",
    "subtitle": "주간보고·회의록·기술과제 발굴을 자동화한 5인 AI 협업조 (사내 3 + 사외 1 + 작성 2)",
    "author": "FA기술담당",
    "created_at": "2026-05-12",
    "audience": "FA기술담당 그룹장 / CPO",

    # 2. Executive Summary
    "summary_pt1": (
        "사내 정량(data) + 사내 운영(ops) + 사외 리서치(tech-research) + "
        "작성(document/ppt) 5명 에이전트가 우편함으로 직접 통신·develop."
    ),
    "summary_pt2": (
        "TaskCompleted Hook 이 outputs/*.md 의 출처 인용을 자동 검증, "
        "미기입 시 작업 자체가 차단된다."
    ),
    "summary_pt3": (
        "meeting-minutes / meeting-merger / fa-task-discovery 3개 도메인 "
        "스킬로 회의·기술 동향까지 자동화."
    ),
    "summary_pt4": (
        "주간보고 1건 작성 시간을 시간 단위 → 분 단위로 단축, "
        "출처 100% 검증된 표준 양식 산출."
    ),
    "summary_source": "CLAUDE.md, .claude/agents/*.md (ROI 정량값은 자료 미확인 — 정성 추정)",

    # 3. 문제 정의
    "problem_before": (
        "• 주간/월간 보고서 1건 작성에 수 시간 소요\n"
        "• 출처 누락된 수치가 보고서마다 반복\n"
        "• 양식이 사람마다 달라 표준 비교 불가\n"
        "• 회의록 정리 지연 → 결정사항·액션 추적 곤란\n"
        "• 신규 기술 동향 모니터링 미체계화\n"
        "• 자료 검색에 평균 30분~1시간 누적"
    ),
    "problem_after": (
        "• 보고서 1건 자동 생성 (분 단위)\n"
        "• 모든 수치에 (출처: 파일 [행번호]) 자동 첨부\n"
        "• templates/ 표준 양식 5종으로 일관성 확보\n"
        "• 회의록·녹취록 → PPTX/Word 자동 변환\n"
        "• fa-task-discovery 가 매일 6분야 자동 스캔\n"
        "• references/ 자동 인덱싱으로 검색 즉시"
    ),
    "problem_source": "자료 미확인 — 정량 비교 데이터는 그룹 내 PoC 후 보강 예정",

    # 4. 비교 (table)
    "comp_left_1": "1명이 단일 호출, 결과만 보고",
    "comp_right_1": "5명이 팀으로 동시 작동·develop, 우편함 통신",
    "comp_left_2": "리더 ↔ 팀원 1:1, 팀원 간 통신 없음",
    "comp_right_2": "팀원 간 직접 우편함 라우팅 (data → writer)",
    "comp_left_3": "순차 실행",
    "comp_right_3": "병렬 실행 (data + ops 동시 spawn)",
    "comp_left_4": "수동 검토",
    "comp_right_4": "TaskCompleted Hook 자동 검증 (출처 미기입 차단)",
    "comparison_source": "참조: outputs/Claude_Agent_Teams_가이드.md",

    # 5. 시스템 구성도
    "architecture_caption": (
        "사용자 → 팀장(메인 세션) → writer/ppt-writer 위임 작성. "
        "수집은 사내 3명(data·ops) + 사외 1명(tech-research) 병렬. "
        "SessionStart Hook 이 references/ 자동 인덱싱 (tech-research 가 저장한 신규 자료 포함), "
        "TaskCompleted Hook 이 출처 인용을 자동 검증한다."
    ),
    "architecture_source": "CLAUDE.md [17-42행], .claude/settings.json",

    # 6. 팀원 카드 (5명 — 사내 3 + 사외 1 + 작성 2)
    "team_card_5": (
        "[Read/Grep/Glob/Bash/Write/WebSearch/WebFetch + MCP] "
        "외부 기술 동향 · 경쟁사 · 기술 평가 담당. "
        "AMR/협동로봇/디지털트윈/스마트물류/AI/피지컬AI 6분야 글로벌+한국어 리서치. "
        "결과를 references/기술자료/ · references/경쟁사/ 에 frontmatter 포함 마크다운으로 "
        "자동 아카이빙 → SessionStart hook 이 다음 세션에 자동 인덱싱. "
        "신규 과제 발굴 · KPI 재설정 · 로드맵 검토 시 핵심 — 본 팀의 '외부 시야'."
    ),
    "team_card_1": (
        "data-teammate  [Read/Grep/Glob/Bash]\n\n"
        "사내 정량 — KPI · 로드맵 · Capex · MRM(양산) · DST(R&D).\n"
        "sample_data/ 와 references/roadmap/ 에서 항상 출처와 함께 인용 회신.\n"
        "외부 동향 요청은 tech-research-teammate 로 위임."
    ),
    "team_card_2": (
        "ops-teammate  [Read/Grep/Glob/Bash]\n\n"
        "사내 운영 — 협력사 · HR(출장·교육·포상) · 결재 큐 · 사내 PDF/PPT 검색.\n"
        "사람 승인 필요한 작업은 '제안' 형태만 가능.\n"
        "KPI·외부 동향 요청은 data / tech-research 로 위임."
    ),
    "team_card_3": (
        "document-writer  [Read/Write/Edit/Glob/Grep]\n\n"
        "templates/*.md 양식의 {{ placeholder }} 채움.\n"
        "data·ops·tech-research 에게 우편함으로 부족 데이터 직접 요청.\n"
        "plan-mode로 시작 → outputs/<이름>_<YYYY-MM-DD>.md 저장."
    ),
    "team_card_4": (
        "ppt-writer  [Read/Write/Edit/Glob/Grep/Bash]\n\n"
        "templates/*.pptx 의 5단계 placeholder 추론.\n"
        "ppt_extract.py + ppt_fill.py 결정론 파이프라인으로 새 PPTX 생성.\n"
        "외부 시장 데이터·경쟁사 도식은 tech-research 에 요청."
    ),
    "team_source": (
        ".claude/agents/data-teammate.md, ops-teammate.md, "
        "tech-research-teammate.md, document-writer.md, ppt-writer.md"
    ),

    # 7. 스킬 카드
    "skill_card_1": (
        "트리거: '녹취록으로 회의록', '회의록 PPTX'\n\n"
        "입력: 녹취록 텍스트\n"
        "출력: LGES 표준 회의록 PPTX\n\n"
        "5단계 워크플로우:\n"
        "  ① 녹취록 입력\n"
        "  ② 참석자 수집 (인터랙티브 UI)\n"
        "  ③ AI 요약 6~10건 생성\n"
        "  ④ 요약 선택·수정 UI\n"
        "  ⑤ generate_minutes.py → PPTX"
    ),
    "skill_card_2": (
        "트리거: 'OneNote랑 녹취록 합쳐줘', '회의 결과 Word'\n\n"
        "입력: 녹취록 + OneNote 메모\n"
        "출력: 7섹션 정제 Word docx\n\n"
        "구성:\n"
        "  ① 회의 기본 정보\n"
        "  ② 회의 요약 (3~5줄)\n"
        "  ③ 결정 사항 (표)\n"
        "  ④ 액션 아이템 (담당·기한)\n"
        "  ⑤ 이슈/리스크 (심각도)\n"
        "  ⑥ 일정표\n"
        "  ⑦ 전체 상세 내용"
    ),
    "skill_card_3": (
        "트리거: 'FA 기술 동향', '신규 과제 발굴', 'FA 브리핑'\n\n"
        "6분야 자동 스캔:\n"
        "  AMR/AGV · 협동로봇 · 디지털트윈\n"
        "  스마트물류 · AI · 피지컬 AI\n\n"
        "산출:\n"
        "  • docx 보고서 (과제 5~10건)\n"
        "  • 카카오톡 요약 자동 전송\n"
        "  • 매일 자동 스케줄"
    ),
    "skills_source": (
        ".claude/skills/meeting-minutes/SKILL.md, meeting-merger/SKILL.md, "
        "fa-task-discovery/SKILL.md"
    ),

    # 8. SOP
    "sop_step_1": (
        "리더가 요청을 도메인별 작업으로 쪼갠다. "
        "수치 ↔ 운영, 작성 ↔ 검토 축으로 분리."
    ),
    "sop_step_2": (
        "필요한 수집 팀원 (data·ops·tech-research) 을 동시 spawn 해 "
        "출처와 함께 데이터 수집. 사내·사외 자료가 같은 라운드에 합쳐진다."
    ),
    "sop_step_3": (
        "두 결과를 통합해 document-writer(.md) 또는 ppt-writer(.pptx) "
        "에게 전달. templates/ 양식의 placeholder 만 치환 — 양식 구조는 변경 금지."
    ),
    "sop_step_4": (
        "TaskCompleted Hook (check_output_citations.py) 이 outputs/*.md 의 "
        "출처 인용을 자동 검증. 위반 시 작업 완료 차단."
    ),
    "sop_caption": (
        "케이스별 구성: 정형보고 3명(data+ops+writer), "
        "브레인스토밍·KPI 재설정 4명(tech-research 추가). "
        "팀원당 5~6 작업이 적정 (Agent Teams 가이드)."
    ),
    "sop_source": "CLAUDE.md [26-42행]",

    # 9. 격리
    "iso_card_1": (
        "토큰 절감\n\n"
        "팀원 호출 시 시스템 프롬프트 + 본인 정의 + 호출자 프롬프트만 전달. "
        "리더의 누적 대화 기록은 미상속되므로 호출 1회당 토큰 사용량이 큰 폭으로 줄어든다."
    ),
    "iso_card_2": (
        "환각 차단\n\n"
        "참조 가능한 자료가 명확히 한정된다. references/INDEX.md 가 자동 인덱싱된 "
        "화이트리스트 역할을 하며, 추측 답변은 '자료 미확인' 으로 명시되어야 통과."
    ),
    "iso_card_3": (
        "스코프 명확\n\n"
        "호출자가 작업 범위·기준일·필요 출처를 프롬프트에 명시 → "
        "결과의 재현성·감사 가능성 확보. 후속 회수에서도 동일 산출 가능."
    ),
    "iso_source": "참조: .claude/agents/data-teammate.md [11행]",

    # 10. 검증
    "hook_snippet": (
        "TaskCompleted Hook 이 outputs/*.md \n"
        "(180초 내 수정, 400+ bytes)을 검사한다.\n\n"
        "검증 패턴 (1회 이상 필수):\n"
        "  • (출처: ...)   — 직접 인용\n"
        "  • [N행]          — 행 번호 인용\n"
        "  • 참조: ...      — 참조 표시\n"
        "  • 자료 미확인    — 명시적 보류\n\n"
        "위반 시 exit 2 → 작업 완료 차단."
    ),
    "pipeline_snippet": (
        "PPT 는 바이너리 직접 수정 금지.\n\n"
        "Step 1) ppt_extract.py <양식.pptx>\n"
        "  → outputs/.cache/<이름>.structure.json\n"
        "    (도형·표·색상·placeholder 5단계 추론)\n\n"
        "Step 2) ppt_fill.py\n"
        "  --template <양식.pptx>\n"
        "  --values   <values.json>\n"
        "  --out      outputs/<이름>_<날짜>.pptx\n"
        "    (원본 보존, 새 파일만 생성)"
    ),
    "verification_caption": (
        "두 메커니즘이 결합되어 출처 없는 보고서·구조 깨진 PPT 가 outputs/ 에 "
        "저장될 수 없다. 산출물의 신뢰성·재현성을 시스템 차원에서 보장한다."
    ),
    "verification_source": (
        "scripts/hooks/check_output_citations.py, "
        "scripts/ppt_extract.py, scripts/ppt_fill.py"
    ),

    # 11. capabilities (6 items — 단순 보고서를 넘어선 전략적 활용)
    "cap_card_1": (
        "정형 보고서 (주간 · 월간 · 회의록)\n\n"
        "산출: 주간보고 .md / 월간보고 .md / 회의록 .md\n"
        "협업: data + ops 병렬 수집 → writer 통합\n"
        "양식: templates/주간보고.md, 월간보고.md\n"
        "사례: outputs/주간보고_2026-05-12.md"
    ),
    "cap_card_2": (
        "과제 진척률 심층 보고\n\n"
        "DST(R&D) · MRM(양산) 진척률을 분기별로 정량화,\n"
        "references/roadmap/ 기준 격차 분석\n\n"
        "협업 흐름:\n"
        "  data → KPI·진척 추출\n"
        "  ↕ (우편함) 격차 검증\n"
        "  ops → 협력사·결재 현황 보완\n"
        "  → writer가 권고사항 종합"
    ),
    "cap_card_3": (
        "신규 과제 브레인스토밍\n\n"
        "외부 동향 + LGES 로드맵 매핑 + 협력사 풀 매칭으로\n"
        "PoC 가능 과제 도출\n\n"
        "협업 흐름:\n"
        "  tech-research → 6분야 외부 동향\n"
        "  + references/기술자료/ 자동 아카이빙\n"
        "  ↓\n"
        "  data ↔ 로드맵 매핑\n"
        "  ops ↔ 벤더 매칭\n"
        "  → writer가 과제 후보 docx + 우선순위"
    ),
    "cap_card_4": (
        "경쟁사 분석 → KPI 재설정 제안\n\n"
        "tech-research 가 references/경쟁사/ + 외부 리서치로\n"
        "격차 분석, 분기별 KPI 시뮬레이션 근거 수집\n\n"
        "협업 흐름:\n"
        "  ops → 사내 현재 KPI 추출\n"
        "  tech-research → 경쟁사 KPI · 글로벌 벤치마크\n"
        "  data → 사내 추이·시뮬 계산\n"
        "  writer ↔ data ↔ tech-research (라운드)\n"
        "  → 격차 분석 + KPI 재설정 제안서"
    ),
    "cap_card_5": (
        "녹취 기반 회의록 (PPTX / Word)\n\n"
        "meeting-minutes: 녹취 → LGES 표준 PPTX\n"
        "meeting-merger: 녹취+OneNote → 정제 Word\n\n"
        "사용자와 스킬이 인터랙티브 UI로 대화:\n"
        "참석자 입력 → AI 요약 → 선택·수정"
    ),
    "cap_card_6": (
        "FA 6분야 매일 자동 리서치\n\n"
        "AMR/협동로봇/디지털트윈/스마트물류/AI/피지컬AI\n\n"
        "협업 흐름:\n"
        "  naver-tech-researcher (한국어)\n"
        "  + tech-web-researcher (글로벌) 병렬 Task\n"
        "  → fa-task-discovery 통합\n"
        "  → docx 보고서 + 카카오톡 요약 (매일)"
    ),
    "capabilities_source": (
        "참조: templates/주간보고.md, 월간보고.md, 회의록.md; "
        ".claude/skills/*/SKILL.md; references/roadmap/, references/경쟁사/"
    ),

    # 12. 협업 시나리오 (NEW)
    "collab_scenario_1": (
        "요청: \"DST 과제 진척률 보고서 작성해줘\"\n\n"
        "[1] 팀장 → data-teammate (우편함)\n"
        "    \"DST 5과제 분기별 진척률 + 목표 격차\"\n\n"
        "[2] data 회신\n"
        "    \"DST-01 Isaac SIM 70%, 목표 -10%pt\n"
        "     DST-02 MMF 45%, 일정 1분기 지연…\"\n\n"
        "[3] 팀장 → ops-teammate (병렬 spawn)\n"
        "    \"5과제 협력사 풀·결재 큐 확인\"\n\n"
        "[4] data ↔ ops 직접 우편함 (develop!)\n"
        "    ops → data: \"MMF 협력사 0개\n"
        "    — 신규 KPI 연동 가능?\"\n"
        "    data → ops: \"references/policy 검토 후 회신\"\n\n"
        "[5] document-writer가 두 결과 통합\n"
        "    + references/과거회의록/ 교차 검증\n"
        "    → 진척률 표 + 격차 + 권고사항\n\n"
        "[6] TaskCompleted Hook 자동 검증\n"
        "    → outputs/DST_진척률_<날짜>.md"
    ),
    "collab_scenario_2": (
        "요청: \"협동로봇 동향 기반 신규 과제 +\n"
        "       KPI 재설정 제안서 만들어줘\"\n\n"
        "[1] 팀장 → tech-research-teammate\n"
        "    NaverSearch(한국어 뉴스) + WebSearch\n"
        "    (Universal Robots · ABB · FANUC) 병렬 수집\n"
        "    + references/기술자료/ 자동 아카이빙\n\n"
        "[2] tech-research → data 위임 (우편함)\n"
        "    \"이 6 동향이 LGES 로드맵 어디 매핑?\n"
        "     현재 KPI 와 격차?\"\n\n"
        "[3] data → references/경쟁사/ 비교\n"
        "    \"CATL 30%, LGES 12% → 격차 18%pt\"\n\n"
        "[4] ops 동시 우편함\n"
        "    \"PoC 가능 협력사 + 결재 큐 현황\"\n\n"
        "[5] writer ↔ tech-research ↔ data 라운드 (develop!)\n"
        "    writer → tech-research: \"글로벌 도입률\n"
        "    추가 사례 5건\"\n"
        "    writer → data: \"분기별 KPI 시뮬 (3년치)\"\n\n"
        "[6] 최종 산출\n"
        "    • 과제 후보 docx (우선순위)\n"
        "    • KPI 재설정 제안서 (외부 인용 포함)\n"
        "    • 카카오톡 요약 자동 전송"
    ),
    "collab_caption": (
        "핵심 포인트: 팀원은 격리되어 있지만 우편함으로 직접 소통한다. "
        "동일 데이터를 여러 라운드에 걸쳐 정제·심화 — 사람이 회의에서 하는 협업과 동일한 패턴을 자동화."
    ),
    "collab_source": (
        "참조: CLAUDE.md [26-42행] 팀 운영 SOP, "
        ".claude/agents/document-writer.md (우편함 라우팅)"
    ),

    # 12. 데모 주간보고
    "demo_weekly_points": (
        "데모 포인트\n\n"
        "• KPI 4지표 (목표·실적·달성률) 표 자동 채움\n"
        "• Capex 4건, 협력사 신규/변경, DST 진척 5과제\n"
        "• 모든 항목 끝에 (출처: references/26FA KPI.md [N행]) 자동 첨부\n"
        "• TaskCompleted Hook 자동 통과 — 출처 미기입 0건\n"
        "• 작성에 사람 손이 닿은 부분: 0 (전 자동)\n\n"
        "사람이 했더라면? 통상 2~3 시간\n"
        "본 팀: 약 10분 이내"
    ),
    "demo_weekly_source": "참조: outputs/주간보고_2026-05-12.md",

    # 13. 데모 로드맵
    "demo_roadmap_caption": (
        "좌: roadmap_v2.pptx 의 4축 KPI 슬라이드 (CAPEX 4트랙·OPEX 3트랙). "
        "우: MMF Pop-up Factory 컨셉 SVG. "
        "두 산출물 모두 본 팀의 결정론 파이프라인으로 생성."
    ),
    "demo_roadmap_source": (
        "참조: outputs/roadmap_v2.pptx, "
        "outputs/MMF_컨셉이미지/03_PopUp_Factory_이동성.svg"
    ),

    # 14. ROI · 다음 단계
    "roi_kpi_1": "수 시간 → 수 분\n(자료 미확인 — 정성 추정)",
    "roi_kpi_2": "100%\n(Hook 자동 차단으로 보장)",
    "roi_kpi_3": "templates/ 5종\n(주간/월간/회의/PPT)",
    "next_step_1": (
        "그룹 내 타 팀 PoC — 동일 SOP 를 품질·구매 보고서 영역으로 확장"
    ),
    "next_step_2": (
        "templates/ 추가 — MRM 양식, Capex 승인 양식, 분기 임원보고 양식"
    ),
    "next_step_3": (
        "sample_data/ 와 references/ 보강 — 실데이터 연동, 자동 ETL 검토"
    ),
    "next_step_4": (
        "fa-task-discovery 매일 스케줄 정착 → 월간 신규 과제 발굴 정량 추적"
    ),

    # 15. 부록 회의록
    "appendix_minutes_points": (
        "회의록_신규R&D과제착수_2026-05-12.md\n\n"
        "• Round 1/2 논의 자동 정리\n"
        "• 결정사항 (1순위 Isaac SIM, 2순위 MMF) 추출\n"
        "• 액션아이템 담당·기한 표 형식\n"
        "• 모든 결정 항목에 references/roadmap/ 출처 명시\n\n"
        "data + ops + document-writer 3명 협업 사례"
    ),
    "appendix_minutes_source": "참조: outputs/회의록_신규R&D과제착수_2026-05-12.md",

    # 16. 부록 FA
    "appendix_fa_points": (
        "FA_기술과제후보_20260512.docx\n\n"
        "• AMR/협동로봇/디지털트윈/스마트물류/AI/피지컬AI 6분야 동시 리서치\n"
        "• 각 과제: 카테고리·과제명·동향 근거·LGES 적용 포인트·글로벌 사례\n"
        "• 동시에 카카오톡 요약 자동 전송\n"
        "• 매일 자동 스케줄로 월간 누적"
    ),
    "appendix_fa_source": (
        "참조: outputs/FA_기술과제후보_20260512.docx, "
        ".claude/skills/fa-task-discovery/SKILL.md"
    ),
}


def build_values(structure_path: Path) -> dict:
    structure = json.loads(structure_path.read_text(encoding="utf-8"))
    fills: list[dict] = []
    unfilled: list[str] = []
    fill_id = 0
    used_keys: set[str] = set()

    for slide in structure["slides"]:
        slide_idx = slide["slide_index"]
        for shape in slide["shapes"]:
            sid = shape["shape_id"]
            # Shape-level placeholders
            for key in shape.get("explicit_placeholders", []):
                if shape.get("is_table"):
                    continue  # handled in table loop
                value = CONTENT.get(key)
                if value is None:
                    unfilled.append(f"{sid} :: {key}")
                    continue
                fill_id += 1
                fills.append({
                    "fill_id": f"f{fill_id:03d}",
                    "mode": "replace_placeholder",
                    "explicit_placeholder": key,
                    "value": value,
                    "source": "scripts/build_intro_values.py CONTENT",
                    "target": {
                        "kind": "shape_text",
                        "shape_id": sid,
                    },
                })
                used_keys.add(key)
            # Table cells
            if shape.get("is_table") and shape.get("table"):
                for row in shape["table"]["cells"]:
                    for cell in row:
                        for key in cell.get("explicit_placeholders", []):
                            value = CONTENT.get(key)
                            if value is None:
                                unfilled.append(
                                    f"{sid} table[{cell['row']},{cell['col']}] :: {key}"
                                )
                                continue
                            fill_id += 1
                            fills.append({
                                "fill_id": f"f{fill_id:03d}",
                                "mode": "replace_placeholder",
                                "explicit_placeholder": key,
                                "value": value,
                                "source": "scripts/build_intro_values.py CONTENT",
                                "target": {
                                    "kind": "table_cell",
                                    "shape_id": sid,
                                    "row": cell["row"],
                                    "col": cell["col"],
                                },
                            })
                            used_keys.add(key)

    # Image slots etc. (intentionally unfilled for manual insertion)
    image_slot_names = [
        "demo_weekly_image_slot",
        "appendix_minutes_slot",
        "appendix_fa_slot",
        "demo_roadmap_slot1",
        "demo_roadmap_slot2",
    ]

    return {
        "schema_version": "1.0",
        "structure_source": str(structure_path),
        "fills": fills,
        "unfilled": unfilled + [
            f"image_slot::{n}" for n in image_slot_names
        ],
        "stats": {
            "fills_total": len(fills),
            "content_keys_total": len(CONTENT),
            "content_keys_used": len(used_keys),
            "content_keys_unused": sorted(set(CONTENT) - used_keys),
        },
    }


def main(argv: list[str] | None = None) -> int:
    structure_path = Path("outputs/.cache/소개자료.structure.json")
    out_path = Path("outputs/.cache/소개자료.values.json")
    if not structure_path.exists():
        sys.stderr.write(f"structure 파일 없음: {structure_path}\n")
        sys.stderr.write(
            "먼저 다음을 실행: python scripts/ppt_extract.py templates/소개자료.pptx\n"
        )
        return 1

    values = build_values(structure_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(values["stats"], ensure_ascii=False, indent=2))
    print(f"unfilled total: {len(values['unfilled'])}")
    print(f"saved: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
