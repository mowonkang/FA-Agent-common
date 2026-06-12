"""FA-Agent 온보딩 가이드 → 전 구성원 발표용 PPTX 빌더.

온보딩_가이드.md (리포 루트) 의 내용을 사내 표준 분류형 레이아웃으로 변환한다.
GitHub·Claude·코드를 모르는 전 구성원 대상 교육/발표 세션용.

  - 메인        : 온보딩 한눈 요약 (무엇인가 / 준비 / 사용법 / 규칙, 분류형)
  - 보조 1/7    : 개념 — 등장인물 3개·AI 팀원 5명·폴더 4개
  - 보조 2/7    : 처음 1회 준비 — 계정 2개 절차·체크리스트
  - 보조 3/7    : 사용 방법 선택 — 웹(권장) vs PC 설치
  - 보조 4/7    : 첫 사용 5분 — 동작 확인 명령·화면 대응·결과물 받기
  - 보조 5/7    : 업무별 요청 문장 치트시트
  - 보조 6/7    : 잘 쓰는 요령·팀 규칙·보안
  - 보조 7/7    : FAQ·문제 해결·오늘 시작하기

데이터 출처(추측 없음): 온보딩_가이드.md · CLAUDE.md · README.md
[106~113, 149~158, 186~196행] · scripts/hooks/check_output_citations.py.
출력: outputs/FA-Agent_온보딩_발표자료_2026-06-12.pptx
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Cm

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ppt_layout import (  # noqa: E402
    render_classified_slide, title_block, add_label_band, add_text,
    fill_table, add_footnote, SLIDE_W_CM, SLIDE_H_CM,
    BLACK, BLUE, MID_GRAY, CHARCOAL, SZ_HEAD, SZ_BODY, SZ_SUB,
)

OUT_PATH = Path("outputs/FA-Agent_온보딩_발표자료_2026-06-12.pptx")
BYLINE = ["FA기술혁신 Part", "(2026.06.12)"]
FULL_X, FULL_W = 0.6, 26.3


def band_table(slide, y, label, headers, rows, col_w, row_h, *,
               body=SZ_BODY, head=SZ_BODY, band_h=0.62, emph_col0=True):
    """전폭 라벨 띠 + 표 블록. 표 하단 y 를 반환."""
    add_label_band(slide, FULL_X, y, FULL_W, label, height=band_h)
    ty = y + band_h + 0.08
    nrows = len(rows) + 1
    shape = slide.shapes.add_table(
        nrows, len(headers), Cm(FULL_X), Cm(ty),
        Cm(sum(col_w)), Cm(nrows * row_h))
    fill_table(shape.table, headers, rows, col_w=col_w,
               body_size=body, head_size=head, emph_col0=emph_col0)
    for r in range(nrows):
        shape.table.rows[r].height = Cm(row_h)
    return ty + nrows * row_h


def band_text(slide, y, label, lines, *, band_h=0.62, box_h=2.0):
    """전폭 라벨 띠 + 텍스트 블록. 블록 하단 y 를 반환."""
    add_label_band(slide, FULL_X, y, FULL_W, label, height=band_h)
    add_text(slide, FULL_X + 0.05, y + band_h + 0.10, FULL_W - 0.1,
             box_h, lines)
    return y + band_h + 0.10 + box_h


# ───────────────────────── 메인 (분류형 요약) ─────────────────────────
def build_main(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    render_classified_slide(
        slide,
        title="FA-Agent 온보딩 — 전 구성원 사용 가이드",
        byline=BYLINE,
        head_message=[
            ("한국어로 말하면 AI 팀원 5명이 사내 양식대로 보고서를 만들어 주는 "
             "시스템입니다.",
             {"size": SZ_HEAD, "bold": True, "color": BLACK}),
            ("처음 1회 계정 2개 준비(약 30분) 후에는, 채팅 한 문장으로 회의록·"
             "주간보고·동향조사를 시킬 수 있습니다.",
             {"size": SZ_HEAD, "bold": True, "color": BLACK}),
        ],
        sections=[
            {
                "label": "1\n무엇인가",
                "session_message": "Claude(AI 직원) + Claude Code(작업 환경) + "
                                   "GitHub(팀 공유 저장소) 위에 구축된 AI 팀원 5명",
                "content": [
                    (" • 데이터·운영·리서치·문서·PPT 5개 역할 분담 — 팀장(메인 "
                     "세션)이 적임자에게 자동 배분 (→ 보조 1)",
                     {"size": SZ_BODY, "color": BLACK}),
                    (" • 코드·명령어 학습 불필요 — 평소 말투의 한국어로 업무를 "
                     "지시하면 됩니다",
                     {"size": SZ_BODY, "bold": True, "color": BLUE}),
                ],
            },
            {
                "label": "2\n준비\n(1회)",
                "session_message": "계정 2개만 준비하면 끝 — GitHub(저장소 초대) + "
                                   "Claude(시트 배정), 약 30분 (→ 보조 2)",
                "content": [
                    (" • GitHub 가입 → 관리자에게 FA-Agent-common 초대 요청 → "
                     "초대 수락",
                     {"size": SZ_BODY, "color": BLACK}),
                    (" • Claude 가입 → Claude Code 권한 포함 시트 배정 (회사 신청 "
                     "절차는 담당자 확인 — 자료 미확인)",
                     {"size": SZ_BODY, "color": BLACK}),
                ],
            },
            {
                "label": "3\n사용법",
                "session_message": "웹(claude.ai/code) 권장 — 브라우저 로그인만으로 "
                                   "시작, 결과물은 outputs/ 폴더 (→ 보조 3·4)",
                "content": [
                    (" • 설치 프로그램 0개 — 첫 명령 한 문장으로 5분 안에 동작 "
                     "확인까지 가능합니다",
                     {"size": SZ_BODY, "bold": True, "color": BLUE}),
                    (" • 회의록·주간/월간보고·KPI 조회·기술동향·엑셀/PDF 등 "
                     "복사해 쓰는 요청 문장 치트시트 제공 (→ 보조 5)",
                     {"size": SZ_BODY, "color": BLACK}),
                ],
            },
            {
                "label": "4\n규칙",
                "session_message": "추측 금지 · 출처 명시 · 원본 보존 — 시스템이 "
                                   "자동 검증, 최종 검토 책임은 사람 (→ 보조 6·7)",
                "content": [
                    (" • 출처 인용이 없는 보고서는 자동 차단 — 모든 숫자는 원본 "
                     "대조 후 보고합니다",
                     {"size": SZ_BODY, "bold": True, "color": BLUE}),
                    (" • 사외비·개인정보 입력 전 회사 AI 보안 지침 확인 (현재 "
                     "데이터는 시연용 샘플)",
                     {"size": SZ_BODY, "color": BLACK}),
                ],
            },
        ],
        footnote=("* 출처 : 온보딩_가이드.md (리포 루트, 2026-06-12) · CLAUDE.md "
                  "\"에이전트 5명\"·\"절대 원칙\" · README.md [106~113·149~158·"
                  "186~189행]"),
    )


# ──────────────────── 보조 1/7 — 개념 ────────────────────
def build_aux_concept(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_block(slide, "[보조 1/7] 개념 — 등장인물 3개와 AI 팀원 5명",
                byline=BYLINE)

    y = band_table(
        slide, 1.60, "등장인물 3개 — 이것만 알면 됩니다",
        ["이름", "비유하자면", "하는 일"],
        [
            ("Claude (클로드)", "말귀를 알아듣는 AI 직원",
             "한국어를 이해하고 글·문서·분석을 수행하는 인공지능"),
            ("Claude Code (클로드 코드)", "그 AI 직원의 사무실",
             "팀 폴더(양식·데이터·자료)를 직접 열어 보고 파일을 만들어 주는 작업 환경"),
            ("GitHub (깃허브)", "우리 팀의 공유 캐비닛",
             "양식·데이터·결과물 보관 + 변경 이력 자동 기록 (언제든 되돌리기 가능)"),
        ],
        col_w=[5.6, 6.6, 14.1], row_h=0.80)

    y = band_table(
        slide, y + 0.30, "AI 팀원 5명 — 누구를 부를지 고를 필요 없음 (팀장이 자동 배분)",
        ["팀원", "역할", "담당 업무"],
        [
            ("data-teammate", "데이터 담당",
             "KPI 진척 · Capex(투자비) · MRM(양산 일정) · DST(R&D 진척) 숫자 조회"),
            ("ops-teammate", "운영 담당",
             "협력사 현황 · HR(출장/교육/포상) · 결재 대기 · 사내문서 검색"),
            ("tech-research-teammate", "리서치 담당",
             "외부 기술 동향 · 경쟁사 · 신규 과제 발굴 (AMR/협동로봇/디지털트윈/AI 등)"),
            ("document-writer", "문서 작성 담당",
             "회의록·주간/월간보고 등 markdown 문서 양식 채움"),
            ("ppt-writer", "PPT 담당",
             "사내 PPT 양식 분석·빈칸 채움 (LGES 표준 색·폰트 자동 적용)"),
        ],
        col_w=[6.0, 4.0, 16.3], row_h=0.75)

    band_table(
        slide, y + 0.30, "폴더 4개 — 어디에 무엇이 있나",
        ["폴더", "의미"],
        [
            ("templates/", "양식 보관함 — 회의록·보고서 빈 양식 (원본은 절대 수정되지 않음)"),
            ("sample_data/", "숫자 데이터 — KPI·협력사·투자비 등"),
            ("references/", "참고 자료 — 로드맵·정책·과거 회의록·기술자료 (넣기만 하면 자동 인식)"),
            ("outputs/", "결과물 수령함 — AI가 만든 완성본은 전부 여기 생성"),
        ],
        col_w=[4.6, 21.7], row_h=0.70)

    add_footnote(slide, "* 출처 : 온보딩_가이드.md §1 · CLAUDE.md \"에이전트 5명\"·"
                        "\"폴더 구조\" · .claude/agents/ 5개 정의 파일")


# ──────────────────── 보조 2/7 — 처음 1회 준비 ────────────────────
def build_aux_setup(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_block(slide, "[보조 2/7] 처음 1회 준비 — 계정 2개, 약 30분",
                byline=BYLINE)

    y = band_table(
        slide, 1.60, "준비 절차 — 막히면 혼자 고민하지 말고 관리자에게 요청",
        ["단계", "할 일", "비고"],
        [
            ("① GitHub 가입", "github.com → Sign up → 이메일 인증",
             "사용자 이름은 동료가 알아볼 수 있게 (예: hong-gildong-lges)"),
            ("② 저장소 초대", "관리자에게 GitHub 이름 전달 → 초대 메일에서 "
             "Accept invitation 클릭",
             "github.com/mowonkang/FA-Agent-common 이 열리면 성공"),
            ("③ Claude 가입", "claude.ai → 회사 이메일로 가입",
             "—"),
            ("④ 시트 배정", "Claude Code 권한이 포함된 요금제/시트 확인",
             "회사 계약: 관리자에게 요청 / 개인: Pro 이상"),
        ],
        col_w=[3.6, 11.4, 11.3], row_h=0.95)

    y = band_text(
        slide, y + 0.30, "준비 완료 체크리스트 — 4개 모두 ✓ 면 끝",
        [
            (" □ github.com 에 로그인할 수 있다          □ github.com/mowonkang/"
             "FA-Agent-common 페이지가 열린다",
             {"size": SZ_BODY, "color": BLACK}),
            (" □ claude.ai 에 로그인할 수 있다           □ Claude Code 권한이 "
             "있는 시트를 배정받았다",
             {"size": SZ_BODY, "color": BLACK}),
        ],
        box_h=1.45)

    band_text(
        slide, y + 0.30, "안심 포인트",
        [
            (" • 잘못 눌러도 팀 자료는 망가지지 않습니다 — GitHub 는 모든 변경 "
             "이력을 보관하며 언제든 복구할 수 있습니다",
             {"size": SZ_BODY, "bold": True, "color": BLACK}),
            (" • AI 는 파일을 바꾸기 전에 항상 허락을 구하도록 설정되어 있습니다",
             {"size": SZ_BODY, "color": BLACK}),
            (" • 회사 라이선스 신청 절차·보안 지침·담당자 연락처는 가이드 문서의 "
             "기입란에 추가 예정 (자료 미확인)",
             {"size": SZ_SUB, "color": MID_GRAY}),
        ],
        box_h=2.2)

    add_footnote(slide, "* 출처 : 온보딩_가이드.md §2 (계정 준비 절차·체크리스트)")


# ──────────────── 보조 3/7 — 사용 방법 선택 ────────────────
def build_aux_methods(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_block(slide, "[보조 3/7] 사용 방법 선택 — 웹(권장) vs PC 설치",
                byline=BYLINE)

    y = band_table(
        slide, 1.60, "두 가지 길 — 처음이라면 방법 A(웹) 권장",
        ["항목", "방법 A : 웹에서 사용 (권장)", "방법 B : 내 PC에 설치"],
        [
            ("설치", "없음 (브라우저만)", "Claude Code 프로그램 + 팀 폴더 내려받기"),
            ("난이도", "쉬움", "중간 (안내대로 하면 가능)"),
            ("결과물 받기", "화면에서 다운로드 / GitHub 에 저장",
             "내 PC 폴더(outputs/)에서 바로 열기"),
            ("사용 장소", "PC·노트북·휴대폰(Claude 앱)", "설치한 PC에서만"),
            ("추천 대상", "모든 구성원, 특히 첫 사용자", "매일 쓰는 파워유저·관리자"),
        ],
        col_w=[4.0, 11.4, 10.9], row_h=0.72)

    y = band_text(
        slide, y + 0.30, "방법 A 시작 4단계 — 설치 없음",
        [
            (" ① claude.ai/code 접속·로그인   ② GitHub 계정 연결 승인 (최초 1회) "
             "  ③ 저장소 FA-Agent-common 선택   ④ 한국어로 업무 지시",
             {"size": SZ_BODY, "bold": True, "color": BLACK}),
            (" • \"Claude 가 GitHub 에 접근하도록 허용할까요?\" 화면은 정상 절차 "
             "— 승인(Authorize)하면 됩니다",
             {"size": SZ_SUB, "color": MID_GRAY}),
            (" • 휴대폰 Claude 앱에서도 같은 세션 사용 가능 — 출장 중 \"주간보고 "
             "초안 잡아줘\" 가능",
             {"size": SZ_SUB, "color": MID_GRAY}),
        ],
        box_h=2.15)

    band_text(
        slide, y + 0.30, "방법 B 요약 — 파워유저·관리자",
        [
            (" ① 설치 명령 1줄 : Windows PowerShell 「irm https://claude.ai/"
             "install.ps1 | iex」 · Mac 터미널 「curl -fsSL https://claude.ai/"
             "install.sh | bash」",
             {"size": SZ_BODY, "color": BLACK}),
            (" ② 팀 폴더 내려받기 : GitHub Desktop(클릭만으로) 또는 git clone "
             "  ③ Python 3 + pip install -r requirements.txt (PPT 기능)",
             {"size": SZ_BODY, "color": BLACK}),
            (" ④ 폴더에서 claude 실행 → 브라우저 로그인 → 사용 시작 (최신 안내: "
             "code.claude.com/docs)",
             {"size": SZ_BODY, "color": BLACK}),
        ],
        box_h=2.3)

    add_footnote(slide, "* 출처 : 온보딩_가이드.md §3 · README.md [149~158행] "
                        "(사전 설치 요구사항) · 설치 절차는 공식 문서 기준, 화면은 "
                        "업데이트로 달라질 수 있음")


# ──────────────── 보조 4/7 — 첫 사용 5분 ────────────────
def build_aux_first_use(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_block(slide, "[보조 4/7] 첫 사용 5분 — 동작 확인과 화면 대응",
                byline=BYLINE)

    y = band_text(
        slide, 1.60, "첫 명령 — 그대로 복사해 붙여넣으세요 (공식 동작 확인 명령)",
        [
            (" \"sample_data/kpi_샘플.md의 Q3 KPI를 회의록 양식으로 정리해줘. "
             "일시는 오늘, 참석자는 테스트 사용자.\"",
             {"size": SZ_BODY, "bold": True, "color": CHARCOAL}),
            (" → 잠시 후 outputs/회의록_<오늘날짜>.md 파일이 생기면 모든 것이 "
             "정상 동작하는 것입니다",
             {"size": SZ_SUB, "color": MID_GRAY}),
        ],
        box_h=1.45)

    y = band_table(
        slide, y + 0.30, "사용 중 화면에서 만나는 것들",
        ["상황", "의미", "어떻게 하나요"],
        [
            ("\"~를 실행해도 될까요?\" 권한 질문",
             "파일 생성·프로그램 실행 전 허락을 구하는 것",
             "내용 읽고 승인 — 이상하면 거부 후 \"왜 필요한지 설명해줘\""),
            ("계획(Plan)을 보여주며 승인 요청",
             "쓰기 전에 \"이렇게 채우겠습니다\" 검사받는 것",
             "방향이 맞으면 승인, 아니면 \"○○는 빼고 △△ 강조해줘\""),
            ("AI 팀원들이 일하는 표시",
             "팀장이 데이터/작성 팀원에게 업무 배분 중",
             "기다리면 됩니다 (보통 수 분 이내)"),
            ("\"자료 미확인\" 표기",
             "근거 자료를 못 찾았다는 정직한 보고 (지어내지 않음)",
             "자료 위치를 알려주거나, 그대로 두고 사람이 채움"),
        ],
        col_w=[7.0, 9.2, 10.1], row_h=0.95)

    band_text(
        slide, y + 0.30, "결과물 받기 · 대화 이어가기",
        [
            (" • 결과물은 항상 outputs/ 폴더 — 파일명 「양식이름_YYYY-MM-DD」, "
             "같은 날 다시 만들면 _v2·_v3 자동 부여",
             {"size": SZ_BODY, "bold": True, "color": BLACK}),
            (" • 웹 : \"결과 파일 보내줘\"(다운로드) · \"GitHub에 저장해줘\"(팀 "
             "보관)   /   PC : 탐색기에서 outputs/ 폴더 열기",
             {"size": SZ_BODY, "color": BLACK}),
            (" • 수정은 그냥 말하기 (\"두 번째 섹션 3줄로 줄여줘\") · 중지 Esc · "
             "이어가기는 세션 목록 또는 claude --resume",
             {"size": SZ_BODY, "color": BLACK}),
        ],
        box_h=2.2)

    add_footnote(slide, "* 출처 : 온보딩_가이드.md §4 · README.md [106~113행] "
                        "(동작 확인 명령) · [186~189행] (결과물 위치 규칙)")


# ──────────────── 보조 5/7 — 요청 문장 치트시트 ────────────────
def build_aux_cheatsheet(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_block(slide, "[보조 5/7] 업무별 요청 문장 치트시트 — 복사해서 쓰세요",
                byline=BYLINE)

    y = band_table(
        slide, 1.60, "날짜·이름만 바꿔 그대로 쓰면 됩니다 (따옴표·특수기호 불필요)",
        ["분야", "하고 싶은 일", "이렇게 말하세요"],
        [
            ("회의록", "녹취록 → 회의록 PPT",
             "\"이 녹취록으로 회의록 PPT 만들어줘\" (녹취록 첨부/붙여넣기)"),
            ("회의록", "녹취록 + OneNote 통합",
             "\"녹취록이랑 OneNote 메모 합쳐서 Word 회의록으로 정리해줘\""),
            ("정기 보고", "주간·월간보고",
             "\"이번 주 주간보고 작성해줘\" · \"5월 월간보고 작성해줘. KPI 중심으로\""),
            ("정기 보고", "임원 보고 PPT",
             "\"○○ 건으로 임원 보고용 PPT 만들어줘. 메인 1장 + 보조 2장으로\""),
            ("데이터 조회", "KPI·투자비·양산 일정",
             "\"KPI 진척 현황 알려줘\" · \"Capex 집행 현황 정리해줘\" · \"MRM 일정 요약해줘\""),
            ("데이터 조회", "협력사·결재",
             "\"협력사 현황이랑 이슈 정리해줘\" · \"내 결재 대기 목록 보여줘\""),
            ("기술 동향", "종합 브리핑",
             "\"FA 기술 동향 조사해서 신규 과제 후보 뽑아줘\""),
            ("기술 동향", "분야 지정",
             "\"AMR 최신 동향 조사해줘\" · \"협동로봇 트렌드 정리해줘\""),
            ("파일 작업", "엑셀",
             "\"이 데이터 엑셀로 정리해줘\" · \"KPI 트래커 시트 만들어줘\""),
            ("파일 작업", "PDF·Word",
             "\"이 PDF에서 표만 뽑아서 엑셀로 만들어줘\" · \"이 내용 목차 포함 Word로 만들어줘\""),
            ("파일 작업", "사내 공지",
             "\"이 내용으로 사내 공지문 초안 써줘\""),
        ],
        col_w=[3.0, 5.6, 17.7], row_h=0.78)

    band_text(
        slide, y + 0.30, "파일 주는 법 · 다시 시키는 법",
        [
            (" • 파일 주기 : 웹은 채팅창에 끌어다 놓기(드래그&드롭) / PC 는 "
             "폴더에 넣고 경로를 말하거나 끌어다 놓기",
             {"size": SZ_BODY, "color": BLACK}),
            (" • \"다시\"·\"더 짧게\"·\"어조를 공손하게\" — 몇 번을 고쳐 시켜도 "
             "AI 는 지치지 않습니다",
             {"size": SZ_BODY, "color": BLACK}),
            (" • 모르면 AI 에게 직접 물어보세요 : \"네가 뭘 할 수 있는지 알려줘\" "
             "· \"이 저장소에 어떤 양식이 있어?\"",
             {"size": SZ_BODY, "color": BLACK}),
        ],
        box_h=2.2)

    add_footnote(slide, "* 출처 : 온보딩_가이드.md §5 · CLAUDE.md \"스킬\" 절 "
                        "(.claude/skills/ 14종 기준 요청 예시)")


# ──────────────── 보조 6/7 — 요령·규칙·보안 ────────────────
def build_aux_rules(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_block(slide, "[보조 6/7] 잘 쓰는 요령과 팀 규칙·보안", byline=BYLINE)

    y = band_table(
        slide, 1.60, "잘 쓰는 요령 — 신입사원에게 일을 시키듯 구체적으로",
        ["원칙", "아쉬운 요청", "좋은 요청"],
        [
            ("목적을 말하기", "\"보고서 써줘\"",
             "\"다음주 월요일 그룹장 보고용 주간보고 써줘\""),
            ("기준·범위 주기", "\"KPI 알려줘\"",
             "\"Q3 KPI 진척률을 목표 대비로 알려줘\""),
            ("아는 값은 직접 주기", "(작성자를 비워둠)",
             "\"작성자는 김OO, 일시는 6/12 14시로 넣어줘\""),
            ("형식을 지정하기", "\"정리해줘\"",
             "\"표 1개 + 요약 3줄로 정리해줘\""),
            ("한 번에 한 묶음", "세 가지 일을 한 문장에",
             "회의록 먼저 → 끝나면 다음 요청"),
        ],
        col_w=[4.6, 7.2, 14.5], row_h=0.72)

    y = band_text(
        slide, y + 0.30, "팀 절대 원칙 4 — 시스템이 자동으로 강제합니다",
        [
            (" ① 추측 금지 — 자료가 없으면 지어내는 대신 \"자료 미확인\"으로 표기",
             {"size": SZ_BODY, "bold": True, "color": BLACK}),
            (" ② 출처 명시 — 답변에 (출처: 파일명 [행 번호]) 부착, 출처 없는 "
             "보고서는 품질 게이트가 완료를 자동 차단",
             {"size": SZ_BODY, "bold": True, "color": BLACK}),
            (" ③ 원본 보존 — templates/(양식)·references/(자료) 불변, 결과물은 "
             "outputs/ 에만 생성",
             {"size": SZ_BODY, "bold": True, "color": BLACK}),
            (" ④ 최종 책임은 사람 — AI 산출물은 초안, 결재·보고 전 원본 대조와 "
             "검토는 작성자의 몫",
             {"size": SZ_BODY, "bold": True, "color": BLACK}),
        ],
        box_h=2.7)

    band_text(
        slide, y + 0.30, "보안 주의",
        [
            (" • 현재 sample_data/ 는 시연용 샘플 — 실제 사외비 데이터 반입 전 "
             "보안/IT 검토 필수",
             {"size": SZ_BODY, "color": BLACK}),
            (" • 사외비 문서·개인정보를 붙여넣기 전 회사 AI 사용 지침 확인 "
             "(지침 링크는 가이드 기입란에 추가 예정 — 자료 미확인)",
             {"size": SZ_BODY, "color": BLACK}),
            (" • 외부 서비스 연동(스크래핑 키 등)은 보안/IT 검토 후에만 활성화",
             {"size": SZ_BODY, "color": BLACK}),
        ],
        box_h=2.2)

    add_footnote(slide, "* 출처 : 온보딩_가이드.md §6~7 · CLAUDE.md \"절대 원칙\" "
                        "· scripts/hooks/check_output_citations.py [1~16행] (인용 "
                        "자동 검증 규칙)")


# ──────────────── 보조 7/7 — FAQ·오늘 시작하기 ────────────────
def build_aux_faq(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title_block(slide, "[보조 7/7] FAQ·문제 해결 — 그리고 오늘 시작하기",
                byline=BYLINE)

    y = band_table(
        slide, 1.60, "자주 묻는 질문 — 대부분 한 문장이면 해결됩니다",
        ["질문 / 증상", "해결"],
        [
            ("영어로 써야 하나요? 명령어 외워야 하나요?",
             "아니요 — 전부 한국어 평소 말투, 외울 명령어 없음"),
            ("\"작업 완료 차단 (missing_citations)\" 표시",
             "출처 인용 누락 — \"누락된 출처를 추가해줘\" 한마디면 해결"),
            ("이어간 세션에서 AI 팀원이 사라짐",
             "\"팀원을 다시 생성해줘\" — 같은 구성으로 재생성 (알려진 제한)"),
            ("PPT 일부 칸이 빈 채로 나옴",
             "근거 없는 칸은 추측 대신 비워둠(unfilled) — 값을 알려주고 다시 시키기"),
            ("새 참고자료를 등록하고 싶음",
             "references/ 폴더에 넣기만 하면 자동 인식 — \"이 파일 등록해줘\"도 가능"),
            ("응답이 멈췄거나 이상함",
             "Esc(정지) 후 재지시 · 새 세션에서 \"아까 하던 ○○ 이어서 해줘\""),
        ],
        col_w=[10.6, 15.7], row_h=0.78)

    y = band_text(
        slide, y + 0.30, "오늘 시작하기 — 3단계",
        [
            (" ① claude.ai/code 로그인 → FA-Agent-common 선택     ② 보조 4 의 첫 "
             "명령 복사·붙여넣기     ③ outputs/ 에서 결과 확인",
             {"size": SZ_BODY, "bold": True, "color": BLACK}),
            (" • 이후에는 치트시트(보조 5)의 문장으로 실제 업무에 바로 적용",
             {"size": SZ_BODY, "color": BLACK}),
        ],
        box_h=1.45)

    band_text(
        slide, y + 0.30, "도움이 필요하면",
        [
            (" • 관리자 문의 (담당자 연락처는 가이드 기입란에 추가 예정 — 자료 "
             "미확인) · 공식 문서(한국어) code.claude.com/docs",
             {"size": SZ_BODY, "color": BLACK}),
            (" • 전체 상세 내용은 저장소의 「온보딩_가이드.md」 — 배포용 Word "
             "사본 outputs/FA-Agent_온보딩_가이드_*.docx",
             {"size": SZ_BODY, "color": BLACK}),
        ],
        box_h=1.45)

    add_footnote(slide, "* 출처 : 온보딩_가이드.md §8~10 · CLAUDE.md \"트러블슈팅\" "
                        "(/resume 팀원 복원·인용 차단·인덱스 갱신)")


def build():
    prs = Presentation()
    prs.slide_width = Cm(SLIDE_W_CM)
    prs.slide_height = Cm(SLIDE_H_CM)
    build_main(prs)
    build_aux_concept(prs)
    build_aux_setup(prs)
    build_aux_methods(prs)
    build_aux_first_use(prs)
    build_aux_cheatsheet(prs)
    build_aux_rules(prs)
    build_aux_faq(prs)
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    if OUT_PATH.exists():
        OUT_PATH.unlink()
    prs.save(OUT_PATH)
    print(f"saved: {OUT_PATH}  (slides: {len(prs.slides)})")


if __name__ == "__main__":
    build()
