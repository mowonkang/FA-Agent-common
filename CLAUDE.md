# 리포지토리 가이드 (Claude Code 세션용)

이 파일은 Claude Code 가 모든 세션 시작 시 자동으로 읽습니다. 폴더 구조와 사용 규약을 짧게 정리합니다.

## 폴더 구조

| 폴더 | 용도 | 형식 |
|---|---|---|
| `templates/` | **보고서 양식** (빈칸을 채워서 결과물 생성) | `.md` (placeholder), `.pptx` |
| `sample_data/` | **수치 데이터** (KPI, 협력사, Capex 등) | `.md`, `.csv`, `.json` |
| `references/` | **참조 자료** (로드맵·정책·과거 회의록·기술자료·용어집) | `.md` (frontmatter 필수) |
| `outputs/` | 결과물 | 자동 생성 |
| `scripts/` | 결정론 처리 스크립트 (PPT 추출/채움, 인덱스 빌드) | `.py` |
| `.claude/agents/` | 4개 서브에이전트 정의 | `.md` |
| `.claude/skills/` | 회의록 등 도메인 스킬 (자동 인식) | `SKILL.md` + assets/scripts |

## 에이전트 5명

- `data-teammate` — KPI·Capex·MRM·DST 데이터 조회 (사내 정량)
- `ops-teammate` — 협력사·HR·결재·사내문서 검색 (사내 운영)
- `tech-research-teammate` — 외부 기술 동향·경쟁사·기술 평가 (사외 리서치)
- `document-writer` — markdown 보고서 (.md 양식 채움)
- `ppt-writer` — PPT 보고서 (.pptx 양식 채움)

모든 에이전트는 작업 전 `references/INDEX.md` 를 확인하여 관련 참조 자료를 인지합니다.

## 팀 운영 패턴 (SOP)

`CLAUDE_CODE_EXPERIMENTAL_AGENT_TEAMS=1` (활성화됨) 환경에서 본 리포의 5-에이전트는 다음 패턴으로 운영합니다. ([Agent Teams 가이드](https://code.claude.com/docs/ko/agent-teams) 모범사례 적용)

1. **분해** — 리더가 요청을 도메인별 작업으로 쪼갠다 (사내 수치 ↔ 사내 운영 ↔ 사외 리서치, 작성 ↔ 검토).
2. **병렬 수집** — 필요한 수집 팀원 (`data-teammate` · `ops-teammate` · `tech-research-teammate`) 을 **동시 스폰**해 출처와 함께 데이터 수집. 호출자는 작업 범위·기준일·필요 출처를 프롬프트에 명시 (리더 대화 기록은 미상속).
3. **위임 작성** — 결과를 모아 `document-writer`(.md) 또는 `ppt-writer`(.pptx) 에 전달, `templates/` 양식의 placeholder 만 치환. writer 는 부족 데이터를 우편함으로 직접 요청해 라운드를 돌릴 수 있다.
4. **품질 게이트** — TaskCompleted hook (`scripts/hooks/check_output_citations.py`) 이 `outputs/*.md` 의 출처 인용을 자동 검증.

### 권장 규모

- **3–5명 팀**, **팀원당 5–6개 작업** (가이드 "적절한 팀 크기").
- 사용 케이스별 권장 구성:
  - **정형 보고서** (주간/월간/회의록) — 3명 (`data` + `ops` + writer)
  - **과제 진척률 심층 보고** — 3명 (`data` + `ops` + writer)
  - **신규 과제 브레인스토밍** — 4명 (`tech-research` + `data` + `ops` + writer)
  - **경쟁사 분석 → KPI 재설정** — 3~4명 (`tech-research` + `data` (+ `ops` 선택) + writer)
- `ppt-writer` 는 .pptx 산출물 필요 시 추가.

### Plan-mode 권장

- `document-writer` 와 `ppt-writer` 는 **`outputs/` 에 쓰기 전 plan-mode** 로 시작하는 것을 권장. 리더가 placeholder 매핑·인용 출처를 승인한 후 구현 진입. (가이드 "팀원을 위한 계획 승인 요구")

  ```text
  Spawn a document-writer teammate to fill templates/주간보고.md.
  Require plan approval before they write to outputs/.
  ```

- 데이터 수집 팀원(`data-teammate`, `ops-teammate`, `tech-research-teammate`)은 기본적으로 읽기 전용이라 plan-mode 불필요. 단 `tech-research-teammate` 가 `references/기술자료/` 또는 `references/경쟁사/` 에 신규 파일을 자동 아카이빙할 때는 plan-mode 권장.

## 스킬

- `.claude/skills/meeting-minutes/` — 녹취록 → LGES 표준 회의록 PPTX 생성
- `.claude/skills/meeting-merger/` — 녹취록 + OneNote 메모 → 통합 Word(.docx) 생성
- `.claude/skills/fa-task-discovery/` — FA 6개 분야(AMR/AGV·협동로봇·디지털트윈·스마트물류·AI·피지컬AI) 동향 리서치 → 과제 후보 docx 보고서 + 카카오톡 요약
- `.claude/skills/xlsx/` — Excel(.xlsx/.csv) 생성·편집·수식·피벗 (KPI/Capex/MRM 트래커, Anthropic 공식 + FA 트리거 추가)
- `.claude/skills/pdf/` — PDF 본문/표/이미지 추출, 병합·분할·OCR·신규 생성 (협력사 사양서·도면, 공식 + FA 트리거)
- `.claude/skills/skill-creator/` — 신규 스킬 scaffold 메타 스킬, eval-viewer 포함 (공식 + FA 트리거)
- `.claude/skills/mcp-builder/` — MCP 서버 설계·구현 가이드, 사내 ERP/MES/PLM 연동 PoC (공식 + FA 트리거)
- `.claude/skills/brand-guidelines/` — **LGES 사내 브랜드 가이드 (자체 작성)** — 색·폰트·풋노트·레이아웃 룰을 SKILL 로 캡슐화 (`templates/사내양식/LGES_PPT_작업_가이드.md` 기반)
- `.claude/skills/docx/` — Word 문서 생성·편집·합치기·트래킹 체인지 (정책·SOP·협력사 평가서, 공식 + FA 트리거)
- `.claude/skills/doc-coauthoring/` — 장문 문서 협업 워크플로 (제안서·기술 스펙·임원 결재 문서, 공식 + FA 트리거)
- `.claude/skills/internal-comms/` — 사내 공지·주간 메일·3P 업데이트·FAQ·이슈 리포트 (공식 + FA 트리거)
- `.claude/skills/webapp-testing/` — Playwright 기반 웹 앱 테스트 (장기 검토군, 사내 대시보드 개발 시 활성화)
- `.claude/skills/firecrawl/` — Firecrawl SaaS 안정 스크래핑 (PoC 단계, 보안/IT 검토 후 키 활성화)
- `.claude/skills/sentry-code-review/` — 운영 인식 코드 리뷰 (Python 빌더 품질 게이트, Sentry SDK 선택적)

`document-writer` 가 녹취록 입력을 감지하면 회의록 스킬로 위임합니다. "FA 기술 동향 조사", "신규 과제 발굴", "FA 브리핑" 같은 요청에는 `fa-task-discovery` 스킬을 사용합니다. PPT/DOCX/MD 산출물 생성 시 `brand-guidelines` 스킬이 자동 적용됩니다 (LGES 사내 표준 색·폰트·풋노트). 신규 `scripts/build_*.py` 빌더 추가 시 `sentry-code-review` 스킬로 품질 게이트 적용 권장.

## references/ 사용 규약

- 6개 카테고리: `roadmap/`, `policy/`, `과거회의록/`, `기술자료/`, `경쟁사/`, `용어집/`. 카테고리가 애매하면 `references/` 직속에 두면 자동으로 "기타" 분류됨.
- **자료는 그냥 폴더에 두면 자동 인식됩니다.** 제목·카테고리·날짜·요약은 본문 H1·파일명·mtime 으로 휴리스틱 추출. frontmatter 는 선택 사항(고급).
- 인덱스 갱신:

  ```bash
  python3 scripts/build_reference_index.py
  ```

  Claude Code 세션 시작 시 `.claude/settings.json` 의 SessionStart hook 으로 자동 실행됩니다.
- 인용 형식: `references/<카테고리>/<파일명> [N행]`

## PPT 작업 기본 룰

`.pptx` 산출물을 만드는 모든 에이전트(특히 `ppt-writer`) 는 **`templates/사내양식/LGES_PPT_작업_가이드.md` 를 반드시 따른다.**

핵심 요약:
- **구조 = 메인 1~2장(임원 요약) + 보조 N장(상세).** 메인은 박스 안에 모두 들어가야 하고, 넘치는 내용은 보조로 분리
- **기본색 = 화이트 + 검정 + 회색 계열** (DIM_GRAY 라벨 띠, MID_GRAY 카드 보더, GRAY 표 헤더)
- **강조색 = 파랑 `#0000FF` (R0 G0 B255)** — 메인에서 특이 포인트에만 **3~5개 이내**, 보조는 가능한 회색
- **풋노트 = 별도 텍스트박스 + 초록곰팡이 `#006600` (G102) Bold 8pt**
- 폰트: 한글 `LG스마트체 Regular` + 영문 `Arial Narrow`, run-level 자동 분리
- 폰트 사이즈 6단 상수 `SZ_TITLE 16 / SZ_BAND 12 / SZ_SECTION 11 / SZ_BODY 10 / SZ_SUB 9 / SZ_FOOT 8` 만 사용

### 표준 레이아웃 = 분류형 (사내 보고양식 슬라이드 1)

보고용 `.pptx` 의 **표준 메인 레이아웃은 "분류형"** 이다 (기존 2×2 사분면은 가독성 이슈로 대체).
구조 = 제목(상단) → **배경 없는 검정 서술형 Head message(15pt, 1~2줄)** → **좌측 분류 레일
(분류 1~N, 높이 자동 가변)** → 우측 세션메시지+내용(표/이미지 선택) → 초록 풋노트. 우상단 조직은
**검정 + `(YYYY.MM.DD)`**. 구성은 **메인(분류형 요약) + 보조 N장(표 중심 심층)**.

- **반드시** 재사용 헬퍼 `scripts/_ppt_layout.py` 의 `render_classified_slide()` 로 메인을 그린다
  (드로잉 프리미티브·색·폰트·`SZ_*` 상수도 본 모듈 import — 빌더별 중복 정의 금지).
- 좌측 분류 박스 높이는 우측 내용에서 자동 산정, 표 폰트 기본 `SZ_BODY`(10pt), `fill_to_bottom` 으로 하단 채움.
- **상세 표준·작성 워크플로우(데이터 수집→메인→보조→프로 엔지니어 톤→방향성 확인→검증)** :
  `references/policy/분류형_보고_레이아웃_표준.md`.
- 기준 구현: `scripts/build_classified_layout_demo.py` · `build_flexibility_definition.py` ·
  `build_flexibility_competitor_compare.py`.

기준 구현(레거시 사분면): `scripts/build_flexibility_summary.py` (산출물: `outputs/유연성지표_요약_2026-05-14.pptx`).

상위(임원) 보고 요청 시 진입점: `templates/사내양식/보고양식_카탈로그.md`
(+ `.json`) — 12종 양식을 8개 보고 블록으로 정리. 양식·톤은 **참고**이며
가독성·스토리 위해 자율 조정 가능하나, **텍스트 무오버플로·임원 즉시 활용
품질·스토리는 필수**. 적용 예시: `scripts/build_fa_exec_report.py` →
`outputs/FA자동화_임원보고_예시_2026-05-18.pptx`.

## 절대 원칙

- **추측 금지** — 자료가 없으면 "찾지 못했습니다"라고 명시
- **출처 명시** — 파일명·행 번호 함께 적기
- **양식 ≠ 참조 자료** — `templates/` 와 `references/` 를 혼동하지 않기
- **`outputs/`** 만 결과물, 다른 폴더 원본은 보존

## 트러블슈팅

- **`/resume` 후 팀원이 사라짐** — Agent Teams 가이드의 알려진 제한: in-process 팀원은 세션 재개 시 복원되지 않습니다. 리더에게 "팀원을 다시 생성해줘" 라고 지시해 동일 구성으로 재스폰하세요.
- **TaskCompleted hook 이 차단함** — `outputs/*.md` 에 `(출처: ...)` 또는 `[N행]` 또는 `자료 미확인` 이 1회도 없으면 차단됩니다. 누락 항목에 인용을 추가하거나 명시적으로 `자료 미확인` 으로 표기.
- **인덱스가 오래되어 보임** — `python3 scripts/build_reference_index.py --quiet` 로 수동 갱신 (세션 시작 시 자동 갱신됨).
