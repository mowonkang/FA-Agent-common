# claude_FA-Agent_1
Claude code FA기술혁신 Part Agent

## 구성

- 서브에이전트
  - `data-teammate` — KPI · Capex · MRM(양산) · DST(R&D) 데이터 조회
  - `ops-teammate` — 협력사 · HR · 결재 · 사내문서 검색
  - `document-writer` — markdown(.md) 양식의 빈칸을 채워 보고서 작성
  - `ppt-writer` — PPT(.pptx) 양식을 분석해 빈칸을 추론·채워 PPTX 생성
- 양식: `templates/`
- 데이터: `sample_data/`
- **참조 자료**: `references/` (로드맵·정책·과거 회의록·기술자료·용어집) — [`references/README.md`](references/README.md)
- 결과물: `outputs/`
- 캐시(추출 JSON, 값 JSON): `outputs/.cache/` (gitignore)

## 에이전트 팀 구성

> **문서장 + 데이터 팀원 + 운영 팀원 + 문서 팀원** 4명이 협업해서 회의록·주간보고·월간보고를 자동으로 만들어 주는 Claude Code 셋업입니다. 별도 서버·DB·코드 없이 마크다운 파일 몇 개로만 동작합니다.

이 레포에는 **PPT 팀원**(`ppt-writer`) 이 추가돼 총 4명의 서브에이전트 + 메인 세션(팀장) 구조입니다.

- 기반: **Claude Code Agent Teams** 공식 기능 (2026.02 릴리즈, 실험적)
- 필수 설정: `.claude/settings.json` 의 `CLAUDE_CODE_EXPERIMENTAL_AGENT_TEAMS=1` — 변경 후 Claude Code 재시작 필요
- 요구사항: Claude Code v2.1.32+

### Agent Teams 공식 용어 ↔ 본 레포 매핑

영상 ["클로드 코드 Agent Teams 완벽 정리 | Subagent와 차이점"](https://youtu.be/qGm8odiBkBg) (짐코딩, 2026-02-28) 의 용어를 기준으로:

| 영상 용어 | 본 레포 매핑 | 상태 |
|---|---|---|
| Team Lead (팀장) | 메인 Claude Code 세션 (사용자가 직접 대화) | 암묵적 |
| Teammate (팀원) | `.claude/agents/*.md` 4개 | ✅ |
| Spawn (팀원 세션 생성) | `CLAUDE_CODE_EXPERIMENTAL_AGENT_TEAMS=1` 환경변수로 자동 | ✅ |
| Mailbox (메시지 전송) | 요청 `to:/subject:/payload:` · 응답 `key:/value:/source:` 양방향 표준화 | ✅ |
| 직접 소통 (팀장 거치지 않음) | writer 가 data/ops-teammate 에게 직접 요청 | ✅ |
| 컨텍스트 유지 (독립 세션) | Agent Teams 기능이 자동 처리 | ⚪ |
| 공유 Task List | (미적용) writer 가 일회성 dispatcher 가 되는 패턴 | ❌ |
| 비용 (독립 세션 토큰 증가) | 동시 spawn 2-3 명 수준이라 영향 제한적 | ⚪ |

### 한눈 요약 표

| 에이전트 | 역할 | 담당 영역 | tools | 협업 채널 |
|---|---|---|---|---|
| `data-teammate` | 데이터 조회 | KPI·Capex·MRM(양산)·DST(R&D) | Read, Grep, Glob, Bash | document-writer / ppt-writer 의 요청 처리 |
| `ops-teammate` | 운영 조회 | 협력사·HR·결재 큐·사내문서 | Read, Grep, Glob, Bash | document-writer / ppt-writer 의 요청 처리 |
| `document-writer` | markdown 보고서 | 회의록·주간보고·월간보고 (.md 채움) | Read, Write, Edit, Glob, Grep | data/ops-teammate 에게 직접 요청 |
| `ppt-writer` | PPT 보고서 | 주간보고·회의록 (.pptx 채움, 결정론 스크립트 호출) | Read, Write, Edit, Glob, Grep, Bash | data/ops-teammate 에게 직접 요청 |

### 우편함 라우팅

```
사용자
  ↓
팀장(Claude Code 메인) ──── dispatch
                  │
       ┌──────────┴──────────┐
       ↓                     ↓
document-writer (.md)   ppt-writer (.pptx)
       │                     │
       └─────────┬───────────┘  ← 팀장 거치지 않고 직접 요청
                 ↓
       ┌─────────┴─────────┐
       ↓                   ↓
 data-teammate        ops-teammate
       │                   │
       ↓                   ↓
  sample_data/        sample_data/ + 사내문서
```

### 에이전트별 카드

#### `data-teammate`
- **담당**: KPI·Roadmap / Capex·면적·구매 / MRM(양산 로드맵) / DST(R&D 미래기술)
- **원칙**: 추측 금지(실파일 인용) · 출처 명시(`파일명 [N행]`) · 자기 영역만(협력사·HR 은 ops-teammate 에게 위임) · 추이는 Q1→Q2→Q3
- **자료 위치**: `sample_data/`
- **상세**: [`.claude/agents/data-teammate.md`](.claude/agents/data-teammate.md)

#### `ops-teammate`
- **담당**: 협력사 / HR(출장·교육·포상) / 결재 큐 / 사내문서 검색
- **원칙**: 사람 승인 작업은 "제안"만 · 사내문서 결과는 `[파일명 페이지/슬라이드]` 출처 · 자기 영역만(KPI·Capex 는 data-teammate 에게 위임) · 추측 금지
- **자료 위치**: `sample_data/`
- **상세**: [`.claude/agents/ops-teammate.md`](.claude/agents/ops-teammate.md)

#### `document-writer`
- **담당**: markdown 양식 채움 (회의록·주간보고·월간보고)
- **원칙**: 추측 금지(부족값은 팀원에게 요청) · 양식 구조 보존 · 양식에 없는 섹션 추가 금지 · 결과는 `outputs/<양식>_YYYY-MM-DD.md`
- **양식 위치**: `templates/**/*.md`
- **상세**: [`.claude/agents/document-writer.md`](.claude/agents/document-writer.md)

#### `ppt-writer`
- **담당**: PPT(.pptx) 양식 채움. PPTX 바이너리 직접 수정 금지, 항상 `scripts/ppt_extract.py` + `scripts/ppt_fill.py` 호출
- **원칙**: 추측 금지(unfilled 처리) · 원본 보존(`templates/` 덮어쓰기 금지) · 양식 구조 변경 금지 · 자연어 자리표시 매치는 결정적이지 않으므로 모호하면 unfilled
- **양식 위치**: `templates/**/*.pptx`
- **상세**: [`.claude/agents/ppt-writer.md`](.claude/agents/ppt-writer.md)

### 팀 셋업 진화 흐름

- **초기 셋업** (`7c13cd0`, 2026-05-10): 3개 에이전트(data/ops/document-writer) + 3개 markdown 양식 + 2개 sample_data
- **PPT 파이프라인 추가** (`4fed800`, 2026-05-11): `ppt-writer` + `scripts/ppt_extract.py`/`ppt_fill.py` + `templates/주간보고.pptx`
- **자연어 자리표시 + 하위 폴더 양식 지원** (`07d98ee`, 2026-05-11): `templates/**/*.pptx` 재귀 검색, `font_color_rgb` 추출, "이름"/"20YY년" 등 자연어 자리표시 인식

### 동작 확인

원본 setup 프롬프트에서 인용한 첫 명령:

```
sample_data/kpi_샘플.md의 Q3 KPI를 회의록 양식으로 정리해줘.
일시는 오늘, 참석자는 테스트 사용자.
```

`outputs/회의록_<YYYY-MM-DD>.md` 가 생기면 팀 라우팅이 정상 동작하는 것입니다.

> `.claude/settings.json` 의 `CLAUDE_CODE_EXPERIMENTAL_AGENT_TEAMS=1` 가 적용되려면 Claude Code 를 한 번 재시작해야 합니다.

### 다른 레포에 같은 팀을 재현하려면

main 브랜치의 [`previous setup prompt.md`](https://github.com/mowonkang/claude_FA-Agent_1/blob/main/previous%20setup%20prompt.md) 가 배포용 튜토리얼입니다.

- 빈 폴더에서 `claude` 실행 → 그 문서의 "셋업 프롬프트" 섹션을 통째 복사·붙여넣기 → Claude Code 재시작 → 동작 확인 명령 입력 → 3개 에이전트 + 양식 · sample_data 가 자동 생성됨
- 그 위에 **PPT 팀원**을 더하려면 이 레포의 `scripts/ppt_extract.py`, `scripts/ppt_fill.py`, `.claude/agents/ppt-writer.md` 를 그대로 복사

### 양식 / 에이전트 / 참조 자료 추가 가이드

- 새 markdown 양식: `templates/<이름>.md` (placeholder `{{ key }}`) → document-writer 가 자동 인식
- 새 PPT 양식: `templates/<이름>.pptx` 또는 `templates/<폴더>/<이름>.pptx` → ppt-writer 가 자동 인식
- 새 데이터 출처: `sample_data/` 에 .md/.csv/.json 추가 → data/ops-teammate 가 글로브 검색으로 발견
- **새 참조 자료** (로드맵·정책 등): `references/<카테고리>/<파일명>.md` 추가 + frontmatter(title/category/date/tags/summary) 필수 → `python3 scripts/build_reference_index.py` 로 인덱스 갱신 (세션 시작 시 자동 실행). 자세히는 [`references/README.md`](references/README.md)
- 새 에이전트: `.claude/agents/<이름>.md` 추가 (front matter `name`, `description`, `tools` 필수)

### 의도적으로 제외한 패턴

- **공유 Task List**: writer 가 일회성 dispatcher 가 되는 단순한 흐름. 보고서 작성 외 워크플로우(예: 코드 리뷰)가 추가될 때 도입 검토.
- **디스플레이 모드(Split panes)**: tmux/iTerm2 환경 의존이라 셋업 가이드에 두지 않음.

### 출처

- `.claude/agents/{data-teammate,ops-teammate,document-writer,ppt-writer}.md` — 각 에이전트의 행동 원칙·작업 절차 전체 (source of truth)
- 배포용 튜토리얼 + 셋업 프롬프트 전문: main 브랜치 `previous setup prompt.md`
- 영상: [클로드 코드 Agent Teams 완벽 정리 | Subagent와 차이점](https://youtu.be/qGm8odiBkBg) (짐코딩, 2026-02-28)
- git log: `7c13cd0` (초기 팀 셋업) · `4fed800` (ppt-writer 추가) · `07d98ee` (자연어 자리표시 보강)

## PPT 양식 자동 채움

회사 양식이 PPT(.pptx) 인 경우, `ppt-writer` 에이전트가 양식을 분석해 빈칸을 추론하고
다른 팀원(`data-teammate`, `ops-teammate`)에게 데이터를 받아 채운 PPTX 를 만들어 줍니다.

### 1. 사전 설치

**사전 요구사항**
- Claude Code v2.1.32+ (`claude --version` 로 확인)
- `.claude/settings.json` 의 환경변수 `CLAUDE_CODE_EXPERIMENTAL_AGENT_TEAMS=1` (이미 설정됨). 처음 클론한 직후라면 Claude Code 를 한 번 재시작해야 적용됩니다.
- PPT 기능을 쓰려면 아래 `pip install` 로 `python-pptx` 까지 설치.

```
pip install -r requirements.txt
```

### 2. 양식 업로드

1. 채우고 싶은 PPT 양식을 `templates/` 폴더 또는 그 **하위 폴더** 에 둡니다.
   - 직접 아래: `templates/주간보고.pptx`
   - 하위 폴더: `templates/사내양식/회의록.pptx`, `templates/외부양식/제안서.pptx`
   - ppt-writer 는 `templates/**/<이름>.pptx` 로 재귀 검색합니다.
2. Placeholder 표시는 **선택사항** 입니다.
   - 명시적으로 표시하려면 텍스트박스/셀에 `{{ key_name }}` 형태로 적어두면 됩니다.
   - 표시하지 않아도, Claude 가 다음 단서로 빈칸을 추론합니다:
     - 슬라이드 구조·도형 이름·layout name
     - 표 헤더 + 빈 데이터 셀
     - 라벨 도형("작성자:") 인접 빈 도형
     - **자연어 자리표시 텍스트 패턴** (회사 표준 양식 지원): `이름`, `직책or소속`, `장소`, `회의실`, `20YY년`, `MM월DD일`, `xx:00~xx:00`, `YY.MMDD`, `YY.MM`, `요약 내용 N`, `안건 N`, `Comment or YY.MM`, `주관 부서(cf. ...)`, `*회사or법인(cf. ...)`, `*조직(cf. ...)`
     - **색상 강조** (font color RGB): 노랑 `FFFF00`/`FFFF99`, 초록 `00B050`/`92D050`, 주황 `FFC000`, 빨강 `FF0000`, 하늘 `00B0F0` — 회사 양식이 강조색으로 "여기 채우라"를 표시한 경우 자동 인식
3. 같은 이름의 markdown 양식이 있으면 (`templates/**/<이름>.md`) 호출 시 어떤 형식인지 되묻습니다.
4. 양식 폴더에 임시 파일·백업본을 두지 마세요. 재귀 글로브가 같은 이름의 다른 파일을 잡으면 호출 시 어느 경로인지 되묻습니다.

### 3. 호출 예시

- "주간보고 PPT 작성해줘"
- "templates/월간보고.pptx 채워줘"
- "회의록 PPT, 작성자 김OO, 일시 2026-05-11 14:00 으로 만들어줘"

사용자가 메시지에 직접 준 값(작성자·일시 등) 은 그대로 들어가고, 나머지는 ppt-writer 가
data-teammate / ops-teammate 에게 받아서 채웁니다.

### 4. 결과물 위치

- `outputs/<양식이름>_<YYYY-MM-DD>.pptx`
- 동일 날짜에 다시 실행하면 `_v2`, `_v3` suffix 가 자동으로 붙습니다.

### 5. 어떤 항목은 채워지지 않을 수 있습니다

- SmartArt, 이미지 내 텍스트, 차트 내부 데이터는 자동 채움 대상이 아닙니다.
- 빈 도형의 의미를 추론할 단서가 없으면 비워둡니다 (라벨도 placeholder 도 자리표시 패턴도 없는 경우).
- 자연어 자리표시 패턴 매치는 결정적이지 않습니다. 같은 단어가 본문의 일반 표현일 수도 있으면 ppt-writer 가 추측 대신 사용자에게 확인하거나 `unfilled` 로 남깁니다.
- 채워지지 않은 항목은 ppt-writer 의 보고 메시지 `unfilled` 섹션에서 확인할 수 있습니다.

### 6. 처리 파이프라인 (내부 동작)

```
templates/<이름>.pptx
   ↓ scripts/ppt_extract.py
outputs/.cache/<이름>.structure.json     ← Claude 가 분석
   ↓ Claude 추론 + 팀원 데이터 요청
outputs/.cache/<이름>.values.json
   ↓ scripts/ppt_fill.py
outputs/<이름>_<YYYY-MM-DD>.pptx          ← 최종 결과
```

Claude 는 PPTX 바이너리를 직접 만지지 않습니다. 결정론적 파이썬 스크립트가 추출/채움을
담당하고, Claude 는 그 사이에서 "어디에 무엇을 채울지" 판단만 합니다. 원본 PPTX 는 절대
덮어쓰지 않습니다.

### 7. 스크립트 직접 사용 (선택)

에이전트를 거치지 않고 스크립트만 직접 쓸 수도 있습니다.

```
# 구조 추출
python scripts/ppt_extract.py templates/주간보고.pptx
# → outputs/.cache/주간보고.structure.json 생성

# 값 JSON 을 수기로 작성한 뒤 채움
python scripts/ppt_fill.py \
  --template templates/주간보고.pptx \
  --values   outputs/.cache/주간보고.values.json \
  --out      outputs/주간보고_$(date +%F).pptx
```

`scripts/build_sample_template.py` 로 검증·시연용 샘플 양식을 다시 만들 수 있습니다
(이미 `templates/주간보고.pptx` 가 있으면 `--force` 필요).

### 8. JSON 스키마 요약

**StructureDoc** (`outputs/.cache/<이름>.structure.json`)
- `slides[].shapes[]`: `shape_id` (`s{slide_idx}_sh{shape_idx}`), `text`, `runs`, `is_placeholder`,
  `is_empty`, `is_table`/`table`, `is_image`, `is_chart`, `is_smartart`, `explicit_placeholders`,
  `highlight_color_rgb`
- `slides[].shapes[].runs[]`: `text`, `font_name`, `font_size_pt`, `bold`, `italic`, `font_color_rgb`
- `slides[].shapes[].table.cells[][]`: `text`, `is_empty`, `explicit_placeholders`, `runs`, `highlight_color_rgb`
- `summary`: `explicit_placeholder_keys`, `shapes_empty`, `tables_total`, `images_total`,
  `charts_total`, `smartart_total`, `emphasis_runs_total`

**FillValuesDoc** (`outputs/.cache/<이름>.values.json`)
- `fills[]`: `target.kind` ∈ {`shape_text`, `table_cell`}, `shape_id`, `value`, `mode`, `source`
- `mode` ∈ {`replace_placeholder`, `replace_text`, `replace_empty`, `set_cell_text`, `append_paragraph`}
- `unfilled[]`: `shape_id`, `reason`, `hint`

### 9. 알려진 한계

- python-pptx 로 SmartArt / 이미지 내 텍스트 / 차트 내부 데이터는 갱신할 수 없습니다 → `unfilled` 처리만.
- `{{ key }}` 가 PPT 의 run 경계로 쪼개진 경우, 채움 시 paragraph 전체 텍스트를 재조립해 첫 run 에 결과를 몰아넣고 나머지 run 은 비웁니다 (스타일은 첫 run 기준 유지).
- 한글 폰트는 환경(Office / LibreOffice) 에 따라 다르게 렌더링될 수 있습니다. 양식 폰트는
  사용자 표준 폰트로 통일·임베드 권장.
