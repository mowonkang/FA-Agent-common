---
name: sentry-code-review
description: 운영(production-aware) 인식 코드 리뷰 스킬. 에러 핸들링·로깅·모니터링 표준·예외 처리·신뢰성 관점에서 Python 스크립트와 신규 PPT/DOCX 빌더 코드를 자동 점검한다. LGES FA 자동화 프로젝트의 `scripts/` 17+개 빌더와 신규 추가되는 빌더(`build_*.py`) 의 품질 게이트 역할. 사용자가 '이 스크립트 리뷰', '코드 검토', '에러 핸들링 점검', '운영 안전성 체크', '로그 누락 확인', 'PR 리뷰' 같이 말하거나 신규 `scripts/build_*.py` 가 추가되었을 때 트리거. 본 리포의 TaskCompleted hook(`scripts/hooks/check_output_citations.py`) 과 상호 보완 — 인용 검증 외 코드 품질 검증을 추가. 외부 Sentry SaaS 가입은 선택적이며, 본 스킬은 Sentry 의 코드 리뷰 규칙을 LGES 환경에 캡슐화한 가이드.
license: 자체 작성 (커뮤니티 인기 도구 Sentry Code Review 가이드의 LGES FA 적용 버전)
---

# Sentry-style 코드 리뷰 스킬 (LGES FA)

> **상태**: 본 스킬은 **검토 추천군** — Sentry SaaS 가입 전이라도 본 SKILL.md 의 리뷰 체크리스트를 가이드로 활용 가능. Sentry 가입 시 SDK 연동으로 자동 리뷰까지 확장.

## 1. 사용 시점

다음 상황에서 본 스킬을 호출하거나, 신규 PR 머지 전 자동 적용:

- 신규 `scripts/build_*.py` 빌더가 추가될 때
- 기존 빌더가 운영 환경에서 에러를 일으킬 가능성 (수식 평가 실패·파일 경로 누락·외부 호출 실패) 이 있을 때
- TaskCompleted hook (인용 검증) 통과했지만 코드 품질이 의심스러울 때
- Python 스크립트의 에러 핸들링·로깅이 부족하다고 판단될 때

## 2. 리뷰 체크리스트 (수동 적용 가능, Sentry 미가입 시)

### 2-1. 에러 처리

- [ ] try/except 가 너무 광범위(`except Exception:`)하지 않은가? 가능한 좁은 예외만
- [ ] `except:` (bare except) 사용 금지
- [ ] 실패 시 사용자에게 actionable 메시지를 출력하는가?
- [ ] 외부 호출(API·파일·shell) 의 timeout 이 설정되었는가?
- [ ] 부분 실패 시 transactional rollback 또는 안전한 cleanup 이 있는가?

### 2-2. 로깅·관찰성

- [ ] 주요 단계마다 `print()` 또는 `logging` 출력이 있는가? (Cron/CI 에서 추적 가능해야)
- [ ] 외부 호출(WebFetch, subprocess) 결과를 어딘가에 기록하는가?
- [ ] 산출물 경로·크기·생성 시간이 출력되는가? (`saved: outputs/...` 패턴)

### 2-3. 신뢰성

- [ ] 동일 입력으로 여러 번 실행해도 동일 결과인가? (idempotency)
- [ ] 외부 의존(인터넷·라이브러리·환경변수) 이 명시되었는가?
- [ ] `python3 scripts/<name>.py` 단독 실행 가능한가? (다른 스크립트 import 의존 없이)
- [ ] 부산물(temp file·캐시) 정리되는가?

### 2-4. 본 리포 컨벤션

- [ ] 출력은 `outputs/` 또는 `references/` 만 (다른 폴더 원본 보존)
- [ ] LGES PPT 가이드 (`brand-guidelines` 스킬) 준수 — 색·폰트·풋노트
- [ ] 출처 인용 룰 (`scripts/hooks/check_output_citations.py`) 정렬
- [ ] 폰트 사이즈 상수 (`SZ_TITLE/SZ_BAND/SZ_SECTION/SZ_BODY/SZ_SUB/SZ_FOOT`) 만 사용

## 3. 자동 리뷰 (Sentry SaaS 가입 후 PoC)

```bash
pip install sentry-sdk
export SENTRY_DSN="<DSN>"   # 보안실 검토 후
```

Python 빌더 진입점에 다음 한 줄 삽입:

```python
import sentry_sdk
sentry_sdk.init(dsn=os.environ.get("SENTRY_DSN"), traces_sample_rate=0.1)
```

이후 운영 환경(cron·CI) 에서 발생하는 모든 예외/에러를 Sentry 대시보드에서 추적.

## 4. 결합 자산

- **scripts/hooks/check_output_citations.py**: 인용 검증 (본 리포 기존 TaskCompleted hook)
- **scripts/build_*.py**: 본 스킬의 점검 대상
- **PR 리뷰 워크플로**: GitHub PR 생성 시 자동 적용 후보 (GitHub Action 등록 PoC)

## 5. 운영 룰

- **DSN 키 보안**: `.env` 또는 환경변수만, git 미커밋
- **SaaS 비용**: Sentry 무료 플랜 한도(5,000 events/월) 내 운영 → 초과 시 IT 협의
- **개인정보**: 사내 데이터·KPI 수치 노출 가능한 traceback 은 scrubbing 처리

## 6. 출처

- 본 스킬 가이드: 본 리포의 `outputs/popular_skills_review_2026-05-20.md` §5 #10 항목
- Sentry 코드 리뷰 모범사례: `nimbleway.com/blog/anthropic-claude-agent-skills` 2026 May
- Sentry 공식: `docs.sentry.io` (PoC 진입 시점에 최신 SDK 확인)
