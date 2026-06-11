---
name: firecrawl
description: Firecrawl SaaS 를 활용한 안정적인 웹 스크래핑·구조화 추출·딥 리서치 스킬. JavaScript 가 많은 페이지(Reuters, Bloomberg, Bloomberg, CATL IR, 경쟁사 IR 페이지) 같은 사이트에서 WebFetch 가 제대로 추출하지 못할 때 사용한다. LGES FA 자동화 프로젝트의 tech-research-teammate 가 외부 동향·경쟁사 정보를 수집할 때 보조적으로 호출. 사용자가 '경쟁사 IR 수집해줘', 'JS 사이트 스크래핑', 'CATL 뉴스 가져와', 'Firecrawl 로 긁어와', '안정 크롤링', '대량 페이지 구조화 추출' 같이 말하면 트리거. 외부 SaaS API 키 필요 → 보안/IT 검토 후 도입. 본 스킬은 도입·키 설정 가이드와 호출 패턴을 정의하며, 실제 운영은 별도 PoC 단계 거친 뒤 활성화.
license: 자체 작성 (커뮤니티 인기 도구 Firecrawl 의 LGES FA 프로젝트 적용 가이드)
---

# Firecrawl 활용 스킬 (LGES FA)

> **상태**: 본 스킬은 **검토 추천군** — 실제 SaaS 키 설정 전까지 가이드 문서로만 동작. PoC 단계에 키 발급·환경변수 등록 후 활성화.

## 1. 사용 시점

`tech-research-teammate` 의 기본 WebFetch/WebSearch 로 다음과 같은 문제가 발생할 때:

- JS 렌더링이 무거운 사이트 (CATL IR, Panasonic IR, Tesla 5K, Bloomberg, FT 등)
- 본문 추출이 잘 안 되는 컨퍼런스/연차보고서 페이지
- 대량 페이지를 구조화된 JSON 으로 일괄 수집해야 할 때
- 동적 컨텐츠 (로그인 후 페이지 변경) 가 있는 페이지

본 리포에서는 **외부 SaaS 키가 활성화되기 전까지** 본 스킬을 활용하지 말고, WebFetch/WebSearch 우선 시도.

## 2. 설치·키 설정 (PoC 진입 시)

```bash
pip install firecrawl-py
export FIRECRAWL_API_KEY="<발급받은 키>"   # 보안실 검토 완료 후
```

키는 `.env` 또는 `.claude/settings.local.json` 에 저장하고 git 에 커밋하지 않는다. 본 리포의 `.gitignore` 룰을 따른다.

## 3. 호출 패턴

### Scrape (단일 페이지 본문 + 메타 추출)

```python
from firecrawl import FirecrawlApp
app = FirecrawlApp(api_key=os.environ["FIRECRAWL_API_KEY"])
result = app.scrape_url("https://www.catl.com/en/news/", params={"formats": ["markdown", "html"]})
print(result["markdown"])
```

### Crawl (사이트 전체 크롤)

```python
crawl_id = app.crawl_url(
    "https://example.com/news",
    params={"limit": 50, "scrapeOptions": {"formats": ["markdown"]}}
)
```

### Map (사이트 URL 목록만 빠르게 수집)

```python
urls = app.map_url("https://www.catl.com/")
```

## 4. 결합 자산

- **tech-research-teammate**: WebFetch/WebSearch 실패 시 폴백
- **references/경쟁사/**: Firecrawl 로 수집한 페이지의 markdown 본문 저장 → `scripts/build_reference_index.py` 자동 색인
- **fa-task-discovery**: 6개 도메인 모니터링에서 JS 사이트 폴백

## 5. 운영 룰

- **출처 명시 필수**: 수집한 본문을 `references/경쟁사/` 에 저장 시 frontmatter 또는 본문에 `(출처: <원문 URL>, 수집일: YYYY-MM-DD, Firecrawl)` 명시
- **rate limit 준수**: 키 별 분당 호출 수 모니터링
- **저작권**: 본 리포에 영구 저장 전 외부 사이트 ToS 검토 — 사내 활용 한정 추출만 진행

## 6. 출처

- 본 스킬 가이드: 본 리포의 `outputs/popular_skills_review_2026-05-20.md` §5 #9 항목
- Firecrawl 공식 문서: `firecrawl.dev` (PoC 진입 시점에 최신 SDK 확인)
- 커뮤니티 추천 자료: `firecrawl.dev/blog/best-claude-code-skills` 2026 May
