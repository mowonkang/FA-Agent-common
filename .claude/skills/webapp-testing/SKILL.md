---
name: webapp-testing
description: Playwright 기반 로컬 웹 애플리케이션 자동 테스트 스킬. 프런트엔드 기능 검증, UI 동작 디버깅, 브라우저 스크린샷 캡처, 브라우저 로그 조회. LGES FA 자동화 프로젝트는 현재 사내 웹 대시보드가 없으므로 **장기 검토군**으로 등록. 향후 사내 FA 대시보드(KPI 시각화·실시간 양산 일정·협력사 현황)·React/Vue 기반 임원 대시보드 개발 시 사용한다. 사용자가 '웹 앱 테스트', '대시보드 동작 확인', '브라우저 스크린샷', 'Playwright 테스트 짜줘', 'UI 자동 검증', '프론트엔드 디버깅' 같이 말하면 트리거. 보고서·PPT·DOCX 산출물에는 트리거하지 않음. Toolkit for interacting with and testing local web applications using Playwright. Supports verifying frontend functionality, debugging UI behavior, capturing browser screenshots, and viewing browser logs.
license: Complete terms in LICENSE.txt (Anthropic skills 공식 원본 + FA 트리거 구절 추가)
---

> **LGES FA 프로젝트 적용 노트 (본 리포 한정)**
>
> - **현재 활용도**: 본 리포에 웹 앱 산출물이 없어 **즉시 활용은 제한적** (장기 PoC 대기).
> - **장기 적용 후보**: ① 사내 FA 대시보드 (KPI·MRM·협력사 현황) ② AI 경진대회 데모 웹 ③ Claude Code 기반 사내 도구 UI.
> - **의존성**: `playwright` Python 패키지 + 브라우저(`playwright install`). 본 리포 가상환경에 사전 설치 안 됨 (사용 시점에 설치).

# Web Application Testing

To test local web applications, write native Python Playwright scripts.

**Helper Scripts Available**:
- `scripts/with_server.py` - Manages server lifecycle (supports multiple servers)

**Always run scripts with `--help` first** to see usage. DO NOT read the source until you try running the script first and find that a customized solution is abslutely necessary. These scripts can be very large and thus pollute your context window. They exist to be called directly as black-box scripts rather than ingested into your context window.

## Decision Tree: Choosing Your Approach

```
User task → Is it static HTML?
    ├─ Yes → Read HTML file directly to identify selectors
    │         ├─ Success → Write Playwright script using selectors
    │         └─ Fails/Incomplete → Treat as dynamic (below)
    │
    └─ No (dynamic webapp) → Is the server already running?
        ├─ No → Run: python scripts/with_server.py --help
        │        Then use the helper + write simplified Playwright script
        │
        └─ Yes → Reconnaissance-then-action:
            1. Navigate and wait for networkidle
            2. Take screenshot or inspect DOM
            3. Identify selectors from rendered state
            4. Execute actions with discovered selectors
```

## Example: Using with_server.py

To start a server, run `--help` first, then use the helper:

**Single server:**
```bash
python scripts/with_server.py --server "npm run dev" --port 5173 -- python your_automation.py
```

**Multiple servers (e.g., backend + frontend):**
```bash
python scripts/with_server.py \
  --server "cd backend && python server.py" --port 3000 \
  --server "cd frontend && npm run dev" --port 5173 \
  -- python your_automation.py
```

To create an automation script, include only Playwright logic (servers are managed automatically):
```python
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True) # Always launch chromium in headless mode
    page = browser.new_page()
    page.goto('http://localhost:5173') # Server already running and ready
    page.wait_for_load_state('networkidle') # CRITICAL: Wait for JS to execute
    # ... your automation logic
    browser.close()
```

## Reconnaissance-Then-Action Pattern

1. **Inspect rendered DOM**:
   ```python
   page.screenshot(path='/tmp/inspect.png', full_page=True)
   content = page.content()
   page.locator('button').all()
   ```

2. **Identify selectors** from inspection results

3. **Execute actions** using discovered selectors

## Common Pitfall

❌ **Don't** inspect the DOM before waiting for `networkidle` on dynamic apps
✅ **Do** wait for `page.wait_for_load_state('networkidle')` before inspection

## Best Practices

- **Use bundled scripts as black boxes** - To accomplish a task, consider whether one of the scripts available in `scripts/` can help. These scripts handle common, complex workflows reliably without cluttering the context window. Use `--help` to see usage, then invoke directly. 
- Use `sync_playwright()` for synchronous scripts
- Always close the browser when done
- Use descriptive selectors: `text=`, `role=`, CSS selectors, or IDs
- Add appropriate waits: `page.wait_for_selector()` or `page.wait_for_timeout()`

## Reference Files

- **examples/** - Examples showing common patterns:
  - `element_discovery.py` - Discovering buttons, links, and inputs on a page
  - `static_html_automation.py` - Using file:// URLs for local HTML
  - `console_logging.py` - Capturing console logs during automation