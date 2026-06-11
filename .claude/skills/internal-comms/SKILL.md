---
name: internal-comms
description: 사내 공지·주간 메일·리더십 업데이트·3P 업데이트(Progress/Plans/Problems)·사내 뉴스레터·FAQ·인시던트 리포트·프로젝트 진행상황 메모 등 내부 커뮤니케이션 문서의 표준 양식·톤으로 작성하는 스킬. LGES FA 자동화 프로젝트의 주간 KPI 메일, 분기 부서 업데이트, 사내 공지, 장애·이슈 리포트, FA 그룹 All-hands 자료 작성 시 사용한다. 사용자가 '사내 공지 만들어줘', '주간 메일 정리', '3P 업데이트 작성', '리더십 업데이트', '사내 뉴스레터', 'FAQ 작성', '이슈 리포트' 같이 말하면 트리거. fa-task-discovery 의 카카오톡 요약을 사내 공식 채널 버전으로 확장할 때, document-writer 가 주간보고를 메일/공지 형태로 환승할 때 사용. A set of resources to help me write all kinds of internal communications, using the formats that my company likes to use. Claude should use this skill whenever asked to write some sort of internal communications (status reports, leadership updates, 3P updates, company newsletters, FAQs, incident reports, project updates, etc.).
license: Complete terms in LICENSE.txt (Anthropic skills 공식 원본 + FA 트리거 구절 추가)
---

> **LGES FA 프로젝트 적용 노트 (본 리포 한정)**
>
> - **결합 자산**: document-writer (양식 채움), fa-task-discovery (카톡 요약 → 사내 공식 채널 버전).
> - **양식 위치**: 사내 메일·공지 양식이 `templates/사내양식/` 에 없으면 본 스킬 가이드를 기준으로 신규 양식 작성 가능.
> - **톤**: LGES 사내 톤은 `templates/사내양식/LGES_양식_톤_지침.md` 참조. 외부 발송용은 별도 검토.
> - **브랜드 룰**: 메일 본문(text/HTML) 에도 가능한 `brand-guidelines` 의 색·강조 룰 일관 적용.

## When to use this skill
To write internal communications, use this skill for:
- 3P updates (Progress, Plans, Problems)
- Company newsletters
- FAQ responses
- Status reports
- Leadership updates
- Project updates
- Incident reports

## How to use this skill

To write any internal communication:

1. **Identify the communication type** from the request
2. **Load the appropriate guideline file** from the `examples/` directory:
    - `examples/3p-updates.md` - For Progress/Plans/Problems team updates
    - `examples/company-newsletter.md` - For company-wide newsletters
    - `examples/faq-answers.md` - For answering frequently asked questions
    - `examples/general-comms.md` - For anything else that doesn't explicitly match one of the above
3. **Follow the specific instructions** in that file for formatting, tone, and content gathering

If the communication type doesn't match any existing guideline, ask for clarification or more context about the desired format.

## Keywords
3P updates, company newsletter, company comms, weekly update, faqs, common questions, updates, internal comms
