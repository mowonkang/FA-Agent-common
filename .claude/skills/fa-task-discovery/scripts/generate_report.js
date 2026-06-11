// FA 기술과제 발굴 보고서 생성 스크립트
// 사용법: node generate_report.js <output_path> <tasks_json_path|-> [news_json_path]
//   - tasks.json + news.json → 제1부 뉴스 + 제2부 과제 2부 docx (기본)
//   - tasks.json 만           → 과제 단일부 docx (구버전 폴백)
//   - news.json 만 (tasks="-") → 뉴스 단일부 docx (디버그용)
// tasks_json: [{num, category, title, trend, apply, reference}, ...] 형식
// news_json:  [{category, title, url, date, summary}, ...] 형식

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, PageBreak,
  ExternalHyperlink
} = require('docx');
const fs = require('fs');
const path = require('path');

const COLOR_NAVY = "1F4E79";
const COLOR_LIGHT = "D9E2F3";
const COLOR_WHITE = "FFFFFF";
const COLOR_GRAY = "CCCCCC";

const border = { style: BorderStyle.SINGLE, size: 1, color: COLOR_GRAY };
const borders = { top: border, bottom: border, left: border, right: border };

function makeHeading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLOR_NAVY, space: 4 } },
    children: [new TextRun({ text, font: "맑은 고딕", size: 26, bold: true, color: COLOR_NAVY })]
  });
}

function makeHeading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text, font: "맑은 고딕", size: 22, bold: true, color: COLOR_NAVY })]
  });
}

function makeBody(text, options = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60, line: 360 },
    children: [new TextRun({ text, font: "맑은 고딕", size: 20, ...options })]
  });
}

function makeSummaryRows(tasks) {
  const header = ["분야", "과제명", "핵심 적용 포인트"];
  const colWidths = [1400, 3200, 4426];
  const tableWidth = 9026;

  const headerRow = new TableRow({
    children: header.map((h, i) => new TableCell({
      borders, width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: COLOR_NAVY, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: h, font: "맑은 고딕", size: 18, bold: true, color: COLOR_WHITE })] })]
    }))
  });

  const dataRows = tasks.map((t, i) => new TableRow({
    children: [
      new TableCell({ borders, width: { size: colWidths[0], type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? COLOR_LIGHT : COLOR_WHITE, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: t.category, font: "맑은 고딕", size: 17, bold: true, color: COLOR_NAVY })] })] }),
      new TableCell({ borders, width: { size: colWidths[1], type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? COLOR_LIGHT : COLOR_WHITE, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: t.title, font: "맑은 고딕", size: 17 })] })] }),
      new TableCell({ borders, width: { size: colWidths[2], type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? COLOR_LIGHT : COLOR_WHITE, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: t.apply, font: "맑은 고딕", size: 17, color: "333333" })] })] }),
    ]
  }));

  return new Table({ width: { size: tableWidth, type: WidthType.DXA }, columnWidths: colWidths, rows: [headerRow, ...dataRows] });
}

function makeTaskCard(t) {
  const colWidths = [1400, 7626];
  const headerRow = new TableRow({
    children: [new TableCell({
      borders, columnSpan: 2, width: { size: 9026, type: WidthType.DXA },
      shading: { fill: COLOR_NAVY, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 140, right: 140 },
      children: [new Paragraph({ children: [
        new TextRun({ text: `[${t.num}] `, font: "맑은 고딕", size: 20, bold: true, color: "FFFF00" }),
        new TextRun({ text: `${t.category}  |  `, font: "맑은 고딕", size: 20, bold: true, color: "AACCFF" }),
        new TextRun({ text: t.title, font: "맑은 고딕", size: 20, bold: true, color: COLOR_WHITE }),
      ]})]
    })]
  });

  const makeRow = (label, content) => new TableRow({
    children: [
      new TableCell({ borders, width: { size: colWidths[0], type: WidthType.DXA },
        shading: { fill: COLOR_LIGHT, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 }, verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ children: [new TextRun({ text: label, font: "맑은 고딕", size: 18, bold: true, color: COLOR_NAVY })] })] }),
      new TableCell({ borders, width: { size: colWidths[1], type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: content, font: "맑은 고딕", size: 18 })] })] }),
    ]
  });

  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: colWidths,
    rows: [headerRow, makeRow("동향 근거", t.trend), makeRow("적용 포인트", t.apply), makeRow("참고 사례", t.reference)] });
}

// ── 제1부 (뉴스 브리핑) 헬퍼 ──────────────────────────────────────────

// 6개 분야 고정 표시 순서
const DOMAIN_ORDER = ["AMR/AGV", "협동로봇", "디지털 트윈", "스마트 물류", "AI", "피지컬 AI"];

// 별칭/부분일치 → 표준 분야명. 매칭 실패 시 null.
function normalizeCategory(raw) {
  if (!raw) return null;
  const s = String(raw).toLowerCase().replace(/[\s()/]/g, '');
  const rules = [
    ["피지컬 AI", ["피지컬", "physical", "휴머노이드", "humanoid", "임바디드", "embodied"]],
    ["AI", ["ai", "인공지능", "비전", "vision", "머신러닝"]],
    ["AMR/AGV", ["amr", "agv", "자율이동", "ammr", "이동로봇"]],
    ["협동로봇", ["협동로봇", "코봇", "cobot", "collaborativerobot"]],
    ["디지털 트윈", ["디지털트윈", "digitaltwin", "트윈", "시뮬레이션"]],
    ["스마트 물류", ["스마트물류", "물류", "logistics", "인트라로지스틱", "창고", "wms"]],
  ];
  for (const [canon, keys] of rules) {
    if (keys.some(k => s.includes(k.replace(/\s/g, '')))) return canon;
  }
  return null;
}

function makeNewsLink(url, label) {
  if (!url || !/^https?:\/\//i.test(url)) {
    return new TextRun({ text: label, font: "맑은 고딕", size: 17 });
  }
  return new ExternalHyperlink({
    link: url,
    children: [new TextRun({ text: label, font: "맑은 고딕", size: 17, color: "1155CC", underline: {} })]
  });
}

function makeNewsTable(items) {
  const header = ["발행일", "기사 제목 (출처 링크)", "핵심 요약"];
  const colWidths = [1300, 4326, 3400];
  const tableWidth = 9026;

  const headerRow = new TableRow({
    tableHeader: true,
    children: header.map((h, i) => new TableCell({
      borders, width: { size: colWidths[i], type: WidthType.DXA },
      shading: { fill: COLOR_NAVY, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: h, font: "맑은 고딕", size: 18, bold: true, color: COLOR_WHITE })] })]
    }))
  });

  const dataRows = items.map((it, i) => {
    const fill = i % 2 === 0 ? COLOR_LIGHT : COLOR_WHITE;
    return new TableRow({
      children: [
        new TableCell({ borders, width: { size: colWidths[0], type: WidthType.DXA },
          shading: { fill, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 120 }, verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ children: [new TextRun({ text: it.date || "-", font: "맑은 고딕", size: 16, color: "333333" })] })] }),
        new TableCell({ borders, width: { size: colWidths[1], type: WidthType.DXA },
          shading: { fill, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [makeNewsLink(it.url, it.title || "(제목 없음)")] })] }),
        new TableCell({ borders, width: { size: colWidths[2], type: WidthType.DXA },
          shading: { fill, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: it.summary || "", font: "맑은 고딕", size: 16, color: "333333" })] })] }),
      ]
    });
  });

  return new Table({ width: { size: tableWidth, type: WidthType.DXA }, columnWidths: colWidths, rows: [headerRow, ...dataRows] });
}

// 메인 실행
const outputPath = process.argv[2] || '/tmp/FA_기술과제후보.docx';
const tasksJsonPath = process.argv[3];
const newsJsonPath = process.argv[4];

// tasks.json (세션2 산출). 세션1(뉴스 only) 에서는 생략 가능 — 인자에 "-" 또는 "" 전달.
let tasks = [];
if (tasksJsonPath && tasksJsonPath !== '-' && fs.existsSync(tasksJsonPath)) {
  const rawTasks = JSON.parse(fs.readFileSync(tasksJsonPath, 'utf-8'));
  if (Array.isArray(rawTasks)) tasks = rawTasks;
}

// news.json (선택) — 없거나 비어 있으면 제1부 생략, 기존 동작으로 폴백
let newsByDomain = null;
let newsCount = 0;
if (newsJsonPath && fs.existsSync(newsJsonPath)) {
  const rawNews = JSON.parse(fs.readFileSync(newsJsonPath, 'utf-8'));
  if (Array.isArray(rawNews) && rawNews.length > 0) {
    newsByDomain = {};
    DOMAIN_ORDER.forEach(d => { newsByDomain[d] = []; });
    const etc = [];
    for (const item of rawNews) {
      const d = normalizeCategory(item.category);
      if (d) newsByDomain[d].push(item);
      else etc.push(item);
    }
    // 분야 매칭 실패 항목은 마지막 분야(피지컬 AI)에 흡수하지 않고 별도 "기타"로 보존
    if (etc.length > 0) newsByDomain["기타"] = etc;
    newsCount = rawNews.length;
  }
}

if (tasks.length === 0 && !newsByDomain) {
  console.error('news.json 또는 tasks.json 중 최소 하나가 필요합니다. 사용법: node generate_report.js <output.docx> <tasks.json|-> [news.json]');
  process.exit(1);
}

const today = new Date().toLocaleDateString('ko-KR', { year: 'numeric', month: '2-digit', day: '2-digit' }).replace(/\. /g, '.').replace('.', '년 ').replace('.', '월 ').replace('.', '일');
const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');

const doc = new Document({
  styles: {
    default: { document: { run: { font: "맑은 고딕", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "맑은 고딕", color: COLOR_NAVY },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "맑은 고딕", color: COLOR_NAVY },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 } },
    ]
  },
  numbering: {
    config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }]
  },
  sections: [{
    properties: {
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } }
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_NAVY, space: 4 } },
        children: [new TextRun({ text: "FA 기술과제 발굴 보고서  |  LG에너지솔루션 FA 생산기술팀", font: "맑은 고딕", size: 16, color: "888888" })]
      })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "- ", font: "맑은 고딕", size: 16, color: "888888" }),
          new TextRun({ children: [PageNumber.CURRENT], font: "맑은 고딕", size: 16, color: "888888" }),
          new TextRun({ text: " -", font: "맑은 고딕", size: 16, color: "888888" }),
        ]
      })] })
    },
    children: buildBody()
  }]
});

function buildBody() {
  const hasNews = !!newsByDomain;
  const hasTasks = tasks.length > 0;
  const twoPart = hasNews && hasTasks;
  const p1 = twoPart ? "제1부 · " : "";
  const p2 = twoPart ? "제2부 · " : "";

  const cover = [
    new Paragraph({ spacing: { before: 800, after: 200 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "FA 기술과제 발굴 보고서", font: "맑은 고딕", size: 52, bold: true, color: COLOR_NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 100 },
      children: [new TextRun({ text: "Factory Automation Technology Task Discovery Report", font: "맑은 고딕", size: 24, color: "555555" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 60 },
      children: [new TextRun({ text: today, font: "맑은 고딕", size: 22, color: "444444" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: "강모원  |  LG에너지솔루션 FA 생산기술", font: "맑은 고딕", size: 22, color: "444444" })] }),
    new Paragraph({ children: [new PageBreak()] }),
  ];

  // ── 제1부 · 최근 1개월 FA 기술 동향 뉴스 ──
  const part1 = [];
  if (hasNews) {
    const ago = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000)
      .toLocaleDateString('ko-KR', { year: 'numeric', month: '2-digit', day: '2-digit' });
    part1.push(
      makeHeading1(`${p1}최근 1개월 FA 기술 동향 뉴스`),
      makeBody(`수집 기준: ${ago} ~ ${today} (최근 30일 이내 발행). 6개 분야별 최신 기사 ${newsCount}건을 출처 링크와 함께 정리했습니다.`, { color: "555555" }),
      new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
    );
    const domains = [...DOMAIN_ORDER];
    if (newsByDomain["기타"]) domains.push("기타");
    for (const d of domains) {
      const items = newsByDomain[d] || [];
      part1.push(makeHeading2(`〔${d}〕`));
      if (items.length === 0) {
        part1.push(makeBody("이번 주기 신규 동향 없음", { color: "888888", italics: true }));
      } else {
        part1.push(makeNewsTable(items));
      }
      part1.push(new Paragraph({ spacing: { before: 100, after: 0 }, children: [] }));
    }
    if (hasTasks) part1.push(new Paragraph({ children: [new PageBreak()] }));
  }

  // ── 제2부 · FA 기술과제 후보 (기존 로직 유지) ──
  const part2 = [
    makeHeading1(`${p2}Executive Summary`),
    makeBody(`본 보고서는 배터리 공장 FA 분야 6개 기술 주제의 최신 글로벌 동향을 분석하고, LGES 배터리 공장 적용 가능성이 높은 신규 기술과제 후보 ${tasks.length}건을 도출한 결과물입니다.`),
    new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
    makeBody("과제 후보 선정 기준: ① 배터리 셀/모듈/팩 생산라인 실질 적용 가능, ② LGES 공식 양산 도입 발표 없음"),
    new Paragraph({ spacing: { before: 120, after: 80 }, children: [] }),

    makeHeading1(`${p2}FA 기술과제 후보 요약`),
    makeSummaryRows(tasks),
    new Paragraph({ children: [new PageBreak()] }),

    makeHeading1(`${p2}FA 기술과제 후보 (${tasks.length}건) 상세`),
    makeBody("아래 과제 후보는 수집된 글로벌·국내 동향에서 도출되었으며, LGES 배터리 공장 적용 타당성을 기준으로 선별되었습니다.", { color: "555555" }),
    new Paragraph({ spacing: { before: 120, after: 0 }, children: [] }),

    ...tasks.flatMap(t => [
      makeTaskCard(t),
      new Paragraph({ spacing: { before: 160, after: 0 }, children: [] }),
    ]),

    new Paragraph({ children: [new PageBreak()] }),
    makeHeading1("시사점 및 Next Step"),
    makeHeading2("제안 Next Step"),
    ...[
      "과제 후보 중 우선순위 선별 후 타당성 검토(FS) 착수",
      "선별 과제는 임원 보고용 1페이지 과제 제안서로 별도 작성",
      "글로벌 선진 사례(CATL, 현대차 등) 벤치마킹 심층 조사",
    ].map(text => new Paragraph({
      numbering: { reference: "bullets", level: 0 },
      spacing: { before: 60, after: 60, line: 340 },
      children: [new TextRun({ text, font: "맑은 고딕", size: 19 })]
    })),
  ];

  return [...cover, ...part1, ...(hasTasks ? part2 : [])];
}

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log('DONE:' + outputPath);
});
