// 일본 Swiftfab vs LGES SMC·MMF 비교분석 — CEO 보고용 Word(.docx) 생성
// 출처: references/경쟁사/2026-06-04_일본_Swiftfab_전지설비연합체_vs_SMC-MMF.md
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, ExternalHyperlink, ImageRun,
} = require("docx");

const FIGDIR = "outputs/figures";
function figure(file, caption) {
  const W = 600, Hh = Math.round(600 / (1240 / 720));
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 30 },
      children: [new ImageRun({ type: "png", data: fs.readFileSync(`${FIGDIR}/${file}`),
        transformation: { width: W, height: Hh },
        altText: { title: caption, description: caption, name: file } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 140 },
      children: [new TextRun({ text: caption, font: "맑은 고딕", size: 16, italics: true, color: "595959" })] }),
  ];
}

const FONT = "맑은 고딕";          // 한글 기본 (LG스마트체 미설치 환경 fallback)
const BLUE = "0000FF";            // 강조 파랑
const GREEN = "006600";           // 풋노트 초록
const GRAY_HEAD = "D9D9D9";       // 표 헤더 회색
const GRAY_BORD = "BFBFBF";
const CW = 9360;                  // content width (US Letter, 1" margin)

const border = { style: BorderStyle.SINGLE, size: 1, color: GRAY_BORD };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 110, right: 110 };

function t(text, opts = {}) { return new TextRun({ text, font: FONT, ...opts }); }
function p(children, opts = {}) {
  return new Paragraph({ children: Array.isArray(children) ? children : [children], ...opts });
}
function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 140 },
    children: [t(text, { bold: true, size: 28, color: "000000" })] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 },
    children: [t(text, { bold: true, size: 24, color: "000000" })] });
}
function bullet(runs) {
  return new Paragraph({ numbering: { reference: "b", level: 0 }, spacing: { after: 40 },
    children: Array.isArray(runs) ? runs : [t(runs)] });
}
function foot(text) {
  return new Paragraph({ spacing: { before: 80, after: 0 },
    children: [t(text, { color: GREEN, bold: true, size: 16 })] });
}

// 표 생성 헬퍼: rows = [[c1,c2,...], ...], 첫 행 헤더
function makeTable(colWidths, rows, opts = {}) {
  const blueCols = opts.blueCols || [];      // 강조 파랑 적용할 (row,col) — 여기선 사용 최소화
  const trs = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => {
      const isHead = ri === 0;
      const runs = (Array.isArray(cell) ? cell : [cell]).map((c) =>
        typeof c === "string"
          ? t(c, { bold: isHead, size: 17, color: isHead ? "000000" : "000000" })
          : c);
      return new TableCell({
        borders, width: { size: colWidths[ci], type: WidthType.DXA }, margins: cellMargins,
        verticalAlign: VerticalAlign.CENTER,
        shading: isHead ? { fill: GRAY_HEAD, type: ShadingType.CLEAR } : undefined,
        children: [new Paragraph({ spacing: { after: 0 }, children: runs })],
      });
    }),
  }));
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: colWidths, rows: trs });
}

const children = [];

// ── 표지 영역 ──
children.push(new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0 },
  children: [t("FA기술담당  |  2026.06.04", { size: 18, color: "000000" })] }));
children.push(new Paragraph({ spacing: { before: 120, after: 40 }, alignment: AlignmentType.CENTER,
  children: [t("일본 Swiftfab '블록형 모듈 전지공장' 연합체 분석", { bold: true, size: 34, color: "000000", font: FONT })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
  children: [t("— 당사(LGES) SMC·MMF 모듈러 기술과의 연관성 검토 —", { size: 24, color: BLUE, font: FONT })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 4 } },
  children: [t("CEO 지시 보고 · 팩트 기반 조사", { size: 18, color: "404040", font: FONT })] }));

// ── 0. 핵심 요약 (Executive Summary) ──
children.push(h1("핵심 요약 (Executive Summary)"));
children.push(bullet([t("일본 9개 설비사가 ", ), t("Swiftfab Energy Systems", { bold: true }),
  t(" 연합체를 결성(2026.4), 전지 생산설비를 "), t("규격화된 컨테이너 블록", { bold: true }),
  t("으로 모듈화해 고객 공장에서 조립 → 총비용 "), t("약 70% 절감", { bold: true, color: BLUE }),
  t(", 공장 건설기간 "), t("4~6년→2~3년", { bold: true, color: BLUE }), t(" 단축을 표방.")]));
children.push(bullet([t("이는 당사 FA기술담당의 "), t("SMC(2026)·MMF(2027) 모듈러 기술과 본질적 방향성이 동일", { bold: true }),
  t(" — 설비를 모듈화(Cube/Rack·이동식 조립)하여 셋업·건설 공수와 투자비를 절감하고 폼팩터 변화에 유연 대응.")]));
children.push(bullet([t("절감 대상은 "), t("설비+설치가 공통", { bold: true, color: BLUE }),
  t(" — 일본 70%는 "), t("건물까지 포함", { bold: true }), t(", 당사 −15%는 "), t("설비+설치 기준", { bold: true }),
  t("(건물은 건설담당 별도 영역, 건설기간 26→18개월 단축 목표). → '건물 포함 여부' 감안 비교.")]));
children.push(bullet([t("당사 강점: "), t("2026년 SMC Demo 실증 선행", { bold: true }),
  t("(일본 모델 2028·가동 2030 대비), 유연성·재활용성 정량지표(Phasing 80%·재활용률 50%) 차별화.")]));

// ── 1. 보고 배경 ──
children.push(h1("1. 보고 배경"));
children.push(p([t("일본경제신문(닛케이) 2026-06-02자 기사 ", ),
  t("\"EV電池、日立やリコーがブロック型工場 コスト7割減で中国に対抗\"", { italics: true }),
  t(" 에 대한 CEO 지시에 따라, 본 보고는 ① 기사 내용을 팩트 기반으로 상세 조사하고 ② 당사 SMC/MMF 모듈러 기술과의 연관성을 검토한다. 닛케이 한/일판·Nikkei Asia·히타치제작소 공식 보도자료(PR TIMES)를 교차 검증하였다.")],
  { spacing: { after: 80 } }));

// ── 2. Swiftfab 사실 요약 ──
children.push(h1("2. 일본 Swiftfab Energy Systems — 사실 요약"));
children.push(h2("2.1 결성 개요"));
children.push(makeTable([2600, 6760], [
  ["항목", "내용"],
  ["사업체명", [t("Swiftfab Energy Systems", { bold: true, size: 17 }), t(" (스위프트팹) — Swift(빠른)+Fabrication(제조)", { size: 17 })]],
  ["설립 시기", "2025-12-18 BASC 합의 발표 → 2026년 4월 법인 설립 (도쿄 미나토구)"],
  ["참여사 / CEO", "전지 서플라이체인협의회(BASC) 가입 9개사 공동 출자 / 사장 Keisuke Kida(키다 케이스케)"],
  ["사업 영역", "자동차용·정치용 리튬이온/차세대 전지 제조장치·라인 개발·설계·판매·운용지원"],
  ["사업 총액", [t("약 180억 엔 (약 1,700억 원) — 정부 보조금 포함", { size: 17 }), t("  ※원문 約180億円 (국문본 '181억'은 환율·번역차)", { size: 15, color: "808080" })]],
  ["일정", "2028년경 모델 설비 완성 → 전지 메이커 판매 / 첫 공장 가동 2030년 말 목표"],
  ["산업 위상", "中 글로벌 전지 ~70% 점유 / 日 설비메이커 글로벌 장비시장 9% vs 中 25%"],
]));

children.push(h2("2.2 참여 9개사 및 역할"));
children.push(makeTable([3000, 6360], [
  ["사", "역할 (공개 범위)"],
  ["히타치제작소(日立)", "디지털 트윈 시뮬레이션, HMAX Industry 축전지 확장, AI·현장데이터 설계→생산 최적화 (주도)"],
  ["리코(Ricoh)", "전지재료 잉크젯 디지털 제조 노하우 보유"],
  ["제이텍트(JTEKT)", "원천공정(源泉工程) 설비, 심플·슬림·콤팩트 설계 (도요타 계열)"],
  ["서부기연(西部技研)", "데시칸트 제습·전열교환, 드라이룸 환경관리"],
  ["고마츠NTC", "탭 성형·검사(120m/min), 레이저 용접·반송·모듈 자동화 (고마츠 자회사)"],
  ["도신(東伸) / 豊電子工業", "전극판 슬리팅·리와인딩 / 전장·자동화 설비"],
  ["히라타기공(平田機工) / 대기사(大気社)", "양산 조립라인(풋프린트 −30%) / 국소 클린룸·유틸리티 일체형 모듈 공법·NMP 회수"],
]));
children.push(foot("출처: Research Report.md [21~32행] / 히타치 보도자료(PR TIMES) / 닛케이 2026-06-02 · Nikkei Asia · Automotive World"));

children.push(h2("2.3 핵심 컨셉 — '블록형(모듈식) 공장'"));
children.push(bullet([t("전지 생산공정(재료가공→전극제조→셀조립→전해액 주액 등)을 잘게 세분화하여 생산장비·컨베이어를 "),
  t("규격화된 컨테이너 상자(블록)", { bold: true }), t("에 일체형으로 담음.")]));
children.push(bullet([t("고객 공장 현장에서 "), t("모듈(블록)처럼 자유롭게 연결", { bold: true }),
  t("해 공장을 완성. 약 1,000개 모듈 = 연 5만 대분 전지 라인.")]));
children.push(bullet([t("표준화된 설비 규격 + 장비 소형화 + "), t("공통 부품(센서·모터) 일괄 조달", { bold: true }), t("로 효율 개선.")]));

children.push(h2("2.4 정량 효과 (기사 주장치) 및 전략 의도"));
children.push(makeTable([3400, 5960], [
  ["지표", "효과"],
  ["총비용(건물+설비+설치)", [t("약 70% 절감", { bold: true, color: BLUE, size: 17 })]],
  ["공장 건설 기간", [t("기존 4~6년 → 2~3년", { bold: true, color: BLUE, size: 17 })]],
  ["운영비(OPEX)", "가동 후 절감 (구체 수치 미제시)"],
]));
children.push(bullet([t("배경: 일본 전지산업이 한·중에 밀려 고전. 원인으로 "), t("공장 1곳당 50곳 이상 공급사", { bold: true }),
  t("의 개별 조율 비효율 지목 → 연합체·표준화로 극복.")]));
children.push(bullet([t("전략: 중국 저가공세 대항, 리튬이온 부진 만회, "), t("차세대 전고체 배터리 시장에서 반격", { bold: true }), t(".")]));

children.push(h2("2.5 심층 리서치 추가 팩트 (출처: references/Research Report.md)"));
children.push(bullet([t("정부 정책 연계: ", { bold: true }), t("경산성 「축전지산업전략」 7대 중점과제 중 "), t("②전지설비산업의 구조변혁", { bold: true }), t("에 직접 대응. 경제안보추진법 기반 보조(설비투자 1/3·기술개발 1/2). 보조금 주체·비중은 미확인.")]));
children.push(bullet([t("전고체 OEM 일정과 정합: ", { bold: true }), t("도요타 2027~2028 양산, 닛산 FY2028 상용화, 혼다 2020년대 후반 → Swiftfab 차세대 라인(모델 2028) 시점 부합.")]));
children.push(bullet([t("컨테이너 모듈의 기술적 한계(FA 관점): ", { bold: true }), t("셀조립·주액 노점 −40℃↓, 권취/조립 ≤RH10%, 클린룸 ISO 7~8, 미크론급 정밀도·저진동 → "), t("모듈 분할 시 노점·차압·기밀·제진이 난제", { bold: true, color: BLUE }), t(" (대기사·서부기연 역할이 성패 관건).")]));
children.push(bullet([t("검증 한계: ", { bold: true }), t("70%·약1,000모듈·기간단축은 단일 1차 출처(닛케이) 의존. 현시점 "), t("'계획·목표'이며 실현 성과 아님", { bold: true }), t(", 수주처 미확보로 사업 리스크 상당.")]));

// ── 3. 당사 SMC·MMF ──
children.push(new Paragraph({ pageBreakBefore: true, children: [] }));
children.push(h1("3. 당사(LGES FA기술담당) SMC·MMF 모듈러 기술"));
children.push(p([t("지향점: "), t("\"시장·폼팩터 변화에 유연 대응 가능한 스마트 혁신 공장 구축으로 매몰 Loss 최소화 + CAPEX/OPEX 절감\"", { italics: true }),
  t(" (중장기 로드맵 v2 [22행]). Cell 단위 모듈화 + Phasing 투자 + 셋업 자동화 + 철거·재활용 가능 구조를 통합 적용.")],
  { spacing: { after: 80 } }));

children.push(h2("3.1 모듈러 과제 로드맵"));
children.push(makeTable([1300, 2600, 5460], [
  ["연도", "과제", "내용"],
  ["2026", [t("SMC", { bold: true, size: 17 }), t(" (Smart Modular Cube)", { size: 16 })], "다품종(Roll,Tray) 대응 모듈형 Cube Rack + High-pick AMR 셋업. Cell 단위 Phasing, 철거 후 재활용"],
  ["2027", [t("MMF", { bold: true, size: 17 }), t(" (Mobile Micro Factory)", { size: 16 })], "셋업 현장 내 이동 가능한 Rack 자동조립 로봇 시스템. Modular 라인으로 조립 라인 이동성 극대화"],
  ["2028", [t("밀폐형 Cube 물류 설비", { bold: true, size: 17 })], "DR Less 공정 밀폐형 Cube Rack + 전용 물류. 폼팩터 변경 대응 표준 Cube 구조"],
]));

children.push(h2("3.2 정량 효과 (As-is → To-be)"));
children.push(makeTable([1600, 2700, 3060, 2000], [
  ["과제", "지표", "As-is → To-be", "효과"],
  ["SMC", "면적 효율", "908 → 1,149개", [t("+25%", { bold: true, color: BLUE, size: 17 })]],
  ["SMC", "설치 공수", "3,234 → 2,100 MD", [t("−33%", { bold: true, color: BLUE, size: 17 })]],
  ["SMC", "투자비", "239.8 → 203.3억", [t("−15%", { bold: true, color: BLUE, size: 17 })]],
  ["MMF", "설치 공수", "2,280 → 1,232 MD", [t("−46%", { bold: true, color: BLUE, size: 17 })]],
  ["MMF", "투자비", "296.0 → 250.6억", [t("−15%", { bold: true, color: BLUE, size: 17 })]],
]));
children.push(foot("출처: references/roadmap/2026_FA기술담당_중장기로드맵_v2.md [62~69행], 동 로드맵 [286·291행]"));
figure("fig4_정량효과.png", "[그림 1] 당사 SMC·MMF 정량효과 (As-is → To-be) — 면적 +25%, 설치공수 −33~46%, 투자비 −15%").forEach((x) => children.push(x));

children.push(h2("3.3 KPI 목표 (현재→2028) 및 디지털 트윈"));
children.push(makeTable([3120, 1620, 1620, 1620, 1380], [
  ["축 / 지표", "현재", "2026", "2027", "2028"],
  ["CAPEX (투자비, 억/GWh)", "—", "88.4", "79.6", [t("70.7", { size: 16 }), t(" (−20%)", { size: 14, color: BLUE })]],
  ["면적 (평/GWh)", [t("1,254", { bold: true, size: 16 })], "1,129", "1,073", [t("1,019", { bold: true, size: 16 }), t(" (−20%)", { size: 14, color: BLUE })]],
  ["건설 기간 (개월)", [t("26", { bold: true, size: 16 })], "25", "23", [t("18", { bold: true, size: 16, color: BLUE }), t(" (−31%)", { size: 14, color: BLUE })]],
  ["유연성 (셋업 LT)", "기준", "—", "−30%", "−50%"],
  ["유연성 (Phasing 비율)", "Pilot", "—", "50%", "80%"],
  ["재활용성 (재활용률)", "기준", "—", "30%", "50%"],
]));
children.push(bullet([t("면적·절감 대상은 "), t("설비+설치 효율화", { bold: true }), t(" (건물 제외). 건설 기간(건물)은 "), t("건설담당 별도 영역", { bold: true }), t("으로 26→18개월 단축 목표 — 일본 Swiftfab 건물 단축(4~6년→2~3년)에 대응.")]));
children.push(bullet([t("디지털 트윈: "), t("CNS Physical Works(Omniverse, NVIDIA Isaac SIM)", { bold: true }),
  t(" 연계 — AI 설계 자동화 + 가상 사전 검증 (로드맵 v2 §4).")]));

// ── 4. 비교 분석 ──
children.push(new Paragraph({ pageBreakBefore: true, children: [] }));
children.push(h1("4. 연관성 · 비교 분석"));
children.push(p([t("양사 모두 "), t("'설비를 규격화·모듈화하여 셋업·건설 공수와 투자비를 절감하고, 시장·폼팩터 변화에 유연 대응'", { bold: true }),
  t(" 한다는 동일한 철학. 일본 기사의 문제의식(공급사 난립·셋업 비효율·중국 저가공세)은 당사 로드맵의 문제의식(매몰 Loss·셋업비·유연성)과 같은 축이다.")],
  { spacing: { after: 100 } }));
figure("fig1_개념비교.png", "[그림 2] 개념 비교 — 같은 '모듈러' 방향, 다른 적용 범위 (Swiftfab 공장 전체 블록 vs LGES 공정 단위 모듈)").forEach((x) => children.push(x));
children.push(makeTable([1900, 3730, 3730], [
  ["관점", "일본 Swiftfab", "LGES SMC·MMF"],
  ["주체", "9개 설비사 연합체 (공동출자·외판 비즈니스)", "FA기술담당 자가 개발·자가 적용"],
  ["모듈 단위", "공정 전체를 컨테이너 블록 (~1,000모듈=5만대분)", "Cube Rack·이동식 조립 셀 (물류·셋업 중심)"],
  ["절감 대상(분모)", [t("건물 + 설비 + 설치", { bold: true, size: 16 })], [t("설비 + 설치", { bold: true, size: 16 }), t(" (건물은 건설담당 별도)", { size: 15, color: "808080" })]],
  ["비용 절감(주장)", [t("총비용 70% (건물 포함)", { bold: true, color: BLUE, size: 16 })], [t("투자비 −15%, 설치공수 −33~46% (설비+설치 기준)", { size: 16 })]],
  ["건설 기간(건물)", "공장 건설 4~6년→2~3년 (≈24~36개월)", [t("26→18개월 (2028, −31%)", { size: 16 }), t(" *건설담당", { size: 15, color: "808080" })]],
  ["셋업/면적", "(별도 명시 없음)", "셋업 LT −50%(2028) / 면적 1,254→1,019 (−20%)"],
  ["유연성·재활용", "모듈 재사용 가능(언급)", [t("명시적 지표화 — Phasing 80%, 재활용률 50%", { bold: true, size: 16 })]],
  ["디지털 트윈", "히타치 시뮬레이션(소프트 디버깅)", "CNS Physical Works·Omniverse·Isaac SIM"],
  ["차세대 타깃", "전고체 시장 반격", "폼팩터 변화 대응 표준 Cube(전고체 포함)"],
  ["시간축", "모델 2028 / 첫 공장 2030", "SMC 2026 Demo / MMF 2027 / 밀폐형 2028"],
]));
children.push(foot("주의: 절감 대상은 설비+설치가 공통. 일본 70%는 건물 포함분 / 당사 −15%는 설비+설치 기준(건물은 건설담당 별도, 26→18개월) → '건물 포함 여부' 감안 비교."));
figure("fig2_분모차이.png", "[그림 3] 절감 대상 범위 비교 — 설비+설치는 공통, 건물 포함 여부가 차이 (당사 건물은 건설담당 별도 영역)").forEach((x) => children.push(x));

children.push(h2("4.1 핵심 차이 3가지"));
children.push(bullet([t("절감 대상 범위: ", { bold: true }), t("설비+설치는 공통, 건물 포함 여부가 차이 — 일본 70%는 건물 포함, 당사는 설비+설치(건물은 건설담당이 26→18개월로 별도 단축).")]));
children.push(bullet([t("비즈니스 모델: ", { bold: true }), t("일본=표준 설비 타사 판매(공급망 표준화·규모 경제) / 당사=내재화 경쟁력(self-use).")]));
children.push(bullet([t("성숙도: ", { bold: true }), t("당사 SMC는 2026년 Demo 실증 진행 중, 일본은 2028 모델·2030 가동 → 당사 실증 선행.")]));
figure("fig5_일정_면적목표.png", "[그림 4] 공장 구축 일정·면적 효율 목표 — 건설기간 26→18개월(−31%), 면적 1,254→1,019평/GWh(−20%)").forEach((x) => children.push(x));
figure("fig3_타임라인.png", "[그림 5] 개발·실증 타임라인 — 당사 실증(2026)이 일본 가동(2030) 대비 약 4년 시간 우위").forEach((x) => children.push(x));

// ── 5. 시사점 및 대응 ──
children.push(h1("5. 시사점 및 대응 방향 (CEO 보고용)"));
children.push(h2("5.1 위협 (Threat)"));
children.push(bullet([t("공급망 표준화 압력: ", { bold: true }), t("일본 연합체가 표준 모듈 설비를 외판하면 후발 메이커 진입장벽이 낮아지고 설비 범용화 가속 → 당사 내재화 경쟁력 상대적 희석 우려.")]));
children.push(bullet([t("전고체 선점 의도: ", { bold: true }), t("일본이 차세대 전고체 라인 표준을 선점하면 폼팩터 전환기 주도권 경쟁에서 불리.")]));
children.push(h2("5.2 기회 (Opportunity) — 당사 강점"));
children.push(bullet([t("방향성 외부검증: ", { bold: true }), t("일본 9개사 연합이 당사 모듈러 방향이 산업 정답임을 입증. 그룹장 코멘트(유연성·재활용성 지표 신설)와 정합.")]));
children.push(bullet([t("실증 선행: ", { bold: true }), t("당사는 2026 SMC Demo·MMF 1차 설계 등 이미 실행 단계 → 일본(2028 모델) 대비 시간 우위.")]));
children.push(bullet([t("차별화 축: ", { bold: true }), t("일본이 명시하지 않은 재활용성(50%)·유연성 정량지표(Phasing 80%)가 당사 고유 차별화 포인트.")]));
children.push(h2("5.3 제언"));
children.push(bullet([t("모듈 인터페이스·공통조달 벤치마킹: ", { bold: true }), t("Swiftfab의 '컨테이너 블록 규격화 + 공통 부품 일괄 조달'을 당사 Cube/Rack 표준화·공통조달에 참고.")]));
children.push(bullet([t("디지털 트윈 사전검증 가속: ", { bold: true }), t("히타치 디지털 트윈 대응, CNS Physical Works·Isaac SIM 가상 검증 로드맵(2027~2028) 조기화 검토.")]));
children.push(bullet([t("전고체 모듈러 대응 명문화: ", { bold: true }), t("2028 밀폐형 Cube·표준 Cube 구조를 전고체 폼팩터 전환 대응으로 명시 연결.")]));
children.push(bullet([t("중국 저가공세 대응 스토리: ", { bold: true }), t("self-use 고도화 + 유연성·재활용성 선도로 TCO 우위 강화 (CATL 95% 자동화·1초/셀 대비 유연성 차원 차별화).")]));

children.push(h2("5.4 판단 전환 기준 (위협 격상 임계값)"));
children.push(p([t("Swiftfab은 현재 "), t("컨셉·로드맵 단계", { bold: true }),
  t("로, 즉각 위협이라기보다 2028~2030 설비 모듈화·표준화 경쟁의 신호탄이다. 아래 3개 중 "),
  t("2개 이상 충족 시 '실질 위협'으로 격상", { bold: true, color: BLUE }),
  t("하고 당사 모듈러 대응 투자를 가속한다.")], { spacing: { after: 60 } }));
children.push(bullet([t("① 실제 "), t("수주처", { bold: true }), t("(비중국 전지 메이커·OEM) 확보 발표")]));
children.push(bullet([t("② 총비용 "), t("70% 절감의 제3자 검증", { bold: true })]));
children.push(bullet([t("③ 첫 공장 "), t("양산 수율 ≥90%", { bold: true }), t(" 달성")]));
children.push(foot("모니터링: 2028 모델 설비 사양(노점 유지·모듈 접합 기밀·CAPEX/GWh), 대기사·서부기연 일체형 모듈 공법, 히타치 HMAX 디지털 트윈 vs 당사 가상 시운전. (출처: Research Report.md [99행])"));

// ── 6. 출처 ──
children.push(h1("6. 출처 (Sources)"));
const srcs = [
  ["일본경제신문(닛케이), \"EV電池、日立やリコーがブロック型工場 コスト7割減で中国に対抗\" (2026-06-02)", "https://www.nikkei.com/article/DGXZQOJC203B90Q6A520C2000000/"],
  ["일본경제신문, \"日立など9社、電池製造設備で新会社 一体開発で低コスト化\"", "https://www.nikkei.com/article/DGXZQOUC186UF0Y5A211C2000000/"],
  ["히타치제작소 보도자료, \"蓄電池製造設備産業の強化をめざす共同事業「Swiftfab」始動\" (PR TIMES)", "https://prtimes.jp/main/html/rd/p/000000047.000141666.html"],
];
srcs.forEach(([label, url]) => children.push(bullet([t(label + " "),
  new ExternalHyperlink({ children: [new TextRun({ text: url, style: "Hyperlink", font: FONT, size: 16 })], link: url })])));
children.push(bullet([t("Nikkei Asia, \"Hitachi, Ricoh to build low-cost modular EV battery plants\" / Automotive World, \"Japanese JV targets a 70% cut in battery plant costs\" / Just-Auto (GlobalData) (2026)")]));
children.push(bullet([t("(내부) ", { bold: true }), t("references/Research Report.md", { bold: true }), t(" — Claude 심층 리서치(9개사 역할·정부정책·전고체 OEM·기술 한계·판단 전환 기준)")]));
children.push(bullet([t("(내부) references/경쟁사/2026-06-02_닛케이_블록형공장_원문아카이브.md — 기사 원문+출처 URL")]));
children.push(bullet([t("(내부) references/roadmap/2026_FA기술담당_중장기로드맵_v2.md · 동 로드맵 원본 · references/경쟁사/유연성_경쟁사_분석_2026-05-21.md")]));
children.push(foot("(출처: references/경쟁사/2026-06-04_일본_Swiftfab_전지설비연합체_vs_SMC-MMF.md — 팩트체크 원본)"));

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: FONT, color: "000000" },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "808080", space: 2 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: FONT, color: "000000" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 480, hanging: 240 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 },
      margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [t("LGES FA기술담당 · CEO 보고 · 2026.06.04          ", { size: 14, color: "808080" }),
        t("- ", { size: 14, color: "808080" }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 14, color: "808080" }),
        t(" -", { size: 14, color: "808080" })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("outputs/일본Swiftfab_vs_SMC-MMF_비교분석_2026-06-04.docx", buf);
  console.log("WROTE outputs/일본Swiftfab_vs_SMC-MMF_비교분석_2026-06-04.docx");
});
