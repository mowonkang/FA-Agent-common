#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""일본 Swiftfab vs LGES SMC·MMF 설명용 다이어그램 4종 (SVG + PNG).
LGES 브랜드: 화이트+검정+회색 / 강조 파랑 #0000FF / 풋노트 초록 #006600.
출처: references/경쟁사/2026-06-04_일본_Swiftfab_전지설비연합체_vs_SMC-MMF.md
"""
import os
import cairosvg

OUT = "outputs/figures"
os.makedirs(OUT, exist_ok=True)

BLUE = "#0000FF"
GREEN = "#006600"
BLACK = "#1a1a1a"
GRAY = "#808080"
LGRAY = "#BFBFBF"
FILL = "#F2F2F2"
HEAD = "#D9D9D9"
BLUEFILL = "#E8EEFF"
FONT = "NanumGothic, 'Noto Sans CJK KR', sans-serif"

W, H = 1240, 720


def svg_open(w=W, h=H):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" '
            f'viewBox="0 0 {w} {h}" font-family="{FONT}">'
            f'<rect width="{w}" height="{h}" fill="white"/>')


def txt(x, y, s, size=20, fill=BLACK, weight="normal", anchor="start", spacing="0"):
    return (f'<text x="{x}" y="{y}" font-size="{size}" fill="{fill}" '
            f'font-weight="{weight}" text-anchor="{anchor}" letter-spacing="{spacing}">{s}</text>')


def rect(x, y, w, h, fill="white", stroke=LGRAY, sw=1.5, rx=6):
    return f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>'


def title_bar(t, sub):
    s = txt(40, 50, t, size=30, weight="bold", fill=BLACK)
    s += txt(40, 82, sub, size=17, fill=GRAY)
    s += f'<line x1="40" y1="98" x2="{W-40}" y2="98" stroke="{BLACK}" stroke-width="2"/>'
    s += txt(W - 40, 50, "FA기술담당 (2026.06.04)", size=15, fill=BLACK, anchor="end")
    return s


def footnote(t):
    return txt(40, H - 22, t, size=15, fill=GREEN, weight="bold")


def save(name, body):
    svg = svg_open() + body + "</svg>"
    p_svg = f"{OUT}/{name}.svg"
    with open(p_svg, "w", encoding="utf-8") as f:
        f.write(svg)
    cairosvg.svg2png(bytestring=svg.encode("utf-8"), write_to=f"{OUT}/{name}.png", scale=2.0)
    print("wrote", p_svg, "+ png")


# ════════════════════════════════════════════════════════════════════
# Fig 1 — 개념 비교도
# ════════════════════════════════════════════════════════════════════
def fig1():
    b = title_bar("개념 비교 — 같은 '모듈러' 방향, 다른 적용 범위",
                  "일본 Swiftfab '블록형 공장 전체' vs LGES SMC·MMF '공정 단위 모듈'")
    # 좌 패널 (Swiftfab)
    lx, pw, py, ph = 40, 565, 120, 470
    b += rect(lx, py, pw, ph, fill="white", stroke=GRAY, sw=2)
    b += rect(lx, py, pw, 44, fill=HEAD, stroke=GRAY, sw=2, rx=6)
    b += txt(lx + 20, py + 30, "일본 Swiftfab — 블록형 공장 전체 모듈화", size=19, weight="bold")
    # 공장 외곽 + 컨테이너 블록 4개
    fx, fy, fw = lx + 40, py + 80, pw - 80
    b += f'<rect x="{fx}" y="{fy}" width="{fw}" height="240" rx="4" fill="none" stroke="{BLACK}" stroke-width="2" stroke-dasharray="6 4"/>'
    b += txt(fx + 8, fy - 8, "고객 공장(현장 조립)", size=14, fill=GRAY)
    blocks = ["재료 가공", "전극 제조", "셀 조립", "전해액 주액"]
    bw = (fw - 30) / 4
    for i, name in enumerate(blocks):
        bx = fx + 12 + i * bw
        b += rect(bx, fy + 30, bw - 14, 110, fill=FILL, stroke=BLACK, sw=1.5, rx=4)
        # 컨테이너 빗금
        b += f'<line x1="{bx}" y1="{fy+55}" x2="{bx+bw-14}" y2="{fy+55}" stroke="{LGRAY}" stroke-width="1"/>'
        b += txt(bx + (bw - 14) / 2, fy + 95, "컨테이너", size=12, fill=GRAY, anchor="middle")
        b += txt(bx + (bw - 14) / 2, fy + 120, name, size=14, weight="bold", anchor="middle")
        if i < 3:
            cx = bx + bw - 10
            b += txt(cx, fy + 92, "+", size=22, fill=BLUE, weight="bold", anchor="middle")
    b += txt(fx, fy + 175, "규격화 컨테이너 블록을 현장에서 레고처럼 연결", size=15)
    b += txt(fx, fy + 200, "→ 약 1,000모듈 = 연 5만 대분 라인", size=15)
    # blue tag
    b += rect(lx + 40, py + ph - 70, pw - 80, 46, fill=BLUEFILL, stroke=BLUE, sw=2, rx=8)
    b += txt(lx + pw / 2, py + ph - 40, "총비용 ▼ 70% (건물+설비+설치)", size=20, weight="bold", fill=BLUE, anchor="middle")

    # 우 패널 (LGES)
    rxp = lx + pw + 30
    b += rect(rxp, py, pw, ph, fill="white", stroke=BLUE, sw=2)
    b += rect(rxp, py, pw, 44, fill=HEAD, stroke=BLUE, sw=2, rx=6)
    b += txt(rxp + 20, py + 30, "LGES SMC·MMF — 공정 단위 모듈화", size=19, weight="bold")
    # SMC Cube Rack (stacked cells)
    sx, sy = rxp + 35, py + 80
    b += txt(sx, sy + 4, "SMC (2026) — 모듈형 Cube Rack", size=15, weight="bold")
    for r in range(4):
        for c in range(3):
            b += rect(sx + c * 46, sy + 20 + r * 30, 42, 26, fill=FILL, stroke=BLACK, sw=1.2, rx=3)
    b += txt(sx + 160, sy + 50, "Cell 단위", size=14)
    b += txt(sx + 160, sy + 72, "Phasing 셋업", size=14)
    b += txt(sx + 160, sy + 94, "철거 후 재활용", size=14, fill=BLUE, weight="bold")
    # High-pick AMR
    b += rect(sx + 290, sy + 30, 90, 70, fill="white", stroke=GRAY, sw=1.5, rx=4)
    b += f'<circle cx="{sx+308}" cy="{sy+108}" r="7" fill="{BLACK}"/><circle cx="{sx+362}" cy="{sy+108}" r="7" fill="{BLACK}"/>'
    b += txt(sx + 335, sy + 70, "High-pick", size=13, anchor="middle")
    b += txt(sx + 335, sy + 88, "AMR", size=13, anchor="middle", weight="bold")
    # MMF
    my = sy + 150
    b += txt(sx, my + 4, "MMF (2027) — 이동식 Rack 자동조립 로봇", size=15, weight="bold")
    b += rect(sx, my + 18, pw - 110, 60, fill=FILL, stroke=BLACK, sw=1.5, rx=4)
    b += txt(sx + 16, my + 44, "Modular 라인 구성", size=14)
    b += txt(sx + 16, my + 66, "→ 조립 라인 이동성 극대화", size=14)
    b += f'<circle cx="{sx+330}" cy="{my+84}" r="6" fill="{BLACK}"/><circle cx="{sx+370}" cy="{my+84}" r="6" fill="{BLACK}"/>'
    # blue tag
    b += rect(rxp + 40, py + ph - 70, pw - 80, 46, fill=BLUEFILL, stroke=BLUE, sw=2, rx=8)
    b += txt(rxp + pw / 2, py + ph - 40, "투자비 ▼15% · 설치공수 ▼33~46%", size=19, weight="bold", fill=BLUE, anchor="middle")

    # 하단 주의 배너
    by = py + ph + 18
    b += rect(40, by, W - 80, 40, fill="#FFF7E6", stroke=GREEN, sw=2, rx=8)
    b += txt(W / 2, by + 26, "⚠ 70%와 15%는 분모가 다름 (전체 TCO vs 공정 투자비) → 절감률 직접 비교 불가",
             size=18, weight="bold", fill=GREEN, anchor="middle")
    b += footnote("출처: 닛케이 2026-06-02 · 히타치 보도자료(PR TIMES) · 중장기 로드맵 v2 [62~69행]")
    save("fig1_개념비교", b)


# ════════════════════════════════════════════════════════════════════
# Fig 2 — 범위(분모) 차이: 왜 70% ≠ 15%
# ════════════════════════════════════════════════════════════════════
def fig2():
    b = title_bar("절감 대상 범위 비교 — 설비+설치는 공통, '건물 포함 여부'가 차이",
                  "양사 모두 설비·설치 절감 / 일본은 건물까지 포함, 당사 건물은 건설담당 별도 영역")
    bx, bw = 90, 1000
    # 비용 구성: 건물 30%, 설비 50%, 설치 20% (설명용 예시 비율)
    segs = [("건물(Building)", 0.30), ("설비(Equipment)", 0.50), ("설치(Install)", 0.20)]
    BL = "#9AB8FF"

    # Swiftfab 막대 — 건물+설비+설치 전체가 절감 대상
    y1 = 165
    b += txt(bx, y1 - 14, "일본 Swiftfab — 절감 대상 = 건물 + 설비 + 설치 (전체)", size=18, weight="bold")
    x = bx
    for name, frac in segs:
        w = bw * frac
        b += rect(x, y1, w, 64, fill=BL, stroke=BLUE, sw=1.6, rx=0)
        b += txt(x + w / 2, y1 + 40, name, size=15, anchor="middle")
        x += w
    b += f'<line x1="{bx}" y1="{y1+80}" x2="{bx+bw}" y2="{y1+80}" stroke="{BLUE}" stroke-width="2"/>'
    b += f'<line x1="{bx}" y1="{y1+72}" x2="{bx}" y2="{y1+88}" stroke="{BLUE}" stroke-width="2"/>'
    b += f'<line x1="{bx+bw}" y1="{y1+72}" x2="{bx+bw}" y2="{y1+88}" stroke="{BLUE}" stroke-width="2"/>'
    b += txt(bx + bw / 2, y1 + 108, "대상 = 총비용 전체  →  약 70% 절감 주장", size=18, weight="bold", fill=BLUE, anchor="middle")

    # LGES 막대 — 설비+설치만 절감 대상, 건물은 건설담당(범위 외)
    y2 = 410
    b += txt(bx, y2 - 14, "LGES SMC·MMF — 절감 대상 = 설비 + 설치  (건물은 건설담당 별도 영역)", size=18, weight="bold")
    x = bx
    seg_x = None
    # 빗금 패턴 정의 (건물=범위 외)
    b += '<defs><pattern id="hatch" width="8" height="8" patternTransform="rotate(45)" patternUnits="userSpaceOnUse"><line x1="0" y1="0" x2="0" y2="8" stroke="#B0B0B0" stroke-width="2"/></pattern></defs>'
    for name, frac in segs:
        w = bw * frac
        if name.startswith("건물"):
            b += rect(x, y2, w, 64, fill="url(#hatch)", stroke=LGRAY, sw=1.2, rx=0)
            b += txt(x + w / 2, y2 + 34, "건물", size=14, anchor="middle", fill=GRAY)
            b += txt(x + w / 2, y2 + 54, "건설담당(범위 외)", size=12, anchor="middle", fill=GRAY)
        else:
            b += rect(x, y2, w, 64, fill=BL, stroke=BLUE, sw=2.0, rx=0)
            b += txt(x + w / 2, y2 + 40, name, size=15, anchor="middle")
            if seg_x is None:
                seg_x = x
            seg_end = x + w
        x += w
    b += f'<line x1="{seg_x}" y1="{y2+80}" x2="{seg_end}" y2="{y2+80}" stroke="{BLUE}" stroke-width="2"/>'
    b += f'<line x1="{seg_x}" y1="{y2+72}" x2="{seg_x}" y2="{y2+88}" stroke="{BLUE}" stroke-width="2"/>'
    b += f'<line x1="{seg_end}" y1="{y2+72}" x2="{seg_end}" y2="{y2+88}" stroke="{BLUE}" stroke-width="2"/>'
    b += txt((seg_x + seg_end) / 2, y2 + 108, "대상 = 설비+설치  →  투자비 -15% (설치공수 -33~46%)",
             size=18, weight="bold", fill=BLUE, anchor="middle")
    # 건물 별도 일정 화살표
    b += txt(bx + bw * 0.15, y2 + 108, "건물: 건설 26→18개월 단축", size=14, fill=GREEN, weight="bold", anchor="middle")

    b += rect(bx, 565, bw, 44, fill="#FFF7E6", stroke=GREEN, sw=2, rx=8)
    b += txt(bx + bw / 2, 593, "∴ 공통 절감 대상 = 설비+설치. 일본 70%는 건물 포함분 / 당사 -15%는 설비+설치 기준 → '건물 포함 여부' 감안 비교",
             size=16, weight="bold", fill=GREEN, anchor="middle")
    b += footnote("※ 막대 구성비(건물30·설비50·설치20)는 설명용 예시 비율. 건물은 당사 건설담당 별도 관리(일정: 26→18개월). 출처: 분석 아카이브 §3·§4")
    save("fig2_분모차이", b)


# ════════════════════════════════════════════════════════════════════
# Fig 3 — 타임라인 비교
# ════════════════════════════════════════════════════════════════════
def fig3():
    b = title_bar("개발·실증 타임라인 — 당사 실증 선행",
                  "Swiftfab(결성→모델→가동) vs LGES(실증→확대)")
    years = [2026, 2027, 2028, 2029, 2030]
    ax0, ax1 = 110, 1150
    span = ax1 - ax0
    def xof(y):
        return ax0 + (y - 2026) / (2030 - 2026) * span
    # 축
    for lane_y in (250, 470):
        b += f'<line x1="{ax0}" y1="{lane_y}" x2="{ax1}" y2="{lane_y}" stroke="{LGRAY}" stroke-width="2"/>'
    for y in years:
        x = xof(y)
        b += f'<line x1="{x}" y1="200" x2="{x}" y2="520" stroke="#ECECEC" stroke-width="1"/>'
        b += txt(x, 180, str(y), size=18, weight="bold", anchor="middle", fill=BLACK)

    def milestone(lane_y, x, label, sub, color, up=True):
        s = f'<circle cx="{x}" cy="{lane_y}" r="11" fill="{color}" stroke="white" stroke-width="2"/>'
        ty = lane_y - 24 if up else lane_y + 38
        s += txt(x, ty, label, size=17, weight="bold", anchor="middle", fill=color)
        s += txt(x, ty + (-20 if up else 20), sub, size=13, anchor="middle", fill=GRAY)
        return s

    # Swiftfab lane (top, y=250)
    b += rect(40, 222, 60, 56, fill=HEAD, stroke=GRAY, rx=6)
    b += txt(70, 248, "Swift", size=13, weight="bold", anchor="middle")
    b += txt(70, 266, "fab", size=13, weight="bold", anchor="middle")
    b += milestone(250, xof(2026), "설립", "9개사 연합", BLACK, up=True)
    b += milestone(250, xof(2028), "모델 설비 완성", "→ 메이커 판매", BLACK, up=False)
    b += milestone(250, xof(2030), "첫 공장 가동", "목표", BLACK, up=True)

    # LGES lane (bottom, y=470)
    b += rect(40, 442, 60, 56, fill=BLUEFILL, stroke=BLUE, rx=6)
    b += txt(70, 476, "LGES", size=13, weight="bold", anchor="middle", fill=BLUE)
    b += milestone(470, xof(2026), "SMC Demo 실증", "면적+25%/공수-33%", BLUE, up=True)
    b += milestone(470, xof(2027), "MMF", "공수-46%", BLUE, up=False)
    b += milestone(470, xof(2028), "밀폐형 Cube", "전고체 대응", BLUE, up=True)

    # 시간 우위 화살표
    b += f'<defs><marker id="ah" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L8,3 L0,6 Z" fill="{GREEN}"/></marker></defs>'
    b += f'<line x1="{xof(2026)}" y1="565" x2="{xof(2030)}" y2="565" stroke="{GREEN}" stroke-width="2.5" marker-end="url(#ah)" stroke-dasharray="2 0"/>'
    b += txt((xof(2026)+xof(2030))/2, 555, "당사 실증(2026) → 일본 가동(2030): 약 4년 시간 우위",
             size=17, weight="bold", fill=GREEN, anchor="middle")
    b += footnote("출처: 닛케이/히타치 보도자료(Swiftfab 모델 2028·가동 2030) · 로드맵 v2 (SMC 2026 / MMF 2027 / 밀폐형 2028)")
    save("fig3_타임라인", b)


# ════════════════════════════════════════════════════════════════════
# Fig 4 — 당사 정량효과
# ════════════════════════════════════════════════════════════════════
def fig4():
    b = title_bar("당사 SMC·MMF 정량효과 (As-is → To-be)",
                  "모듈러·자동화 적용에 따른 면적·공수·투자비 개선")
    data = [
        ("SMC 면적 효율", 908, 1149, "+25%", True),
        ("SMC 설치 공수(MD)", 3234, 2100, "-33%", False),
        ("SMC 투자비(억)", 239.8, 203.3, "-15%", False),
        ("MMF 설치 공수(MD)", 2280, 1232, "-46%", False),
        ("MMF 투자비(억)", 296.0, 250.6, "-15%", False),
    ]
    x0 = 330
    barmax = 540
    badge_x = x0 + barmax + 150
    gmax = max(max(a, c) for _, a, c, _, _ in data)
    y = 150
    rowh = 100
    for label, asis, tobe, eff, up in data:
        b += txt(40, y + 30, label, size=18, weight="bold")
        wa = barmax * asis / gmax
        wt = barmax * tobe / gmax
        # As-is (회색)
        b += rect(x0, y, wa, 26, fill="#D9D9D9", stroke=GRAY, sw=1, rx=3)
        b += txt(x0 + wa + 8, y + 20, f"As-is {asis:,.0f}", size=13, fill=GRAY)
        # To-be (파랑)
        b += rect(x0, y + 34, wt, 26, fill="#9AB8FF", stroke=BLUE, sw=1.4, rx=3)
        b += txt(x0 + wt + 8, y + 54, f"To-be {tobe:,.0f}", size=13, fill=BLUE, weight="bold")
        # 효과 배지
        b += rect(badge_x, y + 12, 96, 36, fill=BLUEFILL, stroke=BLUE, sw=2, rx=8)
        b += txt(badge_x + 48, y + 37, eff, size=20, weight="bold", fill=BLUE, anchor="middle")
        y += rowh
    # 범례
    b += rect(40, H - 70, 18, 14, fill="#D9D9D9", stroke=GRAY)
    b += txt(64, H - 58, "As-is", size=14, fill=GRAY)
    b += rect(150, H - 70, 18, 14, fill="#9AB8FF", stroke=BLUE)
    b += txt(174, H - 58, "To-be (모듈러 적용 후)", size=14, fill=BLUE)
    b += footnote("출처: 중장기 로드맵 v2 [67~69행] · 동 로드맵 원본 [286·291행]")
    save("fig4_정량효과", b)


# ════════════════════════════════════════════════════════════════════
# Fig 5 — 공장 구축 일정 & 면적 효율 목표
# ════════════════════════════════════════════════════════════════════
def fig5():
    b = title_bar("공장 구축 일정 · 면적 효율 목표 (2026 → 2028)",
                  "건설기간 26→18개월(-31%) · 면적 효율 1,254→1,019평/GWh(-20% 목표)")

    def hpanel(y0, header, rows, maxval, unit, badge, x0=300, barmax=620, note=None):
        s = txt(40, y0, header, size=19, weight="bold")
        yy = y0 + 28
        for i, (label, val) in enumerate(rows):
            cur = (i == 0)
            last = (i == len(rows) - 1)
            w = barmax * val / maxval
            col = "#D9D9D9" if cur else ("#6E8BE8" if last else "#9AB8FF")
            stroke = GRAY if cur else BLUE
            s_ = txt(40, yy + 22, label, size=15, weight=("bold" if last else "normal"),
                     fill=(BLACK if (last or cur) else GRAY))
            s += s_
            s += rect(x0, yy, w, 30, fill=col, stroke=stroke, sw=(2.0 if last else 1.2), rx=3)
            s += txt(x0 + w + 10, yy + 22, f"{val:,} {unit}", size=15,
                     weight=("bold" if last else "normal"), fill=(BLUE if last else BLACK))
            yy += 46
        # badge
        s += rect(x0 + barmax + 150, y0 + 24, 150, 40, fill="#E8EEFF", stroke=BLUE, sw=2, rx=8)
        s += txt(x0 + barmax + 225, y0 + 51, badge, size=18, weight="bold", fill=BLUE, anchor="middle")
        if note:
            s += txt(x0 + barmax + 150, y0 + 90, note, size=13, fill=GREEN, weight="bold")
        return s

    b += hpanel(130,
                "① 공장 건설 기간 (개월) — 건설담당 목표 (건물 영역)",
                [("현재(As-is)", 26), ("2026 목표", 25), ("2027 목표", 23), ("2028 목표", 18)],
                26, "개월", "26 → 18 (-31%)",
                note="참고: Swiftfab 건물 48~72개월 → 24~36개월")
    b += f'<line x1="40" y1="395" x2="{W-40}" y2="395" stroke="#ECECEC" stroke-width="1.5"/>'
    b += hpanel(420,
                "② 면적 효율 (평/GWh) — 설비·설치 효율화, 2028년 -20% 목표",
                [("현재(As-is)", 1254), ("2026 목표", 1129), ("2027 목표", 1073), ("2028 목표", 1019)],
                1254, "평/GWh", "-20% 목표",
                note="설비+설치 효율 / 건물 제외")

    b += footnote("출처: FA기술담당 면적·건설일정 목표 (사용자 제공) · 중장기 로드맵 v2 [41행]. 절감 대상=설비+설치(건물은 건설담당 별도).")
    save("fig5_일정_면적목표", b)


if __name__ == "__main__":
    fig1(); fig2(); fig3(); fig4(); fig5()
    print("DONE — outputs/figures/")
