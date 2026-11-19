# -*- coding: utf-8 -*-
"""
3회기 PPT 생성기 (디미고 수준 편차 반영 개정판)

실행:  python tools/make-lesson03-v2.py
결과:  디미고_바이브코딩_3차시_PRD를화면으로_v2.pptx   (프로젝트 최상위에 새로 생성)

- 기존 PPT 파일은 열지도, 수정하지도 않습니다. 새 파일만 만듭니다.
- 색·글꼴·여백은 폴더 안 1~3차시 PPT에서 뽑아낸 값을 그대로 씁니다.
"""
import os
import sys

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "디미고_바이브코딩_3차시_PRD를화면으로_v2.pptx")

# ---------------------------------------------------------------- 디자인 토큰
BG        = RGBColor(0xFB, 0xFA, 0xF7)
CARD      = RGBColor(0xFF, 0xFF, 0xFF)
SOFT      = RGBColor(0xF7, 0xF4, 0xEE)
TINT      = RGBColor(0xF2, 0xED, 0xE7)
INK       = RGBColor(0x22, 0x21, 0x1E)
BODY      = RGBColor(0x4A, 0x44, 0x3E)
MUTED     = RGBColor(0x6F, 0x6A, 0x63)
FAINT     = RGBColor(0x90, 0x88, 0x7F)
ACCENT    = RGBColor(0xC6, 0x60, 0x3C)
ACCENT_D  = RGBColor(0xA8, 0x41, 0x2B)
ACCENT_L  = RGBColor(0xE9, 0xB9, 0x9F)
LINE      = RGBColor(0xE4, 0xDC, 0xD1)
LINE2     = RGBColor(0xC9, 0xC1, 0xB7)
DARK      = RGBColor(0x1E, 0x22, 0x33)
DARK_TX   = RGBColor(0xDF, 0xE3, 0xF5)
DARK_MUT  = RGBColor(0x9A, 0xA1, 0xBD)
BLUE      = RGBColor(0x8F, 0xA8, 0xF0)
GREEN     = RGBColor(0x7F, 0xD1, 0xA8)
PURPLE    = RGBColor(0xC4, 0xA0, 0xE8)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)

KR = "맑은 고딕"
MONO = "Consolas"

# 레이아웃 (inch)
SW, SH = 13.333, 7.5
ML = 0.73
CW = SW - ML * 2
HEADER_Y = 0.30
RULE_Y = 0.62
TITLE_Y = 0.84
SUB_Y = 1.52
BAR_Y = 1.96
TOP = 2.34
FOOT_Y = 6.95


# ---------------------------------------------------------------- 기본 도구
def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    return prs


def blank(prs):
    CUR["n"] += 1
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(SW), Inches(SH))
    bg.fill.solid(); bg.fill.fore_color.rgb = BG
    bg.line.fill.background(); bg.shadow.inherit = False
    return s


def rect(s, x, y, w, h, fill=None, line=None, lw=1.0, radius=None):
    shape = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    if radius:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid(); shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line; shape.line.width = Pt(lw)
    shape.shadow.inherit = False
    return shape


def text(s, x, y, w, h, parts, size=12, color=BODY, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, line=1.35, space_after=0, font=KR):
    """parts: 문자열 | [문단] ; 문단 = 문자열 | [(텍스트, {opts})...]"""
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    paras = parts if isinstance(parts, list) else [parts]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.line_spacing = line
        if space_after:
            para.space_after = Pt(space_after)
        runs = p if isinstance(p, list) else [(p, {})]
        for t, o in runs:
            r = para.add_run()
            r.text = t
            f = r.font
            f.name = o.get("font", font)
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.color.rgb = o.get("color", color)
            if o.get("space"):
                # 자간(letter-spacing) — XML 직접 설정
                r.font._rPr.set("spc", str(int(o["space"] * 100)))
    return box


def header(s, right="DIMIGO · VIBE CODING"):
    text(s, ML, HEADER_Y, 4, 0.3,
         [[("A I   ×   B U I L D", {"bold": True, "size": 10.5, "color": ACCENT, "space": 1.2})]],
         size=10.5)
    text(s, SW - ML - 5, HEADER_Y + 0.01, 5, 0.3,
         [[(right, {"bold": True, "size": 9.5, "color": FAINT, "space": 1.0})]],
         align=PP_ALIGN.RIGHT)
    rect(s, ML, RULE_Y, CW, 0.018, fill=LINE)


def footer_rule(s):
    rect(s, ML, FOOT_Y, CW, 0.018, fill=LINE)


def title_block(s, title, sub=None, bar=True):
    text(s, ML, TITLE_Y, CW, 0.62, [[(title, {"bold": True, "size": 27, "color": INK})]])
    if sub:
        text(s, ML, SUB_Y, CW, 0.34, [[(sub, {"size": 12, "color": MUTED})]])
    if bar:
        rect(s, ML, BAR_Y, 0.80, 0.045, fill=ACCENT)


def banner(s, runs, y=5.42, h=0.95, fill=TINT, align=PP_ALIGN.CENTER):
    rect(s, ML, y, CW, h, fill=fill, radius=0.12)
    text(s, ML + 0.4, y + 0.17, CW - 0.8, h - 0.3, runs,
         size=12.5, color=BODY, align=align, anchor=MSO_ANCHOR.MIDDLE, line=1.5)


def page(s, cur, total):
    text(s, SW - ML - 3, FOOT_Y + 0.12, 3, 0.26,
         [[("PART %s / %s" % (cur, total), {"size": 9, "color": FAINT, "space": 0.8})]],
         align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------- 카드 컴포넌트
# ---------------------------------------------------------------- 넘침 자동 검사
WARN = []
CUR = {"n": 0}


def _wrapped_lines(t, width_in, size_pt):
    """한글·영문이 섞인 문자열이 몇 줄로 접힐지 대략 계산한다."""
    if not t:
        return 1
    units = 0.0
    for ch in t:
        units += 1.0 if ord(ch) > 0x2000 else 0.55   # 한글 1em, 영문/숫자 0.55em
    per_line = (width_in * 72.0) / size_pt
    return max(1, int(units / per_line) + (1 if units % per_line else 0))


def check_fit(kind, x, y, w, h, cy, lines, note, size, note_size=10):
    """카드 안 내용이 실제로 들어가는지 계산해 넘치면 경고를 남긴다."""
    inner = w - (0.52 if kind == "w" else 0.60)
    used = 0.0
    for t in lines:
        n = _wrapped_lines("· " + t, inner, size)
        used += n * size * 1.42 / 72.0 + 5 / 72.0
    if note:
        used += 4 * 1.42 / 72.0
        n = _wrapped_lines(note, inner, note_size)
        used += n * note_size * 1.42 / 72.0
    used *= 1.20          # 맑은 고딕 실제 행높이가 계산값보다 크므로 안전계수
    avail = h - (cy - y) - 0.16
    if used > avail:
        WARN.append("슬라이드 %2d  %s카드(x=%.2f)  내용 %.2f in > 공간 %.2f in  (%.2f in 부족)"
                    % (CUR["n"], "흰 " if kind == "w" else "검은", x, used, avail, used - avail))


def _body_paras(lines, note, bullet_color, note_color, note_size=10):
    """본문 목록과 보조 문구를 '하나의 텍스트 흐름'으로 만든다.
       (따로 배치하면 줄 수에 따라 서로 겹치기 때문)"""
    paras = [[("· ", {"color": bullet_color}), (t, {})] for t in lines]
    if note:
        paras.append([("", {"size": 4})])          # 사이 여백 한 줄
        paras.append([(note, {"size": note_size, "color": note_color})])
    return paras


def wcard(s, x, y, w, h, topbar, label, heading, lines, note=None,
          strong=False, label_color=None, size=11.5):
    """흰 카드 + 상단 컬러바"""
    rect(s, x, y, w, h, fill=CARD, line=ACCENT_L if strong else LINE,
         lw=1.4 if strong else 1.0, radius=0.06)
    rect(s, x, y, w, 0.075, fill=topbar)
    cy = y + 0.30
    if label:
        text(s, x + 0.26, cy, w - 0.52, 0.26,
             [[(label, {"bold": True, "size": 9.5, "color": label_color or ACCENT_D, "space": 1.0})]])
        cy += 0.30
    if heading:
        text(s, x + 0.26, cy, w - 0.52, 0.4,
             [[(heading, {"bold": True, "size": 15.5, "color": INK})]])
        cy += 0.50
    rect(s, x + 0.26, cy, w - 0.52, 0.012, fill=LINE)
    cy += 0.20
    check_fit("w", x, y, w, h, cy, lines, note, size)
    text(s, x + 0.26, cy, w - 0.52, h - (cy - y) - 0.22,
         _body_paras(lines, note, ACCENT_L, FAINT),
         size=size, color=BODY, line=1.42, space_after=5)


def dcard(s, x, y, w, h, label, label_color, heading, lines, note=None, size=11.5):
    """어두운 카드 (레벨/미션용)"""
    rect(s, x, y, w, h, fill=DARK, radius=0.06)
    cy = y + 0.30
    text(s, x + 0.30, cy, w - 0.6, 0.26,
         [[(label, {"bold": True, "size": 10, "color": label_color, "space": 1.1})]])
    cy += 0.32
    text(s, x + 0.30, cy, w - 0.6, 0.42,
         [[(heading, {"bold": True, "size": 16, "color": WHITE})]])
    cy += 0.52
    rect(s, x + 0.30, cy, w - 0.6, 0.012, fill=RGBColor(0x3A, 0x40, 0x55))
    cy += 0.20
    check_fit("d", x, y, w, h, cy, lines, note, size)
    text(s, x + 0.30, cy, w - 0.6, h - (cy - y) - 0.22,
         _body_paras(lines, note, DARK_MUT, DARK_MUT),
         size=size, color=DARK_TX, line=1.42, space_after=5)


def codebox(s, x, y, w, h, title, code, tint=False):
    rect(s, x, y, w, h, fill=DARK, radius=0.05)
    text(s, x + 0.28, y + 0.22, w - 0.56, 0.26,
         [[(title, {"bold": True, "size": 9.5, "color": ACCENT_L, "space": 1.0})]])
    text(s, x + 0.28, y + 0.60, w - 0.56, h - 0.85,
         [[(l, {})] for l in code],
         size=10.5, color=DARK_TX, font=MONO, line=1.55)


def step_row(s, y, items, h=1.5):
    """번호 + 제목 + 설명이 가로로 이어지는 흐름"""
    n = len(items)
    gap = 0.20
    w = (CW - gap * (n - 1)) / n
    for i, (num, head, desc) in enumerate(items):
        x = ML + i * (w + gap)
        rect(s, x, y, w, h, fill=CARD, line=LINE, radius=0.06)
        text(s, x + 0.24, y + 0.22, w - 0.48, 0.3,
             [[(num, {"bold": True, "size": 11, "color": ACCENT, "space": 0.8})]])
        text(s, x + 0.24, y + 0.56, w - 0.48, 0.34,
             [[(head, {"bold": True, "size": 14, "color": INK})]])
        text(s, x + 0.24, y + 0.96, w - 0.48, h - 1.05,
             [[(desc, {"size": 11, "color": MUTED})]], line=1.45)
        if i < n - 1:
            text(s, x + w + 0.02, y + h / 2 - 0.16, gap - 0.04, 0.3,
                 [[("→", {"size": 13, "color": ACCENT_L})]], align=PP_ALIGN.CENTER)


def part_slide(s, kicker, title, sub, cur, total):
    header(s)
    text(s, ML, 2.55, CW, 0.34,
         [[(kicker, {"bold": True, "size": 11, "color": ACCENT, "space": 1.6})]])
    text(s, ML, 2.98, CW, 0.85, [[(title, {"bold": True, "size": 38, "color": INK})]])
    rect(s, ML, 3.98, 0.80, 0.045, fill=ACCENT)
    text(s, ML, 4.22, CW * 0.8, 0.4, [[(sub, {"size": 13.5, "color": MUTED})]])
    footer_rule(s)
    page(s, cur, total)


# ================================================================ 슬라이드
def build():
    prs = new_deck()
    S = lambda: blank(prs)

    # ---------------------------------------------------------- 1. 표지
    s = S()
    text(s, ML, 0.44, 5, 0.3,
         [[("A I   ×   B U I L D", {"bold": True, "size": 11, "color": ACCENT, "space": 1.3})]])
    text(s, SW - ML - 7.4, 0.30, 7.4, 0.7,
         [[("AI 노코드 창작 클래스", {"bold": True, "size": 30, "color": RGBColor(0xC9, 0xC1, 0xB7)})]],
         align=PP_ALIGN.RIGHT)
    rect(s, ML, 1.06, CW, 0.018, fill=LINE)

    steps = [("1", "IDEA"), ("2", "SPEC"), ("3", "BUILD"), ("4", "TEST")]
    cx = ML + 0.28
    for i, (num, nm) in enumerate(steps):
        on = (i == 2)
        d = 0.62
        c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx), Inches(1.48), Inches(d), Inches(d))
        c.fill.solid(); c.fill.fore_color.rgb = ACCENT if on else RGBColor(0xEF, 0xE5, 0xDC)
        c.line.fill.background(); c.shadow.inherit = False
        text(s, cx, 1.62, d, 0.3, [[(num, {"bold": True, "size": 12,
             "color": WHITE if on else ACCENT_D})]], align=PP_ALIGN.CENTER)
        text(s, cx - 0.2, 2.22, d + 0.4, 0.26,
             [[(nm, {"bold": True, "size": 8.5, "color": ACCENT_D if on else FAINT, "space": 0.9})]],
             align=PP_ALIGN.CENTER)
        if i < 3:
            rect(s, cx + d, 1.77, 0.42, 0.012, fill=LINE2)
        cx += d + 0.42

    text(s, ML, 2.92, 4, 1.0, [[("03", {"bold": True, "size": 58, "color": INK})]])
    text(s, ML, 4.00, 9.5, 1.3,
         [[("PRD를 화면으로", {"bold": True, "size": 36, "color": INK})],
          [("내 기획서가 내 화면을 결정합니다", {"size": 18, "color": BODY})]], line=1.35)
    rect(s, ML, 5.56, 1.5, 0.05, fill=ACCENT)
    text(s, ML, 5.86, 9, 0.32,
         [[("90분 · 오늘의 산출물 : 작동하는 index.html + 내 PRD 검증표", {"size": 12.5, "color": MUTED})]])
    text(s, SW - ML - 7, 6.62, 7, 0.3,
         [[("한국디지털미디어고등학교 · Vibe Coding Project", {"size": 10.5, "color": FAINT})]],
         align=PP_ALIGN.RIGHT)

    # ---------------------------------------------------------- 2. 오늘의 목표
    s = S(); header(s)
    title_block(s, "오늘의 목표",
                "2회기에 쓴 PRD를 실제로 동작하는 첫 화면으로 바꾸고, 그 결과를 내 기준으로 검증합니다.")
    items = [
        (BLUE, "01", "TRANSLATE", "기획서를 화면으로",
         ["MVP 기능 → 화면 요소", "User Flow → 배치 순서", "Non-goal → 만들지 말 것"]),
        (GREEN, "02", "BUILD", "첫 화면을 만든다",
         ["내 PRD로 요청문을 쓴다", "코드를 받아 파일로 저장", "브라우저로 직접 연다"]),
        (PURPLE, "03", "VERIFY", "내 기준으로 검증한다",
         ["완성 기준 5문장으로 판정", "통과·실패를 구분한다", "AI가 덧붙인 것을 찾는다"]),
        (ACCENT_L, "04", "ITERATE", "불일치를 고친다",
         ["실패한 항목부터 수정", "빈 상태·예외 화면 추가", "고칠 때마다 다시 확인"]),
    ]
    w = (CW - 0.22 * 3) / 4
    for i, (col, num, lab, head, lines) in enumerate(items):
        x = ML + i * (w + 0.22)
        wcard(s, x, TOP, w, 3.0, col, num + " · " + lab, head, lines)
    banner(s, [[("오늘은 같은 주제로 시작하지만, ", {}),
                ("30명이 30개의 다른 화면", {"bold": True, "color": ACCENT_D}),
                ("을 만듭니다. 그 차이를 만드는 것이 PRD입니다.", {})]])
    footer_rule(s)

    # ---------------------------------------------------------- 3. 90분 흐름
    s = S(); header(s)
    title_block(s, "오늘 90분의 흐름", "설명 20분, 직접 만들고 검증하는 시간이 70분입니다.")
    flow = [
        ("00-06", "OPEN", "같은 주제\n다른 결과", False),
        ("06-20", "TRANSLATE", "PRD를 화면으로\n요청문 설계", False),
        ("20-45", "BUILD", "첫 화면 만들어\n저장하고 열기", True),
        ("45-58", "VERIFY", "내 PRD로\n직접 검증하기", True),
        ("58-76", "FIX", "불일치 수정\n예외 화면 추가", True),
        ("76-90", "SHARE", "결과 비교\n4회기 예고", False),
    ]
    n = len(flow); gap = 0.16
    w = (CW - gap * (n - 1)) / n
    for i, (t, lab, desc, hot) in enumerate(flow):
        x = ML + i * (w + gap)
        rect(s, x, TOP, w, 2.62, fill=CARD, line=ACCENT_L if hot else LINE,
             lw=1.4 if hot else 1.0, radius=0.06)
        rect(s, x, TOP, w, 0.075, fill=ACCENT if hot else LINE2)
        text(s, x + 0.20, TOP + 0.30, w - 0.4, 0.36,
             [[(t, {"bold": True, "size": 16, "color": INK})]])
        text(s, x + 0.20, TOP + 0.74, w - 0.4, 0.26,
             [[(lab, {"bold": True, "size": 9.5,
                      "color": ACCENT_D if hot else FAINT, "space": 1.0})]])
        rect(s, x + 0.20, TOP + 1.06, w - 0.4, 0.012, fill=LINE)
        text(s, x + 0.20, TOP + 1.24, w - 0.4, 1.2,
             [[(l, {})] for l in desc.split("\n")],
             size=11, color=BODY, line=1.5)
    banner(s, [[("지금 준비하세요 — 2회기에 만든 ", {}),
                ("PRD.md", {"bold": True, "font": MONO, "color": ACCENT_D}),
                (" 열기 · 바탕화면에 영어 이름 폴더 하나 · Claude 로그인", {})],
               [("PRD를 잃어버렸다면 지금 손을 드세요. 예시 PRD를 드립니다.", {"size": 11, "color": MUTED})]],
           y=5.30, h=1.05)
    footer_rule(s)

    # ---------------------------------------------------------- 4. PART 0
    s = S(); part_slide(s, "PART 0", "같은 주제, 다른 결과",
                        "오늘 교실에 30개의 서로 다른 화면이 생깁니다. 왜 그럴까요?", 0, 4)

    # ---------------------------------------------------------- 5. 오늘의 질문
    s = S(); header(s)
    title_block(s, "오늘의 질문 — 같은 주제인데 왜 다른 앱이 나올까?",
                "2회기에 우리는 모두 같은 주제로 기획했습니다. 그런데 결과는 전부 다릅니다.")
    w = (CW - 0.26 * 2) / 3
    wcard(s, ML, TOP, w, 2.90, LINE2, "같은 것", "주제와 도구",
          ["똑같은 주제로 시작했다", "똑같이 Claude를 쓴다", "똑같이 90분을 쓴다"],
          note="출발점은 모두 같습니다.")
    wcard(s, ML + w + 0.26, TOP, w, 2.90, ACCENT, "다른 것", "내가 내린 결정",
          ["어떤 문제를 고를지", "누구를 첫 사용자로 둘지", "무엇을 MVP에 넣고 뺄지"],
          strong=True, note="이것이 PRD에 적혀 있습니다.")
    wcard(s, ML + (w + 0.26) * 2, TOP, w, 2.90, LINE2, "결과", "화면",
          ["화면에 놓이는 요소", "배치 순서", "만들지 않은 기능"],
          note="오늘 눈으로 확인합니다.")
    banner(s, [[("AI는 ", {}), ("내 기획서 수준만큼", {"bold": True, "color": ACCENT_D}),
                (" 만들어 줍니다. 오늘 나올 화면의 차이는 실력 차이가 아니라 ", {}),
                ("결정의 차이", {"bold": True, "color": ACCENT_D}), ("입니다.", {})]],
           y=5.48, h=0.88)
    footer_rule(s)

    # ---------------------------------------------------------- 6. 오늘 하는 것 / 안 하는 것
    s = S(); header(s)
    title_block(s, "오늘 하는 것과 하지 않는 것",
                "범위를 먼저 좁혀 둡니다. 오늘 다 만들려고 하면 아무것도 끝나지 않습니다.")
    hw = (CW - 0.3) / 2
    wcard(s, ML, TOP, hw, 3.25, ACCENT, "오늘 한다", "보이는 것까지",
          ["내 PRD의 MVP 기능이 화면에 보인다",
           "사용자 흐름 순서대로 배치된다",
           "빈 화면·예외 상태를 만든다",
           "내 완성 기준으로 통과·실패를 판정한다"], strong=True)
    wcard(s, ML + hw + 0.3, TOP, hw, 3.25, LINE2, "오늘 하지 않는다", "동작과 저장",
          ["버튼을 눌러도 데이터가 쌓이지 않는다",
           "새로고침하면 입력이 사라진다",
           "로그인·서버·알림은 만들지 않는다",
           "예쁘게 다듬는 데 시간을 쓰지 않는다"],
          note="버튼이 안 눌리는 것은 고장이 아닙니다.")
    banner(s, [[("1회기에 배운 4단계 — ", {}),
                ("뼈대 → 기능 → 디자인 → 검증", {"bold": True, "color": ACCENT_D}),
                (". 오늘은 첫 번째 뼈대와 네 번째 검증을 함께 합니다.", {})]],
           y=5.76, h=0.74)
    footer_rule(s)

    # ---------------------------------------------------------- 7. PART 1
    s = S(); part_slide(s, "PART 1", "PRD를 화면으로",
                        "기획서의 어느 줄이 화면의 어느 부분이 되는지 · 규칙은 세 개뿐입니다.", 1, 4)

    # ---------------------------------------------------------- 8~10. 번역 규칙
    s = S(); header(s)
    title_block(s, "번역 규칙 ① — MVP 기능은 화면 요소가 된다",
                "기능 하나가 화면 요소 하나 이상을 만듭니다. 기능 목록이 곧 화면 목록입니다.")
    rect(s, ML, TOP, CW, 2.35, fill=CARD, line=LINE, radius=0.04)
    rows = [
        ("PRD 4번 · MVP 기능", "화면에 생기는 것", True),
        ("과제 등록", "입력칸 + 날짜 선택 + 등록 버튼", False),
        ("D-day 목록", "목록 영역 + 항목마다 남은 일수 표시", False),
        ("완료 체크", "항목마다 체크박스", False),
    ]
    ry = TOP + 0.22
    for i, (a, b, head) in enumerate(rows):
        col = FAINT if head else BODY
        sz = 10 if head else 13
        text(s, ML + 0.38, ry, 4.4, 0.34,
             [[(a, {"bold": True, "size": sz, "color": ACCENT_D if head else INK,
                    "space": 1.0 if head else 0})]])
        if not head:
            text(s, ML + 4.9, ry + 0.02, 0.5, 0.3, [[("→", {"size": 13, "color": ACCENT_L})]])
        text(s, ML + 5.5, ry, CW - 5.9, 0.34,
             [[(b, {"size": sz, "bold": head, "color": ACCENT_D if head else col,
                    "space": 1.0 if head else 0})]])
        ry += 0.30 if head else 0.60
        if i == 0:
            rect(s, ML + 0.38, ry - 0.14, CW - 0.76, 0.012, fill=LINE)
    banner(s, [[("내 PRD 4번에 기능이 5개라면 화면 요소도 5덩어리입니다. ", {}),
                ("6개째를 지금 추가하지 마세요.", {"bold": True, "color": ACCENT_D})]])
    footer_rule(s)

    s = S(); header(s)
    title_block(s, "번역 규칙 ② — User Flow는 배치 순서가 된다",
                "사용자가 먼저 하는 일이 화면 위쪽에 옵니다. 순서를 정해 주지 않으면 AI가 마음대로 놓습니다.")
    step_row(s, TOP, [
        ("01 · 연다", "가장 먼저 보이는 것", "앱을 열자마자 눈에 들어와야 하는 것을 맨 위에"),
        ("02 · 입력한다", "입력 영역", "등록·작성처럼 사용자가 값을 넣는 자리"),
        ("03 · 확인한다", "목록 영역", "입력한 결과가 쌓여 보이는 자리"),
        ("04 · 정리한다", "항목별 동작", "완료·삭제처럼 항목마다 붙는 것"),
    ], h=1.9)
    text(s, ML, TOP + 2.15, CW, 0.9,
         [[("내 PRD 5번(사용자 흐름)을 그대로 순서로 쓰세요. ", {}),
           ("\"위에서부터 이 순서로 배치해줘\"", {"bold": True, "font": MONO, "color": ACCENT_D}),
           (" 한 줄이면 됩니다.", {})]], size=12.5, color=BODY)
    banner(s, [[("화면 이름을 명사로 적으면(\"메인 화면\") 무엇을 할 수 있는지가 사라집니다. ", {}),
                ("동사로 적으세요.", {"bold": True, "color": ACCENT_D})]])
    footer_rule(s)

    s = S(); header(s)
    title_block(s, "번역 규칙 ③ — Non-goal은 \"만들지 말 것\"이 된다",
                "오늘 가장 중요한 규칙입니다. 적지 않으면 AI가 친절하게 덧붙입니다.")
    hw = (CW - 0.3) / 2
    wcard(s, ML, TOP, hw, 3.25, LINE2, "적지 않으면", "코드가 세 배가 된다",
          ["묻지 않은 로그인 화면이 생긴다",
           "설정·다크모드가 딸려 온다",
           "코드가 길어져 복사도 수정도 어려워진다",
           "내가 읽고 판단할 수 없는 코드가 된다"],
          note="AI는 불친절해서가 아니라 친절해서 이렇게 합니다.")
    wcard(s, ML + hw + 0.3, TOP, hw, 3.25, ACCENT, "적어 두면", "딱 필요한 만큼만",
          ["요청한 기능만 화면에 생긴다",
           "코드가 짧아 읽을 수 있다",
           "고칠 곳을 바로 찾는다",
           "오늘 안에 끝난다"], strong=True,
          note="PRD 7번을 그대로 프롬프트에 붙여 넣으세요.")
    banner(s, [[("요청문에 이 한 줄을 꼭 넣으세요 — ", {}),
                ("\"Non-goal에 적은 것은 절대 만들지 마\"", {"bold": True, "font": MONO, "color": ACCENT_D})]],
           y=5.76, h=0.74)
    footer_rule(s)

    # ---------------------------------------------------------- 11. 내 PRD에서 오늘 쓰는 것
    s = S(); header(s)
    title_block(s, "내 PRD에서 오늘 꺼내 쓰는 항목",
                "8개 중 5개만 씁니다. 나머지는 버리는 게 아니라 아껴 둡니다.")
    use = [("2", "사용자"), ("4", "MVP 기능"), ("5", "사용자 흐름"), ("7", "만들지 않을 것"), ("8", "완성 기준")]
    keep = [("1", "문제 · 기획 검토용"), ("3", "가치 · 기획 검토용"), ("6", "예외 상황 · 오늘 일부만")]
    hw = (CW - 0.3) / 2
    rect(s, ML, TOP, hw, 2.75, fill=CARD, line=ACCENT_L, lw=1.4, radius=0.05)
    rect(s, ML, TOP, hw, 0.075, fill=ACCENT)
    text(s, ML + 0.3, TOP + 0.30, hw - 0.6, 0.3,
         [[("오늘 쓰는 5개", {"bold": True, "size": 11, "color": ACCENT_D, "space": 1.0})]])
    yy = TOP + 0.72
    for num, nm in use:
        c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(ML + 0.32), Inches(yy - 0.02), Inches(0.3), Inches(0.3))
        c.fill.solid(); c.fill.fore_color.rgb = ACCENT; c.line.fill.background(); c.shadow.inherit = False
        text(s, ML + 0.32, yy + 0.04, 0.3, 0.24, [[(num, {"bold": True, "size": 10, "color": WHITE})]],
             align=PP_ALIGN.CENTER)
        text(s, ML + 0.78, yy + 0.02, hw - 1.1, 0.3, [[(nm, {"size": 13, "color": INK, "bold": True})]])
        yy += 0.40
    rect(s, ML + hw + 0.3, TOP, hw, 2.75, fill=SOFT, line=LINE, radius=0.05)
    rect(s, ML + hw + 0.3, TOP, hw, 0.075, fill=LINE2)
    text(s, ML + hw + 0.6, TOP + 0.30, hw - 0.6, 0.3,
         [[("오늘 아껴 두는 3개", {"bold": True, "size": 11, "color": FAINT, "space": 1.0})]])
    yy = TOP + 0.72
    for num, nm in keep:
        text(s, ML + hw + 0.62, yy, hw - 0.9, 0.3,
             [[(num + "  ", {"bold": True, "size": 12, "color": LINE2}), (nm, {"size": 12, "color": MUTED})]])
        yy += 0.42
    text(s, ML + hw + 0.62, yy + 0.12, hw - 0.9, 0.7,
         [[("6번은 오늘 \"빈 화면\" 하나만 씁니다.",
            {"size": 10.5, "color": FAINT})]], line=1.45)
    banner(s, [[("8번(완성 기준)이 오늘의 숨은 주인공입니다. ", {}),
                ("만드는 것보다 판정하는 데 시간을 더 씁니다.", {"bold": True, "color": ACCENT_D})]])
    footer_rule(s)

    # ---------------------------------------------------------- 12. PART 2
    s = S(); part_slide(s, "PART 2", "요청문과 첫 화면",
                        "출발선이 다르므로 · 요청문도 각자 다르게 씁니다.", 2, 4)

    # ---------------------------------------------------------- 13. 트랙 선택 ★
    s = S(); header(s)
    title_block(s, "오늘의 트랙을 스스로 고르세요",
                "산출물은 모두 같습니다. 가는 길만 다릅니다. 중간에 올려도 내려도 됩니다.")
    w = (CW - 0.26 * 2) / 3
    dcard(s, ML, TOP, w, 2.95, "TRACK A", BLUE, "따라 만들기",
          ["오늘 처음 만들어 본다", "제공된 요청문을 그대로 쓴다", "화면이 뜨는 것이 목표"],
          note="여기까지만 해도 오늘 목표는 100% 달성입니다.")
    dcard(s, ML + w + 0.26, TOP, w, 2.95, "TRACK B", GREEN, "직접 설계하기",
          ["요청문을 내가 쓴다", "데이터 구조까지 지정한다", "빈 화면·예외까지 만든다"],
          note="HTML을 조금 아는 학생에게 권합니다.")
    dcard(s, ML + (w + 0.26) * 2, TOP, w, 2.95, "TRACK C", PURPLE, "따져 묻기",
          ["영어로 제약을 걸어 요청한다", "받은 코드를 리뷰한다", "더 적은 프롬프트로 같은 결과"],
          note="많이 만들기가 아니라 정확히 만들기입니다.")
    banner(s, [[("빨리 끝내는 것이 잘한 것이 아닙니다. ", {}),
                ("내 PRD대로 나왔는지 설명할 수 있는 것", {"bold": True, "color": ACCENT_D}),
                ("이 잘한 것입니다.", {})]])
    footer_rule(s)

    # ---------------------------------------------------------- 14. 요청 프롬프트 A
    s = S(); header(s)
    title_block(s, "요청문 · TRACK A — 이대로 쓰면 됩니다",
                "아래를 복사해 PRD 내용을 끼워 넣고 보내세요. 조건 다섯 줄이 오늘의 핵심입니다.")
    codebox(s, ML, TOP, CW * 0.62, 3.30, "실습 ① 요청문", [
        "아래는 내가 만든 PRD야. 이걸 보고 첫 화면을 만들어줘.",
        "",
        "[ 여기에 PRD.md 내용을 통째로 붙여넣기 ]",
        "",
        "조건:",
        "- HTML, CSS, JavaScript를 한 파일에 넣어줘",
        "- 미리보기 말고 코드로 보여줘",
        "- MVP(4번)에 적은 기능만. Non-goal(7번)은 절대 만들지 마",
        "- 사용자 흐름(5번) 순서대로 위에서부터 배치해줘",
        "- 지금은 화면과 디자인만. 기능은 아직 연결하지 마",
    ])
    bx = ML + CW * 0.62 + 0.3
    bw = CW - CW * 0.62 - 0.3
    wcard(s, bx, TOP, bw, 3.30, ACCENT, "이 줄이 없으면", "이렇게 됩니다",
          ["한 파일이 아니면 → 저장을 세 번 해야 한다",
           "코드로 안 받으면 → 내 컴퓨터에 남지 않는다",
           "Non-goal을 안 막으면 → 로그인이 딸려 온다",
           "순서를 안 주면 → 내 흐름과 다르게 배치된다",
           "기능까지 만들면 → 코드가 길어 읽을 수 없다"], strong=True)
    footer_rule(s)

    # ---------------------------------------------------------- 15. 요청 프롬프트 B·C
    s = S(); header(s)
    title_block(s, "요청문 · TRACK B와 C — 조건을 직접 설계합니다",
                "정답을 주지 않습니다. 무엇을 지정해야 결과가 흔들리지 않는지 스스로 고르세요.")
    hw = (CW - 0.3) / 2
    codebox(s, ML, TOP, hw, 2.9, "TRACK B · 직접 설계 (한국어)", [
        "[PRD 붙여넣기]",
        "",
        "추가 조건:",
        "- 데이터 한 건은 다음 필드로 다룬다:",
        "  id / 제목 / 날짜 / 완료여부",
        "- 목록이 0건일 때 보여줄 안내 화면도 만들어줘",
        "- 제목이 아주 길 때 줄이 깨지지 않게 해줘",
        "- 모바일 390px 기준으로 먼저 설계해줘",
        "- 설명은 빼고 코드만 줘",
    ])
    codebox(s, ML + hw + 0.3, TOP, hw, 2.9, "TRACK C · 제약 설계 (영어)", [
        "[Paste my PRD]",
        "",
        "Constraints:",
        "- Single HTML file, no framework, no CDN",
        "- Semantic HTML5 + WCAG AA contrast",
        "- Mobile-first at 390px, no horizontal scroll",
        "- Build ONLY the MVP list. Respect Non-goals.",
        "- Empty state + long-text edge case required",
        "- Output code only. No explanation.",
    ])
    banner(s, [[("TRACK C 추가 미션 — 받은 코드를 그대로 믿지 말고 이렇게 되물으세요: ", {}),
                ("\"이 코드에서 시맨틱·접근성 문제를 찾아 근거와 함께 알려줘. 고치지는 마.\"",
                 {"bold": True, "font": MONO, "color": ACCENT_D, "size": 11.5})]], y=5.45, h=0.9)
    footer_rule(s)

    # ---------------------------------------------------------- 16. 실습 ①
    s = S(); header(s)
    title_block(s, "실습 ① 첫 화면 만들기 — 25분",
                "예쁘게 만들려 하지 마세요. 화면이 뜨고, 내 파일로 남는 것까지가 목표입니다.")
    step_row(s, TOP, [
        ("STEP 1 · 0-5분", "보낸다", "PRD 전문과 조건을 함께 붙여넣고 요청한다"),
        ("STEP 2 · 5-12분", "받는다", "코드 블록 전체를 복사한다. 맨 끝 </html>까지"),
        ("STEP 3 · 12-18분", "저장한다", "폴더 안에 index.html 이라는 이름으로 저장한다"),
        ("STEP 4 · 18-25분", "연다", "더블클릭해서 브라우저로 열고 화면을 확인한다"),
    ], h=1.85)
    banner(s, [[("20분이 되면 다 같이 멈춥니다. ", {}),
                ("화면이 뜬 사람은 손을 들어 주세요.", {"bold": True, "color": ACCENT_D}),
                ("  안 뜬 사람은 그때 함께 해결합니다. 혼자 붙잡고 있지 마세요.", {})]],
           y=5.05, h=0.95)
    text(s, ML, 6.15, CW, 0.5,
         [[("TRACK C는 저장까지 5분 안에 끝날 수 있습니다. 남는 시간은 다음 장의 코드 리뷰에 쓰세요.",
            {"size": 11, "color": FAINT})]], align=PP_ALIGN.CENTER)
    footer_rule(s)

    # ---------------------------------------------------------- 17. 저장 확인 + 문제 해결
    s = S(); header(s)
    title_block(s, "막히면 여기부터 — 저장과 실행에서 나는 문제 전부",
                "오늘 생기는 문제는 십중팔구 이 넷 중 하나입니다. 손 들기 전에 먼저 확인하세요.")
    w = (CW - 0.24 * 3) / 4
    probs = [
        ("확장자", "index.html", ["메모장이 뒤에 .txt를 붙인다", "저장 창의 파일 형식을", "\"모든 파일\"로 바꾼다"]),
        ("화면이 하얗다", "복사가 잘렸다", ["코드 맨 아래가", "</html> 로 끝나는지 확인", "다시 전체 복사"]),
        ("한글이 깨진다", "인코딩", ["저장할 때 인코딩을", "UTF-8 로 선택", "또는 VS Code로 저장"]),
        ("버튼이 안 눌린다", "정상입니다", ["오늘은 기능을 연결하지", "않았습니다. 고장이 아닙니다", "6회기에 전기를 넣습니다"]),
    ]
    for i, (lab, head, lines) in enumerate(probs):
        x = ML + i * (w + 0.24)
        wcard(s, x, TOP, w, 3.15, LINE2 if i != 3 else ACCENT_L, lab, head, lines)
    banner(s, [[("VS Code를 쓸 줄 아는 학생은 메모장 대신 VS Code로 저장하세요. ", {}),
                ("확장자·인코딩 사고가 아예 생기지 않습니다.", {"bold": True, "color": ACCENT_D})]],
           y=5.66, h=0.78)
    footer_rule(s)

    # ---------------------------------------------------------- 18. PART 3
    s = S(); part_slide(s, "PART 3", "내 기준으로 검증하기",
                        "만드는 것은 AI가 했습니다 · 판정하는 것은 사람이 합니다.", 3, 4)

    # ---------------------------------------------------------- 19. 실습 ② 검증 ★
    s = S(); header(s)
    title_block(s, "실습 ② 내 PRD로 직접 검증하기 — 13분",
                "PRD 8번(완성 기준)에 쓴 문장을 그대로 체크리스트로 씁니다. 통과와 실패를 구분하세요.")
    rect(s, ML, TOP, CW * 0.60, 3.50, fill=CARD, line=ACCENT_L, lw=1.4, radius=0.05)
    rect(s, ML, TOP, CW * 0.60, 0.075, fill=ACCENT)
    text(s, ML + 0.32, TOP + 0.30, CW * 0.6 - 0.64, 0.3,
         [[("검증표 · 내 완성 기준 5문장", {"bold": True, "size": 11, "color": ACCENT_D, "space": 1.0})]])
    checks = [
        "과제명이 비어 있으면 저장되지 않는다",
        "항목이 0건이면 안내 화면이 나타난다",
        "마감이 임박한 항목이 위에 온다",
        "지난 날짜는 다르게 표시된다",
        "모바일 390px에서 가로 스크롤이 없다",
    ]
    yy = TOP + 0.76
    for i, ct in enumerate(checks):
        b = rect(s, ML + 0.34, yy, 0.26, 0.26, fill=None, line=LINE2, lw=1.2)
        text(s, ML + 0.76, yy - 0.01, CW * 0.6 - 1.9, 0.3,
             [[(ct, {"size": 12, "color": BODY})]])
        text(s, ML + CW * 0.6 - 1.05, yy - 0.01, 0.9, 0.3,
             [[("통과 / 실패", {"size": 9.5, "color": FAINT})]], align=PP_ALIGN.RIGHT)
        yy += 0.40
    bx = ML + CW * 0.60 + 0.3
    bw = CW - CW * 0.60 - 0.3
    wcard(s, bx, TOP, bw, 3.50, ACCENT, "판정하는 법", "누가 눌러도 같은 결론",
          ["실제로 눌러 보고 판정한다",
           "\"대충 되는 것 같다\"는 판정이 아니다",
           "실패한 항목이 오늘 고칠 목록이다",
           "판정할 수 없는 문장이 있다면\n그 문장이 잘못 쓰인 것이다"], strong=True,
          note="그렇다면 PRD를 고칠 차례입니다.")
    banner(s, [[("5개 중 몇 개가 통과했는지 적어 두세요. ", {}),
                ("9회기에 이 표를 다시 씁니다.", {"bold": True, "color": ACCENT_D})]],
           y=5.96, h=0.70)
    footer_rule(s)

    # ---------------------------------------------------------- 20. AI의 친절함 탐지
    s = S(); header(s)
    title_block(s, "그리고 하나 더 — AI가 몰래 덧붙인 것 찾기",
                "실패한 항목만 보면 절반입니다. 시키지 않았는데 생긴 것도 찾아야 합니다.")
    hw = (CW - 0.3) / 2
    wcard(s, ML, TOP, hw, 3.25, ACCENT, "찾아볼 것", "내 PRD에 없는데 화면에 있는 것",
          ["PRD 7번(Non-goal)에 적은 기능이 있는가",
           "묻지 않은 버튼·메뉴·설정이 있는가",
           "내가 정하지 않은 문구가 있는가",
           "화면에 있는데 PRD엔 없는 결정은 무엇인가"], strong=True)
    wcard(s, ML + hw + 0.3, TOP, hw, 3.25, LINE2, "찾았다면", "둘 중 하나를 고른다",
          ["① 필요 없다 → 지워 달라고 요청한다",
           "② 있으면 좋겠다 → PRD에 먼저 적는다",
           "그냥 두는 선택지는 없습니다",
           "기록하지 않은 결정은 다음 시간에 사라집니다"],
          note="AI가 대신 내린 결정을 되돌리는 작업입니다.")
    banner(s, [[("1회기에 배운 ", {}), ("환각·맥락 유실", {"bold": True, "color": ACCENT_D}),
                ("이 실제로 어떻게 나타나는지 오늘 처음 눈으로 보게 됩니다.", {})]],
           y=5.76, h=0.74)
    footer_rule(s)

    # ---------------------------------------------------------- 21. 수정 요청 공식
    s = S(); header(s)
    title_block(s, "고칠 때는 세 조각으로 — 어디를 · 무엇을 · 어떻게",
                "세 조각이 다 있으면 한 번에 고쳐집니다. 하나라도 빠지면 멀쩡한 곳까지 바뀝니다.")
    hw = (CW - 0.3) / 2
    codebox(s, ML, TOP, hw, 3.30, "한 번에 안 고쳐지는 요청", [
        "✕  좀 더 예쁘게 해줘",
        "✕  이상해요",
        "✕  안 되는데요",
        "✕  다시 만들어줘",
        "",
        "→ 마음에 들던 부분까지 전부 바뀐다",
    ])
    wcard(s, ML + hw + 0.3, TOP, hw, 3.30, ACCENT, "한 번에 고쳐지는 요청", "어디를 · 무엇을 · 어떻게",
          ["\"목록 영역에서(어디를)",
           "  마감일 표시를(무엇을)",
           "  D-3 형식으로 바꿔줘(어떻게)\"",
           "고쳤더니 다른 데가 망가졌다면 —",
           "\"방금 바꾼 것만 되돌리고 나머지는 그대로 둬\""], strong=True)
    footer_rule(s)

    # ---------------------------------------------------------- 22. 실습 ③ 레벨
    s = S(); header(s)
    title_block(s, "실습 ③ 고치기 — 18분 · 위에서부터 순서대로",
                "레벨 1은 전원 필수입니다. 2·3은 여유가 있으면, 4는 TRACK C 학생에게.")
    w = (CW - 0.22 * 3) / 4
    dcard(s, ML, TOP, w, 3.20, "LEVEL 1 · 필수", BLUE, "불일치 고치기",
          ["검증표에서 실패한 항목", "하나씩 고친다", "고칠 때마다 저장 후 새로고침"],
          note="여기까지가 오늘의 합격선입니다.")
    dcard(s, ML + w + 0.22, TOP, w, 3.20, "LEVEL 2", GREEN, "빈 화면 만들기",
          ["항목이 0건일 때", "무엇을 보여줄지 정한다", "안내 문구와 다음 행동"],
          note="2회기 6번 예외 상황을 꺼내 쓰세요.")
    dcard(s, ML + (w + 0.22) * 2, TOP, w, 3.20, "LEVEL 3", PURPLE, "경계값 처리",
          ["제목이 아주 길 때", "지난 날짜일 때", "항목이 50개일 때"],
          note="정상 상황만 만들면 실제로 못 씁니다.")
    dcard(s, ML + (w + 0.22) * 3, TOP, w, 3.20, "LEVEL 4 · TRACK C", ACCENT_L, "코드 리뷰",
          ["시맨틱 태그로 바꾸기", "대비·초점 표시 점검", "중복 코드 줄이기"],
          note="\"기능은 그대로 두고 구조만 정리해줘\"")
    banner(s, [[("레벨을 건너뛰지 마세요. ", {}),
                ("검증표가 다 통과되기 전에 새 기능을 추가하면 오늘 안에 끝나지 않습니다.",
                 {"bold": True, "color": ACCENT_D})]],
           y=5.70, h=0.78)
    footer_rule(s)

    # ---------------------------------------------------------- 23. 먼저 끝났다면
    s = S(); header(s)
    title_block(s, "먼저 끝났다면 — 더 만들지 말고, 더 정확하게",
                "빨리 끝낸 학생을 위한 과제입니다. 기능을 늘리는 방향이 아닙니다.")
    w = (CW - 0.26 * 2) / 3
    wcard(s, ML, TOP, w, 3.20, ACCENT, "① 줄이기", "더 적은 프롬프트로",
          ["지금 결과를 처음부터 다시",
           "단 한 번의 요청으로 만들어 본다",
           "무엇을 빠뜨렸는지 비교한다"], strong=True,
          note="좋은 명세는 짧고 정확합니다.")
    wcard(s, ML + w + 0.26, TOP, w, 3.20, ACCENT, "② 되묻기", "AI를 검토자로",
          ["\"이 PRD로 구현할 때",
           "  문제가 될 부분을 질문 형태로 알려줘\"",
           "코드는 만들지 말라고 못 박는다"], strong=True,
          note="2회기 PRD 검토를 코드 단계에서 다시.")
    wcard(s, ML + (w + 0.26) * 2, TOP, w, 3.20, ACCENT, "③ 돕기", "옆 사람 지원",
          ["최대 2명까지 지원한다",
           "대신 해 주지 않는다",
           "\"어디가 안 돼?\"라고 묻는다"], strong=True,
          note="설명할 수 있어야 아는 것입니다.")
    banner(s, [[("기능을 더 추가하고 싶다면 ", {}),
                ("PRD 6번 LATER에 적어 두세요.", {"bold": True, "color": ACCENT_D}),
                ("  9회기 나만의 프로젝트에서 그대로 씁니다.", {})]],
           y=5.70, h=0.78)
    footer_rule(s)

    # ---------------------------------------------------------- 24. PART 4
    s = S(); part_slide(s, "PART 4", "같은 주제, 다른 결과",
                        "30개의 화면을 놓고 · 무엇이 그 차이를 만들었는지 확인합니다.", 4, 4)

    # ---------------------------------------------------------- 25. 비교
    s = S(); header(s)
    title_block(s, "결과 비교 — 무엇이 이 차이를 만들었나 (10분)",
                "네 명씩 화면을 보여 주고, 아래 두 질문에만 답합니다. 잘했다·못했다는 말하지 않습니다.")
    hw = (CW - 0.3) / 2
    wcard(s, ML, TOP, hw, 2.75, ACCENT, "질문 1", "내 PRD의 어느 문장이 이 차이를 만들었나?",
          ["같은 주제인데 화면이 다르다",
           "그 차이의 근거를 PRD에서 찾아 짚는다",
           "찾지 못했다면 그건 내가 정하지 않은 것"], strong=True)
    wcard(s, ML + hw + 0.3, TOP, hw, 2.75, ACCENT, "질문 2", "저 화면에 있는데 내 PRD엔 없는 결정은?",
          ["다른 사람 화면에서 좋아 보이는 것",
           "그것은 그 사람이 내린 결정이다",
           "내 PRD에 추가할지 지금 정한다"], strong=True)
    banner(s, [[("동료 피드백은 ", {}), ("\"이게 더 낫다\"", {"bold": True, "color": ACCENT_D}),
                ("가 아니라 ", {}), ("\"너는 이걸 정했구나\"", {"bold": True, "color": ACCENT_D}),
                ("로 말합니다.", {})]])
    footer_rule(s)

    # ---------------------------------------------------------- 26. 오늘의 핵심
    s = S(); header(s)
    title_block(s, "오늘의 핵심", "대화창 속 결과물이, 내 기준으로 검증된 내 파일이 되었습니다.")
    w = (CW - 0.26 * 2) / 3
    wcard(s, ML, TOP, w, 2.75, ACCENT, "SPEC", "명세가 결과를 만든다",
          ["같은 주제, 다른 화면",
           "그 차이는 실력이 아니라",
           "내가 내린 결정의 차이였다"], strong=True)
    wcard(s, ML + w + 0.26, TOP, w, 2.75, ACCENT, "VERIFY", "판정은 사람이 한다",
          ["AI가 만든 것을 그대로 믿지 않는다",
           "완성 기준으로 통과·실패를 가른다",
           "시키지 않은 것도 찾아낸다"], strong=True)
    wcard(s, ML + (w + 0.26) * 2, TOP, w, 2.75, ACCENT, "FILE", "내 것이 되었다",
          ["와이파이가 끊겨도 열린다",
           "내 폴더에 내 파일로 남았다",
           "4회기부터 AI가 이 파일을 직접 고친다"], strong=True)
    banner(s, [[("코드를 얼마나 직접 썼는지가 아니라, ", {}),
                ("AI에게 얼마나 정확하게 지시하고 결과를 검증했는지", {"bold": True, "color": ACCENT_D}),
                ("가 오늘의 성장입니다.", {})]])
    footer_rule(s)

    # ---------------------------------------------------------- 27. NEXT
    s = S(); header(s)
    title_block(s, "NEXT · 4회기 — 복사-붙여넣기가 사라지는 시간",
                "AI가 브라우저 밖으로 나와 내 폴더를 직접 엽니다. 1회기에 배운 세 유형 중 Agent형입니다.")
    step_row(s, TOP, [
        ("01", "VS Code", "코드를 편하게 보고 고치는 프로그램을 설치한다"),
        ("02", "Claude Code", "터미널에서 내 폴더를 직접 읽는 AI를 연결한다"),
        ("03", "폴더 열기", "오늘 만든 index.html이 들어 있는 그 폴더를 연다"),
        ("04", "직접 수정", "복사·붙여넣기 없이 AI가 파일을 고친다"),
    ], h=1.9)
    banner(s, [[("오늘 만든 ", {}), ("index.html", {"bold": True, "font": MONO, "color": ACCENT_D}),
                ("과 ", {}), ("PRD.md", {"bold": True, "font": MONO, "color": ACCENT_D}),
                (" 를 같은 폴더에 그대로 두세요. 4회기에 그대로 씁니다.", {})]],
           y=5.05, h=0.9)
    text(s, ML, 6.15, CW, 0.5,
         [[("오늘 한 번 고칠 때마다 복사-붙여넣기를 몇 번 했나요? 그 횟수가 다음 시간에 0이 됩니다.",
            {"size": 11.5, "color": FAINT})]], align=PP_ALIGN.CENTER)
    footer_rule(s)

    # ---------------------------------------------------------- 발표자 노트
    # 스크립트 원본: tools/lesson03_v2_script.py (Markdown 문서와 같은 내용)
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import lesson03_v2_script as script
    slides = list(prs.slides)
    if len(slides) != len(script.NOTES):
        raise SystemExit("슬라이드 %d장 / 노트 %d개 — 개수가 달라 노트를 넣지 않습니다."
                         % (len(slides), len(script.NOTES)))
    for i, slide in enumerate(slides):
        slide.notes_slide.notes_text_frame.text = script.note_text(i)

    prs.save(OUT)
    return len(slides)


if __name__ == "__main__":
    n = build()
    print("생성 완료: %s" % os.path.basename(OUT))
    print("슬라이드 %d장" % n)
    if WARN:
        print("")
        print("[넘침 경고] %d건" % len(WARN))
        for w in WARN:
            print("  " + w)
    else:
        print("넘치는 카드 없음")
