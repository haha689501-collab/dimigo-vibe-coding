# -*- coding: utf-8 -*-
"""
1~3차시 최종 PPT를 4차시 디자인으로 통일한다.  (내용은 그대로, 디자인 요소만 바꾼다)

실행:  python tools/restyle-to-lesson04.py
결과:  01_수업PPT/  폴더에 같은 이름으로 저장 (원본 파일은 열기만 하고 저장하지 않는다)

바꾸는 것
  · 배경색 · 머리말(A I × B U I L D / DIMIGO · VIBE CODING) · 머리말 선 · 꼬리말 선과 쪽 번호
  · 제목 · 부제 · 강조 막대의 위치와 크기 (본문 배치는 그대로)
  · 글꼴 → 맑은 고딕 (코드용 고정폭 글꼴은 그대로)
  · 색상표 → 4차시 색 (테라코타 · 아이보리 · 남색 코드 상자)
  · 카드 그림자 제거 · 흰 카드 테두리 통일 · 그림으로 된 카드 배경을 도형으로 교체
  · 표지 → 4차시 표지 형식
"""
import copy
import io
import os
import re
import shutil
import sys

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUTDIR = os.path.join(ROOT, "01_수업PPT")

DECKS = [
    ("최종_디미고_바이브코딩_1차시.pptx", "1차시 · 바이브코딩과 프롬프트", "01"),
    ("최종_디미고_바이브코딩_2차시_PRD기획서_수정본.pptx", "2차시 · PRD 기획서", "02"),
    ("최종_디미고_바이브코딩_3차시_완성본.pptx", "3차시 · PRD와 MVP", "03"),
]

# ---------------------------------------------------------------- 4차시 디자인 값
FONT = "맑은 고딕"
BG = "FBFAF7"
INK, BODY, MUTED, FAINT = "22211E", "4A443E", "6F6A63", "90887F"
ACCENT, ACCENT_D, ACCENT_L = "C6603C", "A8412B", "E9B99F"
LINE, TINT, WASH, DARK = "E4DCD1", "F2EDE7", "F6E9E1", "1E2233"
SW, SH = 13.333, 7.5
ML, CW = 0.73, 11.87
FOOT_Y = 6.95

MONO = ("consolas", "d2coding", "courier", "mono", "menlo")

# 도형(채우기·선)에 쓰인 색 → 4차시 색
SHAPE_MAP = {
    "B36F4B": ACCENT, "C9653F": ACCENT, "D98762": ACCENT, "9E5F3E": ACCENT_D, "C0553C": ACCENT_D,
    "F7F4EE": BG, "F1ECE4": TINT, "EFE5DC": WASH, "F0E6DD": WASH, "F2EAE3": WASH, "EDE0D6": WASH,
    "F6EDE6": WASH, "F2DFD3": WASH, "E8D3C7": ACCENT_L, "EEF0FE": WASH, "EAE5DE": TINT,
    "302E2A": DARK, "2B2926": DARK, "242321": DARK, "141B2D": DARK, "191817": DARK,
    "2F7B60": "2F7A55", "C9C1B7": "C9C1B7",
}
# 글자 색 → 4차시 색
TEXT_MAP = {
    "191817": INK, "242220": INK, "111827": INK, "77716B": MUTED, "667085": MUTED, "7B7269": MUTED,
    "B36F4B": ACCENT, "C9653F": ACCENT, "D98762": ACCENT, "C0553C": ACCENT_D, "4F46E5": ACCENT_D,
    "38BDF8": "8FC1F5", "4ADE80": "8FD9B6", "C084FC": "C9B6F2", "F43F5E": "F29B9B",
    "F1F5F9": "FFFFFF", "94A3B8": "9AA1BD", "2F7B60": "2F7A55",
}

HEADER_TEXTS = {"AI×BUILD", "DIMIGO·VIBECODING", "VIBECODING", "한국디지털미디어고등학교·VIBECODINGPROJECT"}


# ---------------------------------------------------------------- 도우미
def inch(v):
    return v / 914400.0 if v is not None else 0.0


def norm(t):
    return re.sub(r"\s+", "", t or "").upper()


def lum(hexv):
    r, g, b = int(hexv[0:2], 16), int(hexv[2:4], 16), int(hexv[4:6], 16)
    return 0.299 * r + 0.587 * g + 0.114 * b


def first_size(sh):
    if not sh.has_text_frame:
        return 0
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size:
                return r.font.size.pt
    return 0


def remove(sh):
    el = sh._element
    el.getparent().remove(el)


def set_fill(sh, hexv):
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor.from_string(hexv)


def no_line(sh):
    sh.line.fill.background()


def add_rect(slide, x, y, w, h, fill, line=None, radius=None):
    shape = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius:
        s.adjustments[0] = radius
    set_fill(s, fill)
    if line:
        s.line.color.rgb = RGBColor.from_string(line)
        s.line.width = Pt(1)
    else:
        no_line(s)
    s.shadow.inherit = False
    if s.has_text_frame:
        s.text_frame.text = ""
    return s


def add_text(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT, spacing=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = RGBColor.from_string(color)
    set_run_font(r._r)
    if spacing:
        r._r.get_or_add_rPr().set("spc", str(int(spacing * 100)))
    return tb


def set_run_font(r_el, shrink=False):
    """run(a:r)의 글꼴을 맑은 고딕으로. 코드용 고정폭 글꼴은 건드리지 않는다.
    shrink=True 이면, 원래 맑은 고딕이 아니던 글자를 5% 줄여 줄바꿈이 늘지 않게 한다 (맑은 고딕이 조금 넓다)."""
    rPr = r_el.find(qn("a:rPr"))
    if rPr is None:
        rPr = etree.SubElement(r_el, qn("a:rPr"))
        r_el.remove(rPr)
        r_el.insert(0, rPr)
    lat = rPr.find(qn("a:latin"))
    face = (lat.get("typeface") if lat is not None else "") or ""
    if any(m in face.lower() for m in MONO):
        return
    if shrink and face and face != FONT and rPr.get("sz"):
        rPr.set("sz", str(int(round(int(rPr.get("sz")) * 0.95 / 50.0) * 50)))
    for tag in ("a:latin", "a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is not None:
            rPr.remove(e)
    anchor = None
    for tag in ("a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"):
        anchor = rPr.find(qn(tag))
        if anchor is not None:
            break
    for tag in ("a:latin", "a:ea", "a:cs"):
        e = etree.Element(qn(tag))
        e.set("typeface", FONT)
        if anchor is not None:
            anchor.addprevious(e)
        else:
            rPr.append(e)


def style_runs(sh, size=None, color=None, bold=None):
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if size:
                r.font.size = Pt(size)
            if color:
                r.font.color.rgb = RGBColor.from_string(color)
            if bold is not None:
                r.font.bold = bold
            set_run_font(r._r)


def place(sh, x, y, w, h):
    sh.left, sh.top, sh.width, sh.height = Inches(x), Inches(y), Inches(w), Inches(h)


def remap_colors(el):
    """도형 · 글자에 쓰인 srgbClr 값을 4차시 색으로 바꾼다."""
    for c in el.iter(qn("a:srgbClr")):
        v = (c.get("val") or "").upper()
        anc = c.getparent()
        is_text = False
        while anc is not None and anc is not el:
            t = etree.QName(anc).localname
            if t in ("rPr", "defRPr", "endParaRPr"):
                is_text = True
                break
            if t in ("spPr", "bg", "bgPr"):
                break
            anc = anc.getparent()
        m = TEXT_MAP if is_text else SHAPE_MAP
        if v in m:
            c.set("val", m[v])


def slide_is_dark(slide, png):
    if png and os.path.exists(png):
        im = Image.open(png).convert("RGB")
        w, h = im.size
        pts = [im.getpixel((5, h // 2)), im.getpixel((w - 5, h // 2)), im.getpixel((w // 2, h - 5))]
        return sum(0.299 * r + 0.587 * g + 0.114 * b for r, g, b in pts) / 3 < 90
    return False


def replace_picture_backgrounds(slide):
    """카드 배경으로 쓰인 큰 그림을 같은 자리의 4차시 카드 도형으로 바꾼다."""
    n = 0
    for sh in list(slide.shapes):
        if sh.shape_type != 13:
            continue
        w, h = inch(sh.width), inch(sh.height)
        if w * h < 1.5:
            continue
        x, y = inch(sh.left), inch(sh.top)
        texts_on_top = [t for t in slide.shapes if t.has_text_frame and t.text_frame.text.strip()
                        and x - 0.05 <= inch(t.left) <= x + w and y - 0.05 <= inch(t.top) <= y + h]
        if not texts_on_top:
            continue
        try:
            im = Image.open(io.BytesIO(sh.image.blob)).convert("RGB").resize((24, 24))
            px = [im.getpixel((i % 24, i // 24)) for i in range(24 * 24)]
            mean = sum(0.299 * r + 0.587 * g + 0.114 * b for r, g, b in px) / len(px)
        except Exception:
            mean = 255
        dark = mean < 110
        card = add_rect(slide, x, y, w, h, DARK if dark else "FFFFFF", None if dark else LINE, radius=0.05)
        sh._element.addprevious(card._element)       # 그림이 있던 겹침 순서에 그대로
        remove(sh)
        n += 1
    return n


# ---------------------------------------------------------------- 머리말 · 꼬리말
def strip_chrome(slide):
    for sh in list(slide.shapes):
        x, y, w, h = inch(sh.left), inch(sh.top), inch(sh.width), inch(sh.height)
        txt = sh.text_frame.text.strip() if sh.has_text_frame else ""
        nt = norm(txt)
        if nt in HEADER_TEXTS:
            remove(sh); continue
        if re.fullmatch(r"PART\d+/\d+", nt) and y > 6.4:
            remove(sh); continue
        if not txt and h <= 0.035 and w >= 9.5 and (y <= 1.2 or y >= 6.3):
            remove(sh); continue                       # 머리말 · 꼬리말 선
        if not txt and y >= 6.85 and h <= 0.3 and sh.shape_type != 13:
            remove(sh); continue


def content_bottom(slide):
    return max([inch(s.top) + inch(s.height) for s in slide.shapes
                if not (inch(s.width) >= 13 and inch(s.height) >= 7)] or [0])


def add_chrome(slide, label, dark=False):
    low = content_bottom(slide) > 6.9          # 맨 아래까지 내용이 있으면 꼬리말을 생략
    hc = ACCENT_L if dark else ACCENT
    rc = "3A4055" if dark else LINE
    fc = "9AA1BD" if dark else FAINT
    add_text(slide, ML, 0.30, 4.0, 0.30, "A I   ×   B U I L D", 10.5, hc, bold=True, spacing=1.3)
    add_text(slide, 7.6, 0.31, 5.0, 0.30, "DIMIGO · VIBE CODING", 9.5, fc, bold=True, align=PP_ALIGN.RIGHT, spacing=1.0)
    add_rect(slide, ML, 0.62, CW, 0.018, rc)
    if not low:
        add_rect(slide, ML, FOOT_Y, CW, 0.018, rc)
        add_text(slide, 7.6, FOOT_Y + 0.12, 5.0, 0.26, label, 9, fc, align=PP_ALIGN.RIGHT, spacing=0.8)


# ---------------------------------------------------------------- 제목 · 부제 · 막대
def fix_title_block(slide, dark=False):
    shapes = list(slide.shapes)
    cands = [s for s in shapes if s.has_text_frame and s.text_frame.text.strip()
             and 0.5 <= inch(s.top) <= 1.35 and inch(s.width) >= 5 and first_size(s) >= 20]
    if not cands:
        return False
    title = min(cands, key=lambda s: inch(s.top))
    if len(title.text_frame.paragraphs) > 1 or first_size(title) > 34:
        return False                                    # 큰 강조 문장(요약 슬라이드)은 그대로
    t_top = inch(title.top)
    # 본문이 일찍 시작하면 부제 · 막대를 그만큼 위로 (본문을 가리지 않게)
    body = [s for s in shapes if s is not title and inch(s.top) >= 1.75 and inch(s.width) > 1.5
            and not (s.has_text_frame and s.text_frame.text.strip() and inch(s.top) < 2.05 and inch(s.height) <= 0.7)]
    body_top = min([inch(s.top) for s in body] or [2.34])
    bar_y = min(1.96, max(1.74, body_top - 0.20))
    sub_y = bar_y - 0.44
    place(title, ML, 0.84, CW, 0.62)
    style_runs(title, size=27, color="FFFFFF" if dark else INK, bold=True)
    tf = title.text_frame
    tf.margin_left = tf.margin_right = 0
    tf.word_wrap = True
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.LEFT

    subs = [s for s in shapes if s is not title and s.has_text_frame and s.text_frame.text.strip()
            and t_top + 0.3 <= inch(s.top) <= 2.05 and inch(s.width) >= 5 and inch(s.height) <= 0.7
            and 9 <= (first_size(s) or 12) <= 17]
    if subs:
        sub = min(subs, key=lambda s: inch(s.top))
        place(sub, ML, sub_y, CW, 0.34)
        style_runs(sub, size=12, color="9AA1BD" if dark else MUTED, bold=False)
        sub.text_frame.margin_left = sub.text_frame.margin_right = 0
        for p in sub.text_frame.paragraphs:
            p.alignment = PP_ALIGN.LEFT

    for s in shapes:                                     # 강조 막대
        if s is title or (s.has_text_frame and s.text_frame.text.strip()):
            continue
        x, y, w, h = inch(s.left), inch(s.top), inch(s.width), inch(s.height)
        if 0.4 <= w <= 1.7 and h <= 0.07 and 1.55 <= y <= 2.15 and x < 1.2:
            remove(s)
    add_rect(slide, ML, bar_y, 0.80, 0.045, ACCENT)
    return True


# ---------------------------------------------------------------- 표지
def fix_cover(slide, deck_label, num):
    shapes = list(slide.shapes)
    number = title = footer = None
    for s in shapes:
        if not s.has_text_frame:
            continue
        t = s.text_frame.text.strip()
        if re.fullmatch(r"\d{2}", t) and first_size(s) >= 48:
            number = s
        elif "노코드 창작 클래스" in t or norm(t) in HEADER_TEXTS:
            remove(s)
        elif "VIBE CODING PROJECT" in t.upper():
            footer = t
            remove(s)
        elif first_size(s) >= 30 and t:
            title = s
    for s in list(slide.shapes):                         # 예전 선 · 막대
        if s.has_text_frame and s.text_frame.text.strip():
            continue
        x, y, w, h = inch(s.left), inch(s.top), inch(s.width), inch(s.height)
        if (h <= 0.06 and w >= 1.2) or (w >= 13 and h >= 7):
            remove(s)

    # 진행 표시(IDEA · SPEC · BUILD · TEST)는 아래쪽으로 옮긴다
    dots = [s for s in slide.shapes if s is not number and s is not title and 1.2 <= inch(s.top) <= 2.7]
    if dots:
        top = min(inch(s.top) for s in dots)
        for s in dots:
            s.top = Inches(inch(s.top) - top + 5.30)

    add_text(slide, ML, 0.44, 5, 0.3, "A I   ×   B U I L D", 11, ACCENT, bold=True, spacing=1.3)
    add_text(slide, SW - ML - 7.4, 0.46, 7.4, 0.3, "AI 노코드 창작 클래스 · %s차시" % int(num), 10, FAINT,
             bold=True, align=PP_ALIGN.RIGHT, spacing=1.0)
    if number is not None:
        place(number, ML, 1.55, 3.0, 1.2)
        style_runs(number, size=60, color=ACCENT_L, bold=True)
        number.text_frame.margin_left = 0
    y = 2.72
    if title is not None:
        paras = title.text_frame.paragraphs
        def sig(p):
            r = p.runs[0] if p.runs else None
            if r is None:
                return None
            try:
                col = str(r.font.color.rgb) if r.font.color and r.font.color.type is not None else None
            except Exception:
                col = None
            return (r.font.size.pt if r.font.size else None, col)
        real = [p for p in paras if p.text.strip()]
        if len(real) >= 3:
            n_title = 1
        elif len(real) == 2 and sig(real[0]) == sig(real[1]):
            n_title = 2
        else:
            n_title = 1
        sub_lines = [p.text for p in paras[n_title:] if p.text.strip()]
        for p in list(paras[n_title:]):
            p._p.getparent().remove(p._p)
        h = 0.68 * n_title + 0.05
        place(title, ML, y, CW, h)
        style_runs(title, size=40, color=INK, bold=True)
        tf = title.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.word_wrap = True
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.LEFT
        y += h + 0.18
        add_rect(slide, ML, y, 0.9, 0.05, ACCENT)
        y += 0.26
        if sub_lines:
            tb = add_text(slide, ML, y, CW * 0.86, 0.36 * len(sub_lines), sub_lines[0], 17, BODY)
            for extra in sub_lines[1:]:
                p = tb.text_frame.add_paragraph()
                r = p.add_run(); r.text = extra
                r.font.size = Pt(17); r.font.color.rgb = RGBColor.from_string(BODY)
                set_run_font(r._r)
    add_rect(slide, ML, FOOT_Y, CW, 0.018, LINE)
    add_text(slide, ML, FOOT_Y + 0.12, 8, 0.26, footer or "한국디지털미디어고등학교 · Vibe Coding Project", 9.5, FAINT, spacing=0.8)


def is_cover(slide):
    return any(s.has_text_frame and "노코드 창작 클래스" in s.text_frame.text for s in slide.shapes)


# ---------------------------------------------------------------- 공통 다듬기
def tidy_shapes(slide):
    for sh in slide.shapes:
        if sh.shape_type in (1,):                        # 도형: 그림자 제거, 흰 카드 테두리
            try:
                sh.shadow.inherit = False
            except Exception:
                pass
            try:
                if sh.fill.type == 1:
                    v = str(sh.fill.fore_color.rgb)
                    full = inch(sh.width) >= 13 and inch(sh.height) >= 7
                    if v in ("FBFAF7", "FBF9F6", "FFFFFF") and not full:
                        set_fill(sh, "FFFFFF")
                        if sh.line.fill.type in (None, 5) or sh.line.width == 0:
                            sh.line.color.rgb = RGBColor.from_string(LINE)
                            sh.line.width = Pt(1)
            except Exception:
                pass
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    set_run_font(r._r, shrink=True)


def set_background(slide, dark):
    if dark:
        return
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = RGBColor.from_string(BG)
    for sh in slide.shapes:
        if sh.shape_type == 1 and inch(sh.width) >= 13 and inch(sh.height) >= 7 and inch(sh.left) <= 0.1:
            try:
                if sh.fill.type == 1 and lum(str(sh.fill.fore_color.rgb)) > 200:
                    set_fill(sh, BG)
            except Exception:
                pass


# ---------------------------------------------------------------- 실행
def strip_layout_chrome(prs):
    n = 0
    for m in prs.slide_masters:
        for lay in list(m.slide_layouts) + [m]:
            for sh in list(lay.shapes):
                if sh.is_placeholder:
                    continue
                txt = sh.text_frame.text.strip() if sh.has_text_frame else ""
                y, w, h = inch(sh.top), inch(sh.width), inch(sh.height)
                if norm(txt) in HEADER_TEXTS or (not txt and h <= 0.035 and w >= 9.5 and (y <= 1.2 or y >= 6.3)):
                    remove(sh); n += 1
    return n


def restyle(src, label, num, png_prefix):
    prs = Presentation(src)
    strip_layout_chrome(prs)
    total = len(prs.slides)
    report = {"title": 0, "cover": 0, "pic": 0, "dark": []}
    for i, slide in enumerate(prs.slides, 1):
        png = "%s-%d.png" % (png_prefix, i) if png_prefix else None
        dark = slide_is_dark(slide, png)
        if dark:
            report["dark"].append(i)
        remap_colors(slide._element)
        report["pic"] += replace_picture_backgrounds(slide)
        set_background(slide, dark)
        if is_cover(slide):
            fix_cover(slide, label, num)
            tidy_shapes(slide)
            report["cover"] += 1
            continue
        strip_chrome(slide)
        if fix_title_block(slide, dark):
            report["title"] += 1
        tidy_shapes(slide)
        add_chrome(slide, "%s   %02d / %02d" % (label, i, total), dark)
    return prs, report


if __name__ == "__main__":
    png_dir = sys.argv[1] if len(sys.argv) > 1 else None   # 원본 슬라이드 그림 폴더 (어두운 슬라이드 판별용)
    os.makedirs(OUTDIR, exist_ok=True)
    for k, (fname, label, num) in enumerate(DECKS, 1):
        src = os.path.join(ROOT, "01_수업PPT", "지난버전", fname)
        prefix = os.path.join(png_dir, "d%d" % k) if png_dir else None
        prs, rep = restyle(src, label, num, prefix)
        out = os.path.join(OUTDIR, fname)
        prs.save(out)
        print("%s → 01_수업PPT/  (슬라이드 %d장 · 제목 정리 %d · 표지 %d · 그림 카드 교체 %d · 어두운 슬라이드 %s)"
              % (fname, len(prs.slides), rep["title"], rep["cover"], rep["pic"], rep["dark"] or "없음"))
