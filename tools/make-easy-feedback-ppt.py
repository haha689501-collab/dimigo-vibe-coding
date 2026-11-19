# -*- coding: utf-8 -*-
"""팀 프로젝트 피드백 — '선생님이 실제로 말할 수 있는 수준' (기술 용어 없는 쉬운 버전)

사용:  python -I tools/make-easy-feedback-ppt.py docs/students/_피드백자료 docs/students <출력.pptx>
- 내용: _피드백자료/easy_teams.json (작품 소개 · 칭찬 3 · 다음 단계 · 질문 · 수업 전 해 볼 것)
- 대본: _피드백자료/easy_notes.json  ({팀id: [표지, 칭찬, 다음 단계, 질문] 대본})
- 팀마다 4장, 각 장 바로 뒤에 '대본' 장 (노트 칸에도 같은 대본)
- 기존 PPT 는 열지도 고치지도 않습니다. 같은 이름 파일이 있으면 멈춥니다.
"""
import io, json, os, re, sys
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

DATA, STUD, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
if os.path.exists(OUT):
    sys.exit("이미 같은 이름의 파일이 있습니다 — 덮어쓰지 않습니다: " + OUT)

FONT = "맑은 고딕"
C = lambda h: RGBColor.from_string(h)
BG, INK, BODY, MUTED, FAINT, LINE = "FBFAF7", "22211E", "4A443E", "6F6A63", "90887F", "E4DCD1"
ACC, ACC_D, ACC_W = "C6603C", "A8412B", "F6E9E1"
GREEN, GREEN_W, GREEN_T = "2F7A55", "F2F8F4", "DDEFE3"
BLUE, BLUE_W, BLUE_T = "3B5B9B", "F2F5FA", "DFE7F5"
NAVY = "1E2233"
SW, SH = 13.333, 7.5

teams = json.load(io.open(os.path.join(DATA, "easy_teams.json"), encoding="utf-8"))
notes_all = json.load(io.open(os.path.join(DATA, "easy_notes.json"), encoding="utf-8"))

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
BLANK = prs.slide_layouts[6]
tmp = []


def rect(s, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = C(fill)
    if line: r.line.color.rgb = C(line); r.line.width = Pt(1)
    else: r.line.fill.background()
    r.shadow.inherit = False
    if radius is not None: r.adjustments[0] = radius
    return r


def text(s, x, y, w, h, runs, size=18, color=BODY, bold=False, align=PP_ALIGN.LEFT, space=1.25, after=6, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    for i, ln in enumerate(runs if isinstance(runs, list) else [runs]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = space; p.space_after = Pt(after)
        for part in (ln if isinstance(ln, list) else [ln]):
            t, o = (part, {}) if isinstance(part, str) else part
            r = p.add_run(); r.text = t
            f = r.font; f.name = FONT; f.size = Pt(o.get("size", size)); f.bold = o.get("bold", bold); f.color.rgb = C(o.get("color", color))
    return tb


def slide(bg=BG, band=None):
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid(); s.background.fill.fore_color.rgb = C(bg)
    if band: rect(s, 0, 0, SW, 0.12, band)
    return s


def pill(s, x, y, w, label, fill, col):
    r = rect(s, x, y, w, 0.46, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    p = r.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    run = p.add_run(); run.text = label; run.font.name = FONT; run.font.size = Pt(15); run.font.bold = True; run.font.color.rgb = C(col)


# ---------- 대본 장 ----------
def fits(items, w, size, h=6.2):
    em = size / 72.0; used = 0.0
    for hd, b in items:
        for ln in [hd] + [x for x in b.split("\n") if x.strip()] + [""]:
            wid = sum(1.0 if ord(c) > 0x2E80 else (0.3 if c == " " else 0.56) for c in ln) * em
            used += max(1, -(-wid // w)) * em * 1.2 * 1.22 + 3 / 72.0
    return used <= h


def script(body, head):
    parts = re.split(r"^\[([^\]\n]+)\]\s*$", body.strip(), flags=re.M)
    secs = [(parts[i].strip(), parts[i + 1].strip()) for i in range(1, len(parts) - 1, 2)] or [("대본", body.strip())]
    main = [x for x in secs if x[0] == "말할 내용"] or secs[:1]
    rest = [x for x in secs if x not in main]
    s = slide("FFFFFF")
    rect(s, 0, 0, SW, 0.62, NAVY)
    text(s, 0.6, 0.17, 1.6, 0.3, "대본", size=14, color="FFFFFF", bold=True)
    text(s, 1.5, 0.19, 11.2, 0.3, head, size=11, color="C9CCD8", bold=True)

    def col(x, w, items):
        size = next((z for z in (17, 16, 15, 14, 13, 12, 11) if fits(items, w, z)), 11)
        lines = []
        for h, b in items:
            lines.append([(h, {"bold": True, "color": ACC_D, "size": size})])
            lines += [ln.strip() for ln in b.split("\n") if ln.strip()]
            lines.append("")
        text(s, x, 0.95, w, 6.3, lines[:-1], size=size, color=INK, space=1.2, after=3)

    if rest:
        col(0.6, 7.0, main); rect(s, 7.85, 1.0, 0.02, 6.0, LINE); col(8.1, 4.65, rest)
    else:
        col(0.6, 12.1, main)
    s.notes_slide.notes_text_frame.text = body


def with_script(s, body, head):
    s.notes_slide.notes_text_frame.text = body
    if body.strip():
        script(body, head)


def png(path):
    if not path or not os.path.isfile(path): return None
    d = os.path.join(os.path.dirname(OUT), "_tmp_" + os.path.basename(path) + ".png")
    Image.open(path).convert("RGB").save(d, "PNG"); tmp.append(d)
    return d


# ---------- 맨 앞 ----------
s = slide()
text(s, 0.9, 2.1, 11, 0.4, "TEAM PROJECT FEEDBACK", size=14, color=ACC_D, bold=True)
text(s, 0.9, 2.6, 11.5, 1.2, "팀 프로젝트 피드백", size=46, color=INK, bold=True)
text(s, 0.9, 3.9, 11.5, 1.0, ["작품 소개 · 칭찬 3가지 · 다음 단계 · 함께 이야기할 질문",
                             " · ".join(t["name"] for t in teams)], size=18, color=BODY)
with_script(s, """[말할 내용]
오늘은 팀 프로젝트를 한 팀씩 같이 볼 거예요.
팀마다 먼저 어떤 작품인지 소개하고, 선생님이 특히 좋았던 점 세 가지를 이야기할게요.
그다음 이렇게 해 보면 더 좋아질 것 같은 다음 단계를 말하고, 마지막에는 여러분에게 질문을 할 거예요.
질문에는 정답이 없어요. 만든 사람이 제일 잘 아니까, 편하게 설명해 주면 돼요.
다른 팀 이야기도 잘 들어 주세요. 다른 팀의 좋은 점이 우리 팀에도 힌트가 될 때가 많아요.

[화면 동작]
→ 방향키로 넘깁니다. 이 PPT의 '대본' 장은 선생님만 보는 장이니, 학생 앞에서는 발표자 보기로 띄우거나 대본 장을 건너뛰세요.""", "시작")

# ---------- 팀별 ----------
for t in teams:
    nt = notes_all.get(t["id"], ["", "", "", ""])
    who = t["name"] + "  ·  " + " · ".join(t["members"])

    # 1 표지
    s = slide()
    text(s, 0.75, 0.42, 9, 0.3, who, size=11, color=MUTED, bold=True)
    text(s, 0.75, 1.6, 6.6, 0.35, "TEAM PROJECT", size=13, color=ACC_D, bold=True)
    text(s, 0.75, 2.0, 6.6, 1.2, t["name"], size=44, color=INK, bold=True)
    text(s, 0.75, 3.3, 6.4, 1.6, t["what"], size=19, color=ACC_D, bold=True, after=2)
    text(s, 0.75, 5.0, 6.4, 0.5, " · ".join(t["members"]), size=16, color=BODY, bold=True)
    img = png(os.path.join(STUD, *t.get("preview", "").split("/"))) if t.get("preview") else None
    if img:
        rect(s, 7.55, 1.75, 5.05, 0.32, "F1ECE4", line=LINE)
        for i, col in enumerate(["E5675B", "EDB84B", "5DBB63"]):
            rect(s, 7.72 + i * 0.22, 1.85, 0.12, 0.12, col, shape=MSO_SHAPE.OVAL)
        s.shapes.add_picture(img, Inches(7.55), Inches(2.07), Inches(5.05), Inches(5.05 * 600 / 960))
    with_script(s, nt[0], "표지 · " + who)

    # 2 칭찬 3가지
    s = slide(GREEN_W, GREEN)
    text(s, 0.75, 0.42, 9, 0.3, who, size=11, color=MUTED, bold=True)
    pill(s, 0.75, 0.95, 2.4, "✓  칭찬 3가지", GREEN_T, "22603F")
    for i, (h, d) in enumerate(t["praise"][:3]):
        x = 0.75 + i * 4.0
        rect(s, x, 1.85, 3.75, 4.9, "FFFFFF", line="CFE5D7", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
        text(s, x + 0.35, 2.15, 3.1, 0.6, "0%d" % (i + 1), size=30, color=GREEN, bold=True)
        text(s, x + 0.35, 2.95, 3.1, 1.2, h, size=22, color=INK, bold=True, space=1.15)
        text(s, x + 0.35, 4.25, 3.1, 2.4, d, size=15, color=BODY, space=1.35)
    with_script(s, nt[1], "칭찬 3가지 · " + t["name"])

    # 3 다음 단계
    s = slide(BLUE_W, BLUE)
    text(s, 0.75, 0.42, 9, 0.3, who, size=11, color=MUTED, bold=True)
    pill(s, 0.75, 0.95, 2.4, "↗  다음 단계", BLUE_T, "2F4C86")
    nx = t["next"][:2]
    for i, (h, d) in enumerate(nx):
        y = 1.95 + i * 2.45 if len(nx) > 1 else 2.6
        rect(s, 0.75, y, 11.85, 2.15, "FFFFFF", line="C9D6EE", shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
        text(s, 1.15, y + 0.3, 0.8, 0.6, str(i + 1), size=28, color=BLUE, bold=True)
        text(s, 1.95, y + 0.3, 10.3, 0.6, h, size=26, color=INK, bold=True)
        text(s, 1.95, y + 1.0, 10.3, 1.0, d, size=17, color=BODY, space=1.3)
    with_script(s, nt[2], "다음 단계 · " + t["name"])

    # 4 함께 이야기할 질문
    s = slide()
    text(s, 0.75, 0.42, 9, 0.3, who, size=11, color=MUTED, bold=True)
    pill(s, 0.75, 0.95, 3.0, "?  함께 이야기할 질문", ACC_W, ACC_D)
    for i, q in enumerate(t["questions"][:3]):
        y = 2.0 + i * 1.55
        rect(s, 0.75, y, 0.75, 0.75, ACC_W, shape=MSO_SHAPE.OVAL)
        text(s, 0.75, y + 0.12, 0.75, 0.5, str(i + 1), size=22, color=ACC_D, bold=True, align=PP_ALIGN.CENTER)
        text(s, 1.8, y + 0.1, 10.8, 1.2, q, size=24, color=INK, bold=True, space=1.2)
    with_script(s, nt[3], "질문 · " + t["name"])

# ---------- 마무리 ----------
s = slide()
text(s, 0.9, 2.3, 11, 0.4, "마무리", size=14, color=ACC_D, bold=True)
text(s, 0.9, 2.8, 11.5, 1.0, "다음 시간까지", size=42, color=INK, bold=True)
text(s, 0.9, 4.0, 11.5, 2.0, ["팀별 '다음 단계' 하나를 정해서 해 보기", "해 본 것과 바뀐 점을 기록에 남기기"], size=21, color=BODY, after=10)
with_script(s, """[말할 내용]
오늘 다섯 팀을 모두 봤어요. 팀마다 아이디어도, 만든 방식도 달라서 선생님도 많이 배웠어요.
다음 시간까지 각 팀은 '다음 단계' 중에 하나를 골라서 해 보세요.
무엇을 했고, 해 보니 무엇이 달라졌는지 꼭 기록해 주세요. 다음 시간에 그 이야기를 듣고 싶어요.
오늘 받은 질문에 대한 답도 팀끼리 한 번 더 이야기해 보면 좋겠어요.
모두 수고 많았어요. 다 같이 박수 한번 칠까요?

[화면 동작]
전체화면은 Esc 로 나옵니다.""", "마무리")

prs.save(OUT)
for f in tmp:
    try: os.remove(f)
    except OSError: pass
print("저장:", OUT, "· 슬라이드", len(prs.slides._sldIdLst))
