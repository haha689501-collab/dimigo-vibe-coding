# -*- coding: utf-8 -*-
"""팀 프로젝트 피드백 — 교사 발표 연습용 PPT (발표자 노트에 실제 말할 대본)

사용:  python -I make_feedback_ppt.py <검증JSON·노트JSON 폴더> <미리보기 폴더> <출력 pptx>
- 슬라이드는 피드백 페이지의 '학생 전달용' 순서와 같습니다 (팀마다 9장).
- 기존 PPT 는 열지도, 고치지도 않습니다. 새 파일만 만듭니다 (같은 이름이 있으면 멈춤).
"""
import io, json, os, re, sys
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC, PREV, OUT = sys.argv[1], sys.argv[2], sys.argv[3]
if os.path.exists(OUT):
    sys.exit("이미 같은 이름의 파일이 있습니다 — 덮어쓰지 않습니다: " + OUT)

FONT = "맑은 고딕"
C = lambda h: RGBColor.from_string(h)
BG, INK, BODY, MUTED, FAINT, LINE = "FBFAF7", "22211E", "4A443E", "6F6A63", "90887F", "E4DCD1"
ACC, ACC_D, ACC_W = "C6603C", "A8412B", "F6E9E1"
GREEN, GREEN_W, GREEN_T = "2F7A55", "F2F8F4", "DDEFE3"
BLUE, BLUE_W, BLUE_T = "3B5B9B", "F2F5FA", "DFE7F5"
SW, SH = 13.333, 7.5
ORDER = ["byeoldam", "useless", "aro", "draft", "team5"]
MEMBERS = {"byeoldam": "국도연 · 서시우 · 하세린 · 한가윤", "useless": "김대현 · 오태현 · 이준수 · 한승화",
           "aro": "유하람 · 이준서 · 차예린 · 홍의준", "draft": "안지호 · 이지후 · 이채호", "team5": "고준형 · 김리온 · 이호연"}
PREVIEW = {"byeoldam": "byeoldam.webp", "useless": "useless-az.webp", "aro": "aro.webp", "draft": "draft-room.webp"}

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
BLANK = prs.slide_layouts[6]


def clean(t):
    """슬라이드에 보이는 글에서 '(644행)' 같은 코드 위치 표시는 뺍니다 (노트·피드백 페이지에는 그대로)"""
    t = re.sub(r"\s*\((?:[^()]*?\d[^()]*?)행\)", "", t)
    t = re.sub(r"\s*\d+(?:\s*[·,–-]\s*\d+)*행", "", t)
    return re.sub(r"\s{2,}", " ", t).strip()


def rect(s, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = C(fill)
    if line: r.line.color.rgb = C(line); r.line.width = Pt(1)
    else: r.line.fill.background()
    r.shadow.inherit = False
    return r


def text(s, x, y, w, h, runs, size=18, color=BODY, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space=1.25, after=6):
    """runs: 문자열 또는 [문자열 | (문자열, {옵션})] 의 줄 목록"""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0); tf.margin_top = tf.margin_bottom = Inches(0)
    lines = runs if isinstance(runs, list) else [runs]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = space; p.space_after = Pt(after)
        parts = ln if isinstance(ln, list) else [ln]
        for part in parts:
            t, o = (part, {}) if isinstance(part, str) else part
            r = p.add_run(); r.text = t
            f = r.font; f.name = FONT; f.size = Pt(o.get("size", size)); f.bold = o.get("bold", bold)
            f.color.rgb = C(o.get("color", color))
    return tb


LAST = {"team": None, "label": None}     # 바로 앞 슬라이드 정보 (대본 장 머리에 씁니다)


def base(kind="plain", team=None, label=None):
    LAST["team"], LAST["label"] = team, label
    s = prs.slides.add_slide(BLANK)
    bgc = {"good": GREEN_W, "imp": BLUE_W}.get(kind, BG)
    s.background.fill.solid(); s.background.fill.fore_color.rgb = C(bgc)
    if kind in ("good", "imp"):
        rect(s, 0, 0, SW, 0.12, GREEN if kind == "good" else BLUE)
    if team:
        text(s, 0.75, 0.42, 8, 0.3, team, size=11, color=MUTED, bold=True)
    if label:
        text(s, SW - 4.75, 0.42, 4, 0.3, label, size=11, bold=True, align=PP_ALIGN.RIGHT,
             color={"good": GREEN, "imp": BLUE}.get(kind, ACC_D))
    return s


def tag(s, kind):
    if kind == "good":
        r = rect(s, 0.75, 0.95, 1.75, 0.45, GREEN_T, shape=MSO_SHAPE.ROUNDED_RECTANGLE); txt, col = "✓  잘한 점", "22603F"
    else:
        r = rect(s, 0.75, 0.95, 2.3, 0.45, BLUE_T, shape=MSO_SHAPE.ROUNDED_RECTANGLE); txt, col = "↗  더 좋아지려면", "2F4C86"
    r.adjustments[0] = 0.5
    tf = r.text_frame; tf.margin_left = tf.margin_right = Inches(0.1)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    run = p.add_run(); run.text = txt; run.font.name = FONT; run.font.size = Pt(15); run.font.bold = True; run.font.color.rgb = C(col)


SCRIPT_PAGES = "--no-script-pages" not in sys.argv   # 기본: 슬라이드마다 바로 뒤에 '대본' 장을 붙임
NAVY, SCRIPT_BG = "1E2233", "FFFFFF"


def script_page(body, team, label):
    """발표자 노트 내용을 본문에서 읽을 수 있게 한 장으로 — 왼쪽 [말할 내용], 오른쪽 나머지"""
    parts = re.split(r"^\[([^\]\n]+)\]\s*$", body.strip(), flags=re.M)
    secs = [(parts[i].strip(), parts[i + 1].strip()) for i in range(1, len(parts) - 1, 2)]
    if not secs:
        secs = [("대본", body.strip())]
    left, right = secs[:1], secs[1:]
    if len(secs) > 2 and len(secs[0][1]) < 260:        # 첫 칸이 짧으면 두 칸씩 나눔
        left, right = secs[:len(secs) // 2], secs[len(secs) // 2:]

    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid(); s.background.fill.fore_color.rgb = C(SCRIPT_BG)
    rect(s, 0, 0, SW, 0.62, NAVY)
    text(s, 0.6, 0.17, 1.6, 0.3, "대본", size=14, color="FFFFFF", bold=True)
    head = " · ".join(x for x in [label, team] if x)
    text(s, 1.5, 0.19, 11.2, 0.3, head, size=11, color="C9CCD8", bold=True)

    def fits(items, w, size, h=6.2):
        """글자 폭으로 줄 수를 어림해 칸 높이에 들어가는지 봅니다 (한글 1em, 영문 0.56em)"""
        em = size / 72.0
        used = 0.0
        for hd, b in items:
            for ln in [hd] + [x for x in b.split("\n") if x.strip()] + [""]:
                wid = sum(1.0 if ord(c) > 0x2E80 else (0.3 if c == " " else 0.56) for c in ln) * em
                used += max(1, -(-wid // w)) * em * 1.2 * 1.22 + 3 / 72.0
        return used <= h

    def col(x, w, items):
        size = next((sz for sz in (15, 14, 13, 12, 11, 10.5, 10) if fits(items, w, sz)), 10)
        lines = []
        for h, b in items:
            lines.append([(h, {"bold": True, "color": ACC_D, "size": size})])
            for ln in b.split("\n"):
                if ln.strip():
                    lines.append(ln.strip())
            lines.append("")
        text(s, x, 0.95, w, 6.3, lines[:-1], size=size, color=INK, space=1.2, after=3)

    if right:
        col(0.6, 6.25, left)
        rect(s, 7.05, 1.0, 0.02, 6.0, LINE)
        col(7.3, 5.45, right)
    else:
        col(0.6, 12.1, left)
    s.notes_slide.notes_text_frame.text = body
    return s


def notes(s, body):
    s.notes_slide.notes_text_frame.text = body
    if SCRIPT_PAGES and body.strip():
        script_page(body, LAST["team"], LAST["label"])


def png_of(name):
    src = os.path.join(PREV, name)
    if not os.path.isfile(src):
        return None
    dst = os.path.join(os.path.dirname(OUT), "_tmp_" + os.path.splitext(name)[0] + ".png")
    Image.open(src).convert("RGB").save(dst, "PNG")
    return dst


tmp_files = []

# ---------------- 맨 앞: 연습 안내 ----------------
s = base()
text(s, 0.9, 2.0, 11, 0.4, "TEAM PROJECT FEEDBACK · 발표 연습", size=14, color=ACC_D, bold=True)
text(s, 0.9, 2.5, 11.5, 1.2, "팀 프로젝트 피드백 발표 연습", size=44, color=INK, bold=True)
text(s, 0.9, 3.75, 11.5, 1.2, ["슬라이드 순서 = 피드백 페이지의 '학생 전달용' 화면 순서",
                               "발표자 노트에 실제로 말할 대본 · 쉬운 용어 풀이 · 화면 동작 · 예상 질문"], size=18, color=BODY)
text(s, 0.9, 6.4, 11.5, 0.4, "자료: docs/students/팀피드백.html  ·  2026-11-18 제출본, 코드로 확인한 내용만 사용", size=12, color=FAINT)
notes(s, """[이 PPT 쓰는 법]
- 이 PPT는 학생에게 보여 주는 자료가 아니라, 선생님이 혼자 연습하는 대본입니다.
- 실제 수업에서는 docs/students/팀피드백.html 을 열고 '학생 전달용'을 누른 다음, 이 PPT의 노트를 보면서 같은 순서로 넘기면 됩니다.
- 팀마다 9장: 표지 → 요약 → 잘한 점 1·2·3 → 더 좋아지려면 1·2·3 → 다음 미션.
- 노트의 네 칸
  [말할 내용] 학생들에게 그대로 읽어도 되는 대본
  [쉬운 풀이] 선생님이 먼저 이해하기 위한 용어 설명 (학생에게 말하지 않아도 됨)
  [화면 동작] 피드백 페이지에서 누를 것
  [학생이 물으면] 예상 질문과 답

[연습 팁]
- 처음에는 [쉬운 풀이]부터 읽고, 그다음 [말할 내용]을 소리 내어 두 번 읽어 보세요.
- 모르는 질문이 나오면 "그건 발표 때 직접 보여 줄래요?"라고 학생에게 넘기면 됩니다. 만든 사람이 가장 잘 압니다.
- 피드백 내용은 학생 코드에서 실제로 확인한 사실만 담았습니다. 추측은 넣지 않았습니다.""")

# ---------------- 수업 시작 멘트 ----------------
s = base()
text(s, 0.9, 1.0, 11, 0.4, "수업 시작", size=14, color=ACC_D, bold=True)
text(s, 0.9, 1.5, 11.5, 1.0, "오늘 피드백은 이렇게 진행해요", size=36, color=INK, bold=True)
text(s, 0.9, 2.9, 11.5, 3.5, [[("1  ", {"bold": True, "color": ACC_D}), "작품 소개와 한눈에 보기"],
                              [("2  ", {"bold": True, "color": GREEN}), "잘한 점 3가지"],
                              [("3  ", {"bold": True, "color": BLUE}), "더 좋아지려면 3가지 — Claude에게 바로 줄 수 있는 요청문 포함"],
                              [("4  ", {"bold": True, "color": ACC_D}), "다음 미션"]], size=22, color=INK, after=14)
notes(s, """[말할 내용]
오늘은 팀 프로젝트를 하나씩 같이 볼 거예요. 선생님이 여러분이 낸 코드를 꼼꼼히 읽어 봤어요.
팀마다 순서는 똑같아요. 먼저 어떤 작품인지 한눈에 보고, 그다음 잘한 점 세 가지, 그리고 더 좋아지려면 할 일 세 가지, 마지막으로 다음 미션이에요.
'더 좋아지려면'에는 Claude에게 그대로 붙여 넣을 수 있는 요청 문장도 같이 적어 뒀어요. 수업 끝나고 바로 복사해서 써 보면 돼요.
초록색 화면은 잘한 점, 파란색 화면은 더 좋아지려면이에요. 색만 봐도 지금 어떤 이야기인지 알 수 있어요.
다른 팀 피드백도 같이 들어 주세요. 다른 팀이 잘한 방법이 우리 팀에도 그대로 도움이 될 때가 많아요.

[화면 동작]
- 팀피드백.html 을 열고 위쪽 '학생 전달용'을 누릅니다.
- '전체화면' 버튼(또는 F 키)을 누르면 화면이 꽉 찹니다. 넘길 때는 → 방향키.
- 팀을 바꿀 때는 위쪽 팀 이름 버튼을 누릅니다.

[학생이 물으면]
Q. 점수도 나와요?
A. 오늘은 점수보다 내용을 같이 보려고 해요. (점수를 보여 주려면 '학생 화면에 점수'를 켜면 요약 장에 나옵니다.)""")

# ---------------- 팀별 9장 ----------------
for tid in ORDER:
    t = json.load(io.open(os.path.join(SRC, tid + ".fixed.json"), encoding="utf-8"))
    nt = json.load(io.open(os.path.join(SRC, "notes_" + tid + ".json"), encoding="utf-8"))
    nmap = {n["slide"]: n["notes"] for n in nt}
    who = t["name"] + "  ·  " + MEMBERS[tid]
    goods, imps = t["strengths"][:3], t["improves"][:3]
    chg = t.get("changes") or []          # 완성본이 들어온 팀만 '달라진 점' 1장 (노트 키는 "chg")
    off = 1 if chg else 0                 # 그 뒤 장 번호를 한 칸씩 밀기
    total = 9 + off

    # 1 표지
    s = base(team=who, label="표지  1 / %d" % total)
    text(s, 0.75, 1.55, 6.6, 0.35, "TEAM PROJECT FEEDBACK", size=13, color=ACC_D, bold=True)
    text(s, 0.75, 1.95, 6.6, 1.2, t["name"], size=40, color=INK, bold=True)
    if t.get("one"):   # 표지 소개가 비어 있는 팀은 비워 둡니다
        text(s, 0.75, 3.2, 6.4, 1.4, [clean(x) for x in t["one"].split("\n")], size=17, color=ACC_D, bold=True, after=2)
    text(s, 0.75, 4.75, 6.4, 0.5, MEMBERS[tid], size=15, color=BODY, bold=True)
    img = png_of(PREVIEW[tid]) if tid in PREVIEW else None
    if img:
        tmp_files.append(img)
        rect(s, 7.55, 1.75, 5.05, 0.32, "F1ECE4", line=LINE)
        for i, col in enumerate(["E5675B", "EDB84B", "5DBB63"]):
            rect(s, 7.72 + i * 0.22, 1.85, 0.12, 0.12, col, shape=MSO_SHAPE.OVAL)
        s.shapes.add_picture(img, Inches(7.55), Inches(2.07), Inches(5.05), Inches(5.05 * 600 / 960))
    notes(s, nmap.get(1, ""))

    # 2 요약
    s = base(team=who, label="요약  2 / %d" % total)
    text(s, 0.75, 1.0, 6, 0.35, "한눈에 보기", size=13, color=ACC_D, bold=True)
    sm = [clean(x) for x in t["summary"].split("\n") if x.strip()]
    text(s, 0.75, 1.7, 11.8, 1.0, sm[0], size=30, color=INK, bold=True)
    text(s, 0.95, 2.9, 11.6, 3.6, ["•  " + x for x in sm[1:]], size=20, color=BODY, after=12)
    notes(s, nmap.get(2, ""))

    # (완성본이 있으면) 달라진 점
    if chg:
        s = base(team=who, label="달라진 점  3 / %d" % total)
        r = rect(s, 0.75, 0.95, 1.9, 0.45, "F6E1D3", shape=MSO_SHAPE.ROUNDED_RECTANGLE); r.adjustments[0] = 0.5
        p = r.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        run = p.add_run(); run.text = "★  달라진 점"; run.font.name = FONT; run.font.size = Pt(15); run.font.bold = True; run.font.color.rgb = C("8E3A22")
        text(s, 0.75, 1.75, 5.5, 0.3, "전 버전", size=12, color=FAINT, bold=True)
        text(s, 7.1, 1.75, 5.5, 0.3, "완성본", size=12, color=FAINT, bold=True)
        for k, (b, a) in enumerate(chg[:5]):
            y = 2.15 + k * 0.95
            bx = rect(s, 0.75, y, 5.6, 0.78, "FFFFFF", line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE); bx.adjustments[0] = 0.15
            ax = rect(s, 7.1, y, 5.5, 0.78, "FFFFFF", line="E9B99F", shape=MSO_SHAPE.ROUNDED_RECTANGLE); ax.adjustments[0] = 0.15
            text(s, 0.95, y + 0.17, 5.2, 0.5, clean(b), size=15, color=MUTED)
            text(s, 6.45, y + 0.12, 0.6, 0.5, "→", size=22, color=ACC, bold=True, align=PP_ALIGN.CENTER)
            text(s, 7.3, y + 0.17, 5.1, 0.5, clean(a), size=15, color=INK, bold=True)
        notes(s, nmap.get("chg", ""))

    # 잘한 점
    for i, g in enumerate(goods):
        n = 3 + i
        s = base("good", team=who, label="잘한 점 %d  %d / %d" % (i + 1, n + off, total))
        tag(s, "good")
        text(s, 0.75, 2.2, 3, 0.4, "%d / %d" % (i + 1, len(goods)), size=15, color=GREEN, bold=True)
        text(s, 0.75, 2.7, 11.6, 1.2, clean(g[0]), size=36, color=INK, bold=True)
        desc = [clean(x) for x in re.split(r"(?<=\.)\s+", g[1]) if x.strip()]
        text(s, 0.75, 4.0, 11.4, 2.8, desc, size=19, color=BODY, after=8)
        notes(s, nmap.get(n, ""))

    # 6~8 더 좋아지려면
    for i, m in enumerate(imps):
        n = 6 + i
        s = base("imp", team=who, label="더 좋아지려면 %d  %d / %d" % (i + 1, n + off, total))
        tag(s, "imp")
        text(s, 0.75, 1.75, 3, 0.4, "%d / %d" % (i + 1, len(imps)), size=15, color=BLUE, bold=True)
        text(s, 0.75, 2.2, 11.6, 0.9, clean(m[1]), size=32, color=INK, bold=True)
        text(s, 0.75, 3.3, 1.2, 0.4, "왜", size=14, color=FAINT, bold=True)
        text(s, 1.95, 3.25, 10.6, 0.9, clean(m[2]), size=18, color=INK)
        text(s, 0.75, 4.25, 1.2, 0.4, "어떻게", size=14, color=FAINT, bold=True)
        text(s, 1.95, 4.2, 10.6, 0.7, clean(m[3]), size=18, color=INK)
        box = rect(s, 0.75, 5.2, 11.85, 1.5, "E6ECF6", shape=MSO_SHAPE.ROUNDED_RECTANGLE); box.adjustments[0] = 0.12
        text(s, 1.0, 5.38, 1.2, 0.35, "Claude", size=12, color=BLUE, bold=True)
        text(s, 1.0, 5.72, 11.4, 0.95, m[4], size=15, color=BODY)
        notes(s, nmap.get(n, ""))

    # 9 다음 미션
    s = base(team=who, label="다음 미션  %d / %d" % (9 + off, total))
    rect(s, 0.75, 1.9, 0.08, 4.2, ACC)
    text(s, 1.15, 1.95, 6, 0.4, "다음 미션", size=14, color=ACC_D, bold=True)
    steps = [clean(x) for x in t["mission"].split("\n") if x.strip()]
    text(s, 1.15, 2.6, 11.2, 3.8, [[("%d   " % (k + 1), {"color": ACC_D}), x] for k, x in enumerate(steps)],
         size=24, color=INK, bold=True, after=16)
    notes(s, nmap.get(9, ""))

# ---------------- 마무리 ----------------
s = base()
text(s, 0.9, 2.3, 11, 0.4, "마무리", size=14, color=ACC_D, bold=True)
text(s, 0.9, 2.8, 11.5, 1.0, "다음 시간까지", size=40, color=INK, bold=True)
text(s, 0.9, 4.0, 11.5, 2.0, ["팀별 '더 좋아지려면' 1번부터 Claude에게 요청해 보기",
                              "바꾼 결과와 사용한 요청문을 개발 기록에 남기기"], size=20, color=BODY, after=10)
notes(s, """[말할 내용]
오늘 다섯 팀을 모두 봤어요. 공통으로 보인 건, 다들 Claude에게 일을 시키는 수준을 넘어서 구조를 고민한 흔적이 코드에 남아 있다는 거예요.
다음 시간까지 각 팀은 '더 좋아지려면' 1번부터 하나씩 해 보세요. 화면에 있는 Claude 요청문을 그대로 복사해서 시작해도 되고, 여러분 말로 바꿔도 돼요.
중요한 건 무엇을 바꿨는지 기록하는 거예요. 어떤 요청을 했고, 결과가 어떻게 달라졌는지 개발 기록에 남겨 주세요.
궁금한 게 있으면 수업 끝나고 팀별로 물어보러 오세요.

[화면 동작]
- 전체화면은 Esc 로 나옵니다.
- 팀별 피드백을 학생에게 보내려면 '이 팀 복사'(글로 복사) 또는 '인쇄 · PDF'를 씁니다.

[학생이 물으면]
Q. 4번, 5번 개선점도 있어요?
A. 있어요. 오늘은 먼저 할 세 가지만 골랐어요. 세 개를 다 하면 선생님이 다음 단계를 알려 줄게요.""")

prs.save(OUT)
for f in tmp_files:
    try: os.remove(f)
    except OSError: pass
print("저장:", OUT, "· 슬라이드", len(prs.slides._sldIdLst))
