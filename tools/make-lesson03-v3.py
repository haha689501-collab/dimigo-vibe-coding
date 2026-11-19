# -*- coding: utf-8 -*-
"""
3회기 v3 PPT 생성기 — 「같은 설계도, 다른 작품」 (관계자 참관 수업용)

실행:  python tools/make-lesson03-v3.py
결과:  디미고_바이브코딩_3차시_같은설계도다른작품_v3.pptx  (프로젝트 최상위, 새 파일)

- 기존 PPT는 열지도 수정하지도 않습니다.
- 카드·배너·단계 줄 등 부품은 v2 생성기(tools/make-lesson03-v2.py)의 것을 그대로 씁니다.
- 발표자 노트는 tools/lesson03_v3_script.py 에서 가져옵니다.
- 저장 후 모든 글상자의 넘침과 슬라이드 밖으로 나간 도형을 자동으로 검사합니다.
"""
import importlib.util
import os
import sys

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

_spec = importlib.util.spec_from_file_location("l3v2", os.path.join(HERE, "make-lesson03-v2.py"))
B = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(B)

import lesson03_v3_script as SC

OUT = os.path.join(ROOT, SC.DECK)

ML, CW, SW, SH, TOP, FOOT_Y = B.ML, B.CW, B.SW, B.SH, B.TOP, B.FOOT_Y
INK, BODY, MUTED, FAINT = B.INK, B.BODY, B.MUTED, B.FAINT
ACCENT, ACCENT_D, ACCENT_L = B.ACCENT, B.ACCENT_D, B.ACCENT_L
LINE, LINE2, CARD, SOFT, TINT = B.LINE, B.LINE2, B.CARD, B.SOFT, B.TINT
DARK, WHITE, BLUE, GREEN, PURPLE = B.DARK, B.WHITE, B.BLUE, B.GREEN, B.PURPLE
KR, MONO = B.KR, B.MONO
rect, text, header, footer_rule, title_block, banner = B.rect, B.text, B.header, B.footer_rule, B.title_block, B.banner
wcard, dcard, codebox, step_row, part_slide = B.wcard, B.dcard, B.codebox, B.step_row, B.part_slide


def oval(s, x, y, d, fill=None, line=None, lw=1.2):
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    if fill is None:
        c.fill.background()
    else:
        c.fill.solid(); c.fill.fore_color.rgb = fill
    if line is None:
        c.line.fill.background()
    else:
        c.line.color.rgb = line; c.line.width = Pt(lw)
    c.shadow.inherit = False
    return c


def cols(n, gap):
    w = (CW - gap * (n - 1)) / n
    return [ML + i * (w + gap) for i in range(n)], w


# ================================================================ 슬라이드
def build():
    prs = B.new_deck()
    S = lambda: B.blank(prs)

    # ---------------------------------------------------------- 1. 표지
    s = S()
    text(s, ML, 0.44, 5, 0.3, [[("A I   ×   B U I L D", {"bold": True, "size": 11, "color": ACCENT, "space": 1.3})]])
    text(s, SW - ML - 7.4, 0.30, 7.4, 0.7,
         [[("AI 노코드 창작 클래스", {"bold": True, "size": 30, "color": LINE2})]], align=PP_ALIGN.RIGHT)
    rect(s, ML, 1.06, CW, 0.018, fill=LINE)
    cx = ML + 0.28
    for i, (num, nm) in enumerate([("1", "IDEA"), ("2", "SPEC"), ("3", "BUILD"), ("4", "TEST")]):
        on = (i == 2)
        d = 0.62
        oval(s, cx, 1.48, d, fill=ACCENT if on else RGBColor(0xEF, 0xE5, 0xDC))
        text(s, cx, 1.62, d, 0.3, [[(num, {"bold": True, "size": 12, "color": WHITE if on else ACCENT_D})]],
             align=PP_ALIGN.CENTER)
        text(s, cx - 0.2, 2.22, d + 0.4, 0.26,
             [[(nm, {"bold": True, "size": 8.5, "color": ACCENT_D if on else FAINT, "space": 0.9})]],
             align=PP_ALIGN.CENTER)
        if i < 3:
            rect(s, cx + d, 1.77, 0.42, 0.012, fill=LINE2)
        cx += d + 0.42
    text(s, ML, 2.86, 4, 1.18, [[("03", {"bold": True, "size": 58, "color": INK})]])
    text(s, ML, 4.00, 10.5, 1.3,
         [[("같은 설계도, 다른 작품", {"bold": True, "size": 36, "color": INK})],
          [("우리 조가 끝까지 키울 작품을 찾습니다", {"size": 18, "color": BODY})]], line=1.35)
    rect(s, ML, 5.56, 1.5, 0.05, fill=ACCENT)
    text(s, ML, 5.86, 11, 0.32,
         [[("90분 · 창작 유형 탐색 · 오늘의 산출물 : 내 첫 화면 + 우리 조 작품 후보", {"size": 12.5, "color": MUTED})]])
    text(s, SW - ML - 7, 6.62, 7, 0.3,
         [[("한국디지털미디어고등학교 · Vibe Coding Project", {"size": 10.5, "color": FAINT})]], align=PP_ALIGN.RIGHT)

    # ---------------------------------------------------------- 2. 오늘의 목표
    s = S(); header(s)
    title_block(s, "오늘의 목표", "같은 설계도로 만든 결과를 비교하고, 잘 만든 작품의 조건을 찾고, 함께 키울 작품을 고릅니다.")
    xs, w = cols(3, 0.26)
    wcard(s, xs[0], TOP, w, 2.75, BLUE, "01 · COMPARE", "설계도가 작품을 만든다",
          ["같은 주제 PRD로 첫 화면 만들기", "조 친구들 화면과 나란히 보기", "차이를 만든 PRD 문장 찾기"])
    wcard(s, xs[1], TOP, w, 2.75, GREEN, "02 · ANALYZE", "잘 만든 작품을 뜯어본다",
          ["같은 게임의 기본판 · 완성판 비교", "완성돼 보이는 이유 찾기", "우리 반 완성 기준 세우기"])
    wcard(s, xs[2], TOP, w, 2.75, PURPLE, "03 · TEAM", "함께 키울 작품을 고른다",
          ["조와 역할 알아보기", "작품 후보 세 개 만들기", "30초 피치와 동료 피드백"])
    banner(s, [[("연결되는 성취 기준   ", {"bold": True, "color": ACCENT_D}),
                ("① 원하는 결과를 말로 끌어낸다    ② 결과를 살펴 더 낫게 만들 방향을 찾는다", {})]])
    footer_rule(s)

    # ---------------------------------------------------------- 3. 10회기 로드맵
    s = S(); header(s)
    title_block(s, "우리가 가는 길 — 10회기 로드맵",
                "오늘부터는 방법을 배우는 시간에서, 작품 하나를 끝까지 키우는 시간으로 바뀝니다.")
    nodes = [
        ("바이브코딩", "오리엔테이션", "done"), ("설계도(PRD)", "프롬프트·개선", "done"),
        ("탐색·조 결성", "창작 유형 탐색", "today"), ("핵심 재미", "프로젝트 기획", ""),
        ("첫 배포", "제작 집중 ①", ""), ("한 판 완성", "제작 집중 ②", ""),
        ("플레이테스트", "중간 점검", ""), ("그림·소리", "디자인·다듬기", ""),
        ("출시 준비", "발표 준비", ""), ("데모데이", "쇼케이스", "final"),
    ]
    colw = CW / 10.0
    d = 0.56
    cy = 2.62
    rect(s, ML + colw / 2, cy + d / 2 - 0.015, colw * 9, 0.03, fill=LINE2)
    for i, (lab, plan, st) in enumerate(nodes):
        cxn = ML + colw * i + colw / 2
        if st == "done":
            oval(s, cxn - d / 2, cy, d, fill=RGBColor(0xE4, 0xDC, 0xD1))
            numc = MUTED
        elif st == "today":
            oval(s, cxn - d / 2 - 0.06, cy - 0.06, d + 0.12, fill=ACCENT)
            numc = WHITE
            rect(s, cxn - 0.42, cy - 0.44, 0.84, 0.27, fill=ACCENT_D, radius=0.5)
            text(s, cxn - 0.42, cy - 0.415, 0.84, 0.24, [[("오늘", {"bold": True, "size": 9.5, "color": WHITE})]],
                 align=PP_ALIGN.CENTER)
        elif st == "final":
            oval(s, cxn - d / 2, cy, d, fill=DARK)
            numc = WHITE
        else:
            oval(s, cxn - d / 2, cy, d, fill=CARD, line=LINE2)
            numc = ACCENT_D
        text(s, cxn - d / 2, cy + 0.14, d, 0.3, [[(str(i + 1), {"bold": True, "size": 13, "color": numc})]],
             align=PP_ALIGN.CENTER)
        text(s, cxn - colw / 2 + 0.03, cy + d + 0.14, colw - 0.06, 0.3,
             [[(lab, {"bold": True, "size": 10.5, "color": INK if st != "done" else MUTED})]], align=PP_ALIGN.CENTER)
        text(s, cxn - colw / 2 + 0.03, cy + d + 0.46, colw - 0.06, 0.28,
             [[(plan, {"size": 8.5, "color": FAINT})]], align=PP_ALIGN.CENTER)
    bands = [(0, 2, "방법 익히기", RGBColor(0xEF, 0xEB, 0xE5), MUTED),
             (2, 4, "작품 정하기", RGBColor(0xF6, 0xE9, 0xE1), ACCENT_D),
             (4, 6, "만들기", RGBColor(0xF2, 0xED, 0xE7), BODY),
             (6, 8, "고치고 다듬기", RGBColor(0xF2, 0xED, 0xE7), BODY),
             (8, 10, "선보이기", RGBColor(0xE6, 0xE8, 0xF2), DARK)]
    by = 4.28
    for a, b, lab, fill, col in bands:
        x = ML + colw * a + 0.05
        rect(s, x, by, colw * (b - a) - 0.1, 0.46, fill=fill, radius=0.3)
        text(s, x, by + 0.11, colw * (b - a) - 0.1, 0.26, [[(lab, {"bold": True, "size": 11, "color": col})]],
             align=PP_ALIGN.CENTER)
    text(s, ML, by + 0.56, CW, 0.26,
         [[("각 칸 아래 작은 글씨 = 학교 제출 지도안의 회차 이름", {"size": 9, "color": FAINT})]], align=PP_ALIGN.RIGHT)
    banner(s, [[("10회기 데모데이 — ", {"bold": True, "color": ACCENT_D}),
                ("QR 하나로 누구나 휴대폰에서 우리 조의 작품을 해 봅니다.", {})]], y=5.42, h=0.9)
    footer_rule(s)

    # ---------------------------------------------------------- 4. 90분 흐름
    s = S(); header(s)
    title_block(s, "오늘 90분의 흐름", "설명은 짧게, 만들고 비교하고 이야기하는 시간은 길게 가져갑니다.")
    flow = [("00-15", "OPEN", "목표 · 규칙\n트랙 · 요청문", False),
            ("15-35", "BUILD", "같은 PRD로\n첫 화면 만들기", True),
            ("35-47", "COMPARE", "조별로 나란히\n무엇이 달랐나", True),
            ("47-65", "ANALYZE", "잘 만든 작품\n뜯어보기", True),
            ("65-87", "TEAM", "조 · 역할 · 후보\n30초 피치", True),
            ("87-90", "WRAP", "정리와\n4회기 예고", False)]
    xs, w = cols(6, 0.16)
    for i, (t, lab, desc, hot) in enumerate(flow):
        x = xs[i]
        rect(s, x, TOP, w, 2.62, fill=CARD, line=ACCENT_L if hot else LINE, lw=1.4 if hot else 1.0, radius=0.06)
        rect(s, x, TOP, w, 0.075, fill=ACCENT if hot else LINE2)
        text(s, x + 0.20, TOP + 0.30, w - 0.4, 0.36, [[(t, {"bold": True, "size": 16, "color": INK})]])
        text(s, x + 0.20, TOP + 0.74, w - 0.4, 0.26,
             [[(lab, {"bold": True, "size": 9.5, "color": ACCENT_D if hot else FAINT, "space": 1.0})]])
        rect(s, x + 0.20, TOP + 1.06, w - 0.4, 0.012, fill=LINE)
        text(s, x + 0.20, TOP + 1.24, w - 0.4, 1.2, [[(l, {})] for l in desc.split("\n")],
             size=11, color=BODY, line=1.5)
    banner(s, [[("지금 준비하세요 — 지난 시간 ", {}), ("PRD", {"bold": True, "color": ACCENT_D}),
                (" 열기 · Claude 로그인 · 우리 조 게시판 열어 두기", {})],
               [("PRD가 없다면 지금 손을 드세요. 예시 PRD를 드립니다.", {"size": 11, "color": MUTED})]],
           y=5.30, h=1.05)
    footer_rule(s)

    # ---------------------------------------------------------- 5. PART 1
    s = S(); part_slide(s, "PART 1", "같은 설계도, 다른 작품", "같은 주제로 쓴 PRD라면 · 결과도 같을까요?", 1, 3)

    # ---------------------------------------------------------- 6. 오늘의 질문
    s = S(); header(s)
    title_block(s, "오늘의 질문 — 같은 주제인데 왜 다른 작품이 나올까?",
                "우리는 모두 같은 주제로 설계도를 썼습니다. 그런데 결과는 같을까요?")
    xs, w = cols(3, 0.26)
    wcard(s, xs[0], TOP, w, 2.90, LINE2, "같은 것", "주제 · 도구 · 시간",
          ["똑같은 주제로 시작했다", "똑같이 Claude를 쓴다", "똑같이 20분을 쓴다"], note="출발점은 모두 같습니다.")
    wcard(s, xs[1], TOP, w, 2.90, ACCENT, "다른 것", "내가 내린 결정",
          ["어떤 불편을 골랐는지", "누구를 위한 것인지", "무엇을 넣고 뺐는지"], strong=True, note="이것이 PRD에 적혀 있습니다.")
    wcard(s, xs[2], TOP, w, 2.90, LINE2, "결과", "화면",
          ["화면에 놓이는 것", "놓이는 순서", "화면에 없는 것"], note="20분 뒤 눈으로 확인합니다.")
    banner(s, [[("옆 친구 화면이 더 멋져 보여도 괜찮습니다. 그건 ", {}),
                ("실력의 차이가 아니라 결정의 차이", {"bold": True, "color": ACCENT_D}),
                ("입니다.", {})]], y=5.48, h=0.88)
    footer_rule(s)

    # ---------------------------------------------------------- 7. 규칙 세 가지
    s = S(); header(s)
    title_block(s, "PRD가 화면이 되는 규칙 세 가지",
                "PRD의 어느 줄이 화면의 어느 부분이 되는지 알면, 요청문을 정확하게 쓸 수 있습니다.")
    xs, w = cols(3, 0.26)
    wcard(s, xs[0], TOP, w, 2.95, BLUE, "규칙 ①  MVP 기능", "화면 요소가 된다",
          ["‘이번에 만들 기능’ 하나 = 화면 한 덩어리", "과제 등록 → 입력칸 + 등록 버튼", "기능 목록이 곧 화면 목록"])
    wcard(s, xs[1], TOP, w, 2.95, GREEN, "규칙 ②  사용 흐름", "배치 순서가 된다",
          ["먼저 하는 일이 화면 위쪽에", "순서를 안 주면 AI가 마음대로", "‘위에서부터 이 순서로 배치해줘’"])
    wcard(s, xs[2], TOP, w, 2.95, ACCENT, "규칙 ③  만들지 않을 것", "금지 목록이 된다",
          ["안 적으면 AI가 친절하게 덧붙인다", "로그인 · 설정이 딸려 온다", "코드가 길어져 확인할 수 없다"], strong=True)
    banner(s, [[("PRD 양식마다 번호가 다릅니다 — ", {}),
                ("번호 말고 ‘MVP’, ‘사용 흐름’, ‘만들지 않을 것’처럼 이름으로 찾으세요.", {"bold": True, "color": ACCENT_D})]],
           y=5.48, h=0.88)
    footer_rule(s)

    # ---------------------------------------------------------- 8. 트랙
    s = S(); header(s)
    title_block(s, "오늘의 트랙을 스스로 고르세요",
                "목표는 모두 같습니다. 가는 길만 다릅니다. 중간에 올려도 내려도 됩니다.")
    xs, w = cols(3, 0.26)
    dcard(s, xs[0], TOP, w, 2.95, "TRACK A", BLUE, "따라 만들기",
          ["처음 만들어 본다", "게시판 요청문을 그대로 쓴다", "화면이 뜨면 목표 달성"],
          note="여기까지만 해도 오늘 목표는 100% 달성입니다.")
    dcard(s, xs[1], TOP, w, 2.95, "TRACK B", GREEN, "직접 설계하기",
          ["요청문을 내가 쓴다", "빈 화면 · 긴 제목까지 정한다", "휴대폰 크기로 먼저 설계"],
          note="HTML을 조금 아는 학생에게 권합니다.")
    dcard(s, xs[2], TOP, w, 2.95, "TRACK C", PURPLE, "따져 묻기",
          ["조건을 짧고 분명하게 건다", "결과를 PRD와 대조해 되묻는다", "더 적은 요청으로 같은 결과"],
          note="많이가 아니라 정확히 만들기입니다.")
    banner(s, [[("빨리 끝내는 것이 잘한 것이 아닙니다. ", {}),
                ("내 화면이 왜 이렇게 나왔는지 설명할 수 있는 것", {"bold": True, "color": ACCENT_D}),
                ("이 잘한 것입니다.", {})]])
    footer_rule(s)

    # ---------------------------------------------------------- 9. 요청문
    s = S(); header(s)
    title_block(s, "요청문 — 이대로 쓰고, 필요하면 더 얹기",
                "게시판의 요청문을 복사해 PRD를 끼워 넣으세요. 조건 다섯 줄이 AI의 추측을 줄여 줍니다.")
    lw_ = CW * 0.58
    codebox(s, ML, TOP, lw_, 3.55, "TRACK A · 요청문", [
        "아래는 내가 만든 PRD야. 이걸 보고 첫 화면을 만들어줘.",
        "",
        "[ 여기에 PRD 내용을 통째로 붙여넣기 ]",
        "",
        "조건:",
        "- HTML, CSS, JavaScript를 한 화면(한 파일)으로",
        "- MVP에 적은 기능만",
        "- 만들지 않기로 한 것은 절대 만들지 마",
        "- 사용 흐름 순서대로 위에서부터 배치",
        "- 오늘은 화면과 디자인만. 기능은 아직 연결하지 마",
    ])
    bx = ML + lw_ + 0.3
    wcard(s, bx, TOP, CW - lw_ - 0.3, 3.55, ACCENT, "TRACK B · C", "이렇게 더 얹기",
          ["B  목록이 비었을 때 안내 화면", "B  긴 제목도 깨지지 않게", "B  휴대폰 390px 먼저",
           "C  조건을 영어로 짧게 (한국어도 OK)", "C  ‘PRD와 다른 점을 근거와 함께 알려줘. 고치지는 마’"],
          strong=True, size=11)
    footer_rule(s)

    # ---------------------------------------------------------- 10. 실습 ①
    s = S(); header(s)
    title_block(s, "실습 ① 같은 주제 PRD로 첫 화면 — 20분",
                "예쁘게보다 정확하게. 오늘의 기준은 ‘내 설계도대로 나왔는가’입니다.")
    step_row(s, TOP, [
        ("STEP 1 · 0–3분", "보낸다", "PRD와 요청문을 함께 붙여넣고 보낸다"),
        ("STEP 2 · 3–8분", "확인한다", "내 PRD의 MVP 기능이 화면에 다 보이는지 본다"),
        ("STEP 3 · 8–16분", "한 번 고친다", "딱 한 곳만 — 어디를 · 무엇을 · 어떻게"),
        ("STEP 4 · 16–20분", "올린다", "화면을 캡처해 우리 조 게시판에 올린다"),
    ], h=1.85)
    ty = TOP + 2.08
    rect(s, ML, ty, CW, 1.02, fill=SOFT, line=LINE, radius=0.08)
    text(s, ML + 0.3, ty + 0.12, 2.2, 0.3, [[("막히면 이렇게", {"bold": True, "size": 11, "color": ACCENT_D, "space": 0.6})]])
    tips = [("화면 대신 코드만 나오면", "‘미리보기로 보여줘’"),
            ("너무 길고 느리면", "‘기능은 빼고 화면만, 짧게’"),
            ("로그인 · 설정이 생기면", "‘만들지 않을 것’에 추가하고 다시")]
    tw = (CW - 0.6) / 3
    for i, (a, b) in enumerate(tips):
        x = ML + 0.3 + i * tw
        text(s, x, ty + 0.44, tw - 0.2, 0.52,
             [[(a, {"size": 11, "color": MUTED})], [("→ " + b, {"size": 11.5, "bold": True, "color": INK})]], line=1.3)
    banner(s, [[("10분 뒤에 한 번 멈춥니다 — ", {}),
                ("화면이 뜬 사람은 손을 들어 주세요.", {"bold": True, "color": ACCENT_D}),
                ("  혼자 오래 붙잡고 있지 마세요.", {})]], y=5.66, h=0.8)
    footer_rule(s)

    # ---------------------------------------------------------- 11. 조별 비교
    s = S(); header(s)
    title_block(s, "조별로 나란히 — 무엇이 달랐나 (10분)",
                "한 사람당 1분씩 화면을 보여 주고, 두 질문에만 답합니다. 잘했다 · 못했다는 말하지 않습니다.")
    hw = (CW - 0.3) / 2
    wcard(s, ML, TOP, hw, 2.60, ACCENT, "질문 1", "가장 크게 달랐던 점은?",
          ["무엇이 놓였나 · 어떤 순서인가 · 무엇이 없나", "조에서 가장 큰 차이 하나를 고른다"], strong=True)
    wcard(s, ML + hw + 0.3, TOP, hw, 2.60, ACCENT, "질문 2", "그 차이를 만든 PRD 문장은?",
          ["PRD에서 손가락으로 짚는다", "찾았다 → 내가 내린 결정", "못 찾았다 → AI가 대신 내린 결정"], strong=True)
    banner(s, [[("조별 한 줄 발표 (8분 뒤)", {"bold": True, "color": INK})],
               [("“우리 조에서 가장 달랐던 건 ___이고, 그건 PRD의 ___ 때문이었다.”",
                 {"bold": True, "color": ACCENT_D, "size": 13.5})]], y=5.16, h=1.2)
    footer_rule(s)

    # ---------------------------------------------------------- 12. 오늘의 발견 ①
    s = S(); header(s)
    text(s, ML, 1.72, CW, 0.34, [[("오늘의 발견 ①", {"bold": True, "size": 12, "color": ACCENT, "space": 1.0})]])
    text(s, ML, 2.14, CW, 0.95, [[("설계도가 작품을 결정합니다", {"bold": True, "size": 40, "color": INK})]])
    rect(s, ML, 3.22, 0.9, 0.05, fill=ACCENT)
    text(s, ML, 3.48, CW, 0.5,
         [[("같은 주제, 같은 AI, 같은 20분. 달랐던 것은 여러분이 내린 결정뿐이었습니다.", {"size": 16, "color": BODY})]])
    rect(s, ML, 4.72, CW, 1.12, fill=TINT, radius=0.1)
    text(s, ML + 0.45, 4.86, CW - 0.9, 0.86,
         [[("그렇다면 —", {"bold": True, "size": 12.5, "color": ACCENT_D})],
          [("7주 뒤 데모데이에서 멋진 작품을 보여 주려면, 우리는 무엇을 결정해야 할까요?",
            {"bold": True, "size": 15, "color": INK})]], line=1.45, anchor=MSO_ANCHOR.MIDDLE)
    footer_rule(s)

    # ---------------------------------------------------------- 13. PART 2
    s = S(); part_slide(s, "PART 2", "잘 만든 작품은 무엇이 다른가",
                        "데모데이에서 무엇으로 평가받는지 보고 · 직접 해 보며 찾습니다", 2, 3)

    # ---------------------------------------------------------- 14. 데모데이 기준
    s = S(); header(s)
    title_block(s, "데모데이 미리보기 — 무엇으로 평가받나",
                "끝을 먼저 봅니다. 7주 뒤 우리 조의 작품은 이 네 가지로 평가받습니다.")
    xs, w = cols(4, 0.22)
    items = [(ACCENT, "재미", "해 보고 싶은가", ["30초 안에 규칙을 안다", "계속하고 싶게 만든다", "점점 어려워진다"]),
             (BLUE, "완성도 · 안정성", "끝까지 되는가", ["처음부터 끝까지 멈추지 않는다", "다시하기 · 저장이 된다", "휴대폰에서도 된다"]),
             (GREEN, "AI 활용의 질", "어떻게 시켰나", ["요청문과 버전을 기록한다", "AI가 틀린 순간을 바로잡았다", "그림 · 소리 출처를 밝힌다"]),
             (PURPLE, "협업", "함께 만들었나", ["역할을 나눴다", "각자의 기여가 보인다", "서로 테스트하고 고쳤다"])]
    for i, (col, lab, head, lines) in enumerate(items):
        wcard(s, xs[i], TOP, w, 2.95, col, lab, head, lines, size=11)
    banner(s, [[("대학생 해커톤 ", {}), ("‘AI 인디게임 챌린지 2026’", {"bold": True, "color": ACCENT_D}),
                ("(서일대학교, 2026년 9월)도 이 네 가지로 심사했습니다.", {})]], y=5.48, h=0.88)
    footer_rule(s)

    # ---------------------------------------------------------- 15. 시연
    s = S(); header(s)
    title_block(s, "시연 — 규칙은 같은데, 왜 한쪽만 ‘게임’처럼 보일까?",
                "같은 게임을 기본판과 완성판으로 해 봅니다. 보면서 여섯 칸을 떠올리세요.")
    lw2 = 4.35
    dcard(s, ML, TOP, lw2, 2.95, "지금 화면에 띄웁니다", ACCENT_L, "별똥별 피하기",
          ["별은 모으고, 운석은 피한다", "두 버전의 규칙은 완전히 같다", "자원자 두 명이 직접 플레이"],
          note="docs/demo/star-dodge/index.html")
    rx = ML + lw2 + 0.3
    rw = CW - lw2 - 0.3
    lens = SC.APPENDIX_LENS
    cw_ = (rw - 0.2 * 2) / 3
    ch = 1.37
    for i, (lab, q) in enumerate(lens):
        r, c = divmod(i, 3)
        x = rx + c * (cw_ + 0.2)
        y = TOP + r * (ch + 0.21)
        rect(s, x, y, cw_, ch, fill=CARD, line=LINE, radius=0.08)
        rect(s, x, y, 0.07, ch, fill=ACCENT)
        text(s, x + 0.24, y + 0.18, cw_ - 0.4, 0.34, [[(lab, {"bold": True, "size": 14, "color": INK})]])
        text(s, x + 0.24, y + 0.56, cw_ - 0.4, 0.72, [[(q, {"size": 10.5, "color": MUTED})]], line=1.35)
    banner(s, [[("보면서 생각해 보세요 — ", {}),
                ("무엇이 한쪽을 ‘게임’처럼 보이게 만들었을까요?", {"bold": True, "color": ACCENT_D})]], y=5.48, h=0.88)
    footer_rule(s)

    # ---------------------------------------------------------- 16. 뜯어보기 활동
    s = S(); header(s)
    title_block(s, "작품 뜯어보기 — 게임답게 만든 이유 TOP 3 (5분)",
                "각자 본 것을 조별로 모아, 완성판을 ‘게임답게’ 만든 이유 세 가지를 고릅니다.")
    step_row(s, TOP, [
        ("1 · 각자", "떠올리기", "여섯 칸 중에서 눈에 띈 것을 적는다"),
        ("2 · 조별", "세 가지 고르기", "‘게임답게’ 만든 이유 TOP 3"),
        ("3 · 게시판", "올리기", "우리 조 게시판에 TOP 3를 올린다"),
    ], h=1.7)
    xs6, w6 = cols(6, 0.15)
    hints = ["뭘 할지 알았나", "누르면 어떤 느낌", "있고 없고의 차이", "점점 어렵게", "다시 하고 싶게", "휴대폰에서도"]
    py = TOP + 1.98
    for i, (lab, _) in enumerate(SC.APPENDIX_LENS):
        rect(s, xs6[i], py, w6, 0.95, fill=SOFT, line=LINE, radius=0.12)
        text(s, xs6[i], py + 0.14, w6, 0.32, [[(lab, {"bold": True, "size": 13, "color": INK})]], align=PP_ALIGN.CENTER)
        text(s, xs6[i] + 0.08, py + 0.52, w6 - 0.16, 0.3, [[(hints[i], {"size": 10, "color": MUTED})]], align=PP_ALIGN.CENTER)
    banner(s, [[("TRACK C는 하나 더 — ", {}),
                ("그중 하나를 AI에게 부탁하는 요청문 한 줄", {"bold": True, "color": ACCENT_D}),
                ("도 함께 적어 보세요.", {})]], y=5.58, h=0.82)
    footer_rule(s)

    # ---------------------------------------------------------- 17. 반 기준 초안
    s = S(); header(s)
    title_block(s, "우리 반 완성 기준 — 초안",
                "조마다 하나씩 모아 네 칸에 적습니다. 마지막에 시연 게임의 ‘달라진 점’과 맞춰 봅니다.")
    xs, w = cols(4, 0.2)
    board = [(BLUE, "완성도 · 안정성", "예) 끝나면 다시하기가 있다"),
             (ACCENT, "재미", "예) 점점 빨라진다"),
             (GREEN, "반응 (손맛)", "예) 부딪히면 화면이 흔들린다"),
             (PURPLE, "어디서나", "예) 휴대폰으로도 된다")]
    bh = 2.95
    for i, (col, lab, ex) in enumerate(board):
        x = xs[i]
        rect(s, x, TOP, w, bh, fill=CARD, line=LINE, radius=0.06)
        rect(s, x, TOP, w, 0.075, fill=col)
        text(s, x + 0.24, TOP + 0.26, w - 0.48, 0.36, [[(lab, {"bold": True, "size": 15, "color": INK})]])
        text(s, x + 0.24, TOP + 0.70, w - 0.48, 0.3, [[(ex, {"size": 10.5, "color": FAINT})]])
        for k in range(4):
            rect(s, x + 0.24, TOP + 1.30 + k * 0.40, w - 0.48, 0.012, fill=LINE)
    banner(s, [[("이 목록이 4회기에 ", {}), ("우리 반의 ‘데모데이 합격선’", {"bold": True, "color": ACCENT_D}),
                ("이 됩니다 — 선생님이 정한 기준이 아니라, 여러분이 찾은 기준으로.", {})]], y=5.48, h=0.88)
    footer_rule(s)

    # ---------------------------------------------------------- 18. PART 3
    s = S(); part_slide(s, "PART 3", "우리 조가 만들 작품", "혼자 하면 과제가 되고 · 함께 하면 작품이 됩니다", 3, 3)

    # ---------------------------------------------------------- 19. 역할
    s = S(); header(s)
    title_block(s, "조와 역할 다섯 가지",
                "작품을 끝까지 만들려면 역할이 필요합니다. 코딩 경험이 없어도 어떤 역할이든 할 수 있습니다.")
    xs, w = cols(5, 0.18)
    roles = [(ACCENT, "ROLE 1", "디렉터", ["무엇을 만들지 정한다", "범위를 지킨다", "결정을 기록한다"]),
             (BLUE, "ROLE 2", "플레이", ["핵심 규칙을 만든다", "AI와 함께 구현", "동작을 확인한다"]),
             (GREEN, "ROLE 3", "그림", ["분위기를 정한다", "캐릭터 · 배경", "화면 꾸미기"]),
             (PURPLE, "ROLE 4", "소리 · 연출", ["효과음 · 배경음", "흔들림 · 파티클", "손맛 만들기"]),
             (LINE2, "ROLE 5", "테스트 · 출시", ["잘 되는지 확인", "버그 기록", "인터넷에 올리기"])]
    for i, (col, lab, head, lines) in enumerate(roles):
        wcard(s, xs[i], TOP, w, 2.80, col, lab, head, lines, size=11)
    banner(s, [[("역할은 4회기에 정합니다 — ", {}),
                ("오늘은 해 보고 싶은 역할을 1지망 · 2지망으로 게시판에 적어 주세요.", {"bold": True, "color": ACCENT_D})]],
           y=5.40, h=0.9)
    footer_rule(s)

    # ---------------------------------------------------------- 20. 작품 유형
    s = S(); header(s)
    title_block(s, "작품 유형 — 무엇을 만들 수 있나",
                "세 가지 중에서 고릅니다. 어떤 것을 고르든 기준은 하나입니다.")
    xs, w = cols(3, 0.26)
    wcard(s, xs[0], TOP, w, 2.80, ACCENT, "유형 1 · 추천", "게임",
          ["QR 찍고 바로 플레이", "데모데이 인상이 가장 강하다", "한 판 3분 이내"], strong=True)
    wcard(s, xs[1], TOP, w, 2.80, BLUE, "유형 2", "실제로 쓰는 도구",
          ["학교생활의 진짜 불편 해결", "처음 써도 바로 이해", "입력한 내용이 저장된다"])
    wcard(s, xs[2], TOP, w, 2.80, PURPLE, "유형 3", "미디어아트",
          ["움직이고 반응하는 작품", "보는 사람이 참여한다", "소리 · 빛 · 움직임"])
    banner(s, [[("공통 기준 하나 — ", {}),
                ("처음 보는 사람이 QR을 찍고 30초 안에 즐길 수 있는가", {"bold": True, "color": ACCENT_D})]],
           y=5.40, h=0.9)
    footer_rule(s)

    # ---------------------------------------------------------- 21. 한 문장 피치
    s = S(); header(s)
    title_block(s, "한 문장 피치 — 작고 선명하게",
                "아이디어는 한 문장으로 말합니다. 작고 선명한 아이디어가 끝까지 갑니다.")
    rect(s, ML, TOP, CW, 0.92, fill=TINT, radius=0.12)
    text(s, ML + 0.4, TOP + 0.12, CW - 0.8, 0.7,
         [[("[누가]  [언제]  [무엇을 하는]  [게임 · 도구 · 아트]", {"bold": True, "size": 17, "color": INK}),
           (".   핵심 재미는  [          ].", {"bold": True, "size": 17, "color": ACCENT_D})]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cy2 = TOP + 1.14
    hw = (CW - 0.3) / 2
    wcard(s, ML, cy2, hw, 3.08, ACCENT, "좋은 예", "급식 밸런스 캐치",
          ["쉬는 시간 3분 동안,", "떨어지는 급식 메뉴를 받아 영양 균형을 맞추는 게임.",
           "핵심 재미 — 먹고 싶은 것과 먹어야 하는 것 사이의 줄타기"],
          strong=True, size=11, note="누가 · 언제 · 무엇 · 재미가 모두 들어 있습니다.")
    wcard(s, ML + hw + 0.3, cy2, hw, 3.08, LINE2, "아쉬운 예", "“재밌는 RPG 게임”",
          ["누가 · 언제 하는지 없음", "무엇이 재밌는지 없음", "7주 안에 만들기엔 너무 큼"],
          size=11, note="작고 선명한 게 이깁니다 — 한 판 3분 크기로.")
    footer_rule(s)

    # ---------------------------------------------------------- 22. 조별 활동
    s = S(); header(s)
    title_block(s, "조별 활동 — 후보 세 개와 30초 피치 준비 (8분)",
                "각자 떠올리고, 조에서 모으고, 발표할 하나를 고릅니다.")
    step_row(s, TOP, [
        ("1 · 3분", "각자 떠올리기", "한 사람이 한 문장씩 · 오늘 해 본 게임이나 학교에서 불편했던 순간에서"),
        ("2 · 3분", "후보 세 개로", "비슷한 건 합치고 · 한 판 3분 크기로 줄이기"),
        ("3 · 2분", "발표 준비", "발표할 하나를 고르고 · 발표자는 오늘 말을 가장 적게 한 친구"),
    ], h=2.0)
    by2 = TOP + 2.28
    rect(s, ML, by2, CW, 1.0, fill=SOFT, line=LINE, radius=0.1)
    text(s, ML + 0.4, by2 + 0.14, CW - 0.8, 0.72,
         [[("게시판에 올릴 것", {"bold": True, "size": 12, "color": ACCENT_D})],
          [("①  작품 후보 세 개 (한 문장씩)      ②  발표할 후보      ③  조원별 역할 1지망 · 2지망",
            {"bold": True, "size": 13.5, "color": INK})]], line=1.45, anchor=MSO_ANCHOR.MIDDLE)
    footer_rule(s)

    # ---------------------------------------------------------- 23. 피치
    s = S(); header(s)
    title_block(s, "조별 30초 피치 + 동료 피드백",
                "조마다 30초 발표, 다른 조 한 명이 30초 피드백. 1조부터 차례로 갑니다.")
    hw = (CW - 0.3) / 2
    wcard(s, ML, TOP, hw, 2.70, ACCENT, "발표하는 조 · 30초", "한 문장 피치",
          ["한 문장 피치로 시작한다", "무엇이 재밌는지 한 번 더 말한다", "시간이 되면 바로 넘긴다"], strong=True)
    wcard(s, ML + hw + 0.3, TOP, hw, 2.70, BLUE, "듣는 조 한 명 · 30초", "좋았던 점 + 궁금한 점",
          ["좋았던 점 하나", "궁금한 점 하나", "‘별로다’는 말하지 않기"])
    banner(s, [[("궁금한 점", {"bold": True, "color": ACCENT_D}),
                ("이 곧 다음 시간에 우리가 풀 숙제입니다.", {})]], y=5.40, h=0.9)
    footer_rule(s)

    # ---------------------------------------------------------- 24. 핵심
    s = S(); header(s)
    title_block(s, "오늘의 핵심", "만들었고, 비교했고, 기준을 세웠고, 함께할 사람을 찾았습니다.")
    xs, w = cols(3, 0.26)
    wcard(s, xs[0], TOP, w, 2.80, ACCENT, "SPEC", "설계도가 작품을 결정한다",
          ["같은 주제, 같은 AI, 같은 시간", "그런데 화면은 모두 달랐다", "차이는 내가 내린 결정"], strong=True)
    wcard(s, xs[1], TOP, w, 2.80, ACCENT, "QUALITY", "완성에는 이유가 있다",
          ["규칙이 같아도", "시작 · 반응 · 소리 · 끝맺음이", "작품을 바꾼다"], strong=True)
    wcard(s, xs[2], TOP, w, 2.80, ACCENT, "TEAM", "함께 고른 작품을 끝까지",
          ["역할을 나눠 함께 만든다", "우리 반이 세운 기준으로", "7주 뒤 데모데이에 선보인다"], strong=True)
    banner(s, [[("코드를 얼마나 직접 썼느냐가 아니라, ", {}),
                ("AI에게 얼마나 정확하게 말하고 결과를 얼마나 꼼꼼히 확인했느냐", {"bold": True, "color": ACCENT_D}),
                ("가 우리의 성장입니다.", {})]], y=5.40, h=0.9)
    footer_rule(s)

    # ---------------------------------------------------------- 25. NEXT
    s = S(); header(s)
    title_block(s, "NEXT · 4회기 — 재미부터 확인합니다",
                "그림 없이 규칙만으로 정말 재밌는지 먼저 확인합니다. 재미없는 규칙은 그림으로 살아나지 않습니다.")
    step_row(s, TOP, [
        ("01", "작품 하나로", "조별 후보 중 하나를 정한다"),
        ("02", "기획 한 장", "핵심 재미 · 한 판의 흐름 · 만들지 않을 것"),
        ("03", "네모와 동그라미", "그림 없이 규칙만 만들어 재미 확인"),
        ("04", "범위 자르기", "한 판을 끝까지 · 역할 정하기"),
    ], h=1.9)
    banner(s, [[("과제 — ", {"bold": True, "color": ACCENT_D}),
                ("우리 조 후보마다 비슷한 게임 · 앱을 하나씩 30초만 해 보고, ‘이건 이래서 재밌다’ 한 줄 적어 오기", {})]],
           y=4.62, h=0.95)
    text(s, ML, 5.85, CW, 0.5,
         [[("오늘 함께해 주신 선생님들께 감사드립니다.", {"size": 12, "color": FAINT})]], align=PP_ALIGN.CENTER)
    footer_rule(s)

    # ---------------------------------------------------------- 발표자 노트
    slides = list(prs.slides)
    if len(slides) != len(SC.NOTES):
        raise SystemExit("슬라이드 %d장 / 노트 %d개 — 개수가 달라 중단합니다." % (len(slides), len(SC.NOTES)))
    for i, sl in enumerate(slides):
        sl.notes_slide.notes_text_frame.text = SC.note_text(i)

    prs.save(OUT)
    return prs


# ================================================================ 저장 후 자동 검사
def lint(prs):
    """모든 글상자에 대해 글이 상자를 넘치는지, 도형이 슬라이드 밖으로 나가는지 검사한다."""
    EMU = 914400.0
    problems = []
    for si, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            x, y = sh.left / EMU, sh.top / EMU
            w, h = sh.width / EMU, sh.height / EMU
            if x < -0.01 or y < -0.01 or x + w > SW + 0.02 or y + h > SH + 0.02:
                problems.append("슬라이드 %2d  도형이 슬라이드 밖으로 나감 (x=%.2f y=%.2f w=%.2f h=%.2f)" % (si, x, y, w, h))
            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            full = "".join(p.text for p in tf.paragraphs).strip()
            if not full:
                continue
            need = 0.0
            for p in tf.paragraphs:
                sizes = [r.font.size.pt for r in p.runs if r.font.size is not None]
                size = max(sizes) if sizes else 12
                ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
                sa = p.space_after.pt if p.space_after is not None else 0
                t = p.text
                if sizes and not t:
                    need += size * ls / 72.0 + sa / 72.0
                    continue
                font = MONO if any((r.font.name or "") == MONO for r in p.runs) else KR
                unit_ascii = 0.60 if font == MONO else 0.55
                units = sum(1.0 if ord(ch) > 0x2000 else unit_ascii for ch in t)
                per_line = max(1.0, (w * 72.0) / size)
                lines = max(1, int(units / per_line) + (1 if units % per_line > 0.01 else 0))
                need += lines * size * ls * 1.20 / 72.0 + sa / 72.0
            if need > h + 0.06 and need > h * 1.08:
                problems.append("슬라이드 %2d  글 넘침 %.2f in > 상자 %.2f in : %s" % (si, need, h, full[:34]))
    return problems


if __name__ == "__main__":
    prs = build()
    print("생성 완료: %s" % SC.DECK)
    print("슬라이드 %d장 · 발표자 노트 %d개" % (len(prs.slides), len(SC.NOTES)))
    if B.WARN:
        print("\n[카드 넘침 경고] %d건" % len(B.WARN))
        for w_ in B.WARN:
            print("  " + w_)
    else:
        print("카드 넘침 없음")
    probs = lint(prs)
    if probs:
        print("\n[글상자 검사] %d건" % len(probs))
        for p in probs:
            print("  " + p)
    else:
        print("글상자 넘침 · 슬라이드 밖 도형 없음")
