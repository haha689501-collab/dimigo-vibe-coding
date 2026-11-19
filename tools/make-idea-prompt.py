# -*- coding: utf-8 -*-
"""
특강 PPT 생성기 — 「아이디어 찾기 · 프롬프트 쓰기」

실행:  python tools/make-idea-prompt.py
결과:  유튜브/아이디어_찾기와_프롬프트_작성법.pptx

- 참고 영상 2편의 내용을 수업용으로 정리한 덱입니다.
  ① 바이브코딩으로 부업 하고 싶은데 아이디어가 없다면?        (아이디어 재료 늘리기 4가지)
  ② 아직도 AI한테 'XX해줘'라고만 쓰세요? 2026 클로드 공식 프롬프트 작성법
- 대상 학교·학과를 적지 않습니다. 누구에게나 쓸 수 있는 일반 버전입니다.
- 기존 차시 덱과 같은 디자인(머리말 · 제목 · 카드 · 남색 표 · 꼬리말)을 씁니다.
"""
import importlib.util
import os
import sys

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)


def _load(name, filename):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, filename))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


B = _load("l3v2", "make-lesson03-v2.py")
V3 = _load("l3v3", "make-lesson03-v3.py")
L4 = _load("l4v2", "make-lesson04-v2.py")

DECK = os.path.join("06_참고·조사", "유튜브", "아이디어_찾기와_프롬프트_작성법.pptx")

ML, CW, SW, SH, TOP, FOOT_Y = B.ML, B.CW, B.SW, B.SH, B.TOP, B.FOOT_Y
INK, BODY, MUTED, FAINT = B.INK, B.BODY, B.MUTED, B.FAINT
ACCENT, ACCENT_D, ACCENT_L = B.ACCENT, B.ACCENT_D, B.ACCENT_L
LINE, LINE2, CARD, SOFT, TINT = B.LINE, B.LINE2, B.CARD, B.SOFT, B.TINT
DARK, WHITE, BLUE, GREEN, PURPLE = B.DARK, B.WHITE, B.BLUE, B.GREEN, B.PURPLE
KR, MONO = B.KR, B.MONO
rect, text, footer_rule, title_block = B.rect, B.text, B.footer_rule, B.title_block
wcard, dcard, codebox = B.wcard, B.dcard, B.codebox
grid, lcard, minicard, checklist = L4.grid, L4.lcard, L4.minicard, L4.checklist
BLUE_D, GREEN_D, PURPLE_D, RED = L4.BLUE_D, L4.GREEN_D, L4.PURPLE_D, L4.RED

NOTES = {}
PAGES = []
RIGHT = "V I B E   C O D I N G"


def head(s):
    B.header(s, right=RIGHT)


def note(i, lines):
    NOTES[i] = "\n".join(lines)


def page(s, section):
    PAGES.append((s, section))


def table(s, x, y, w, cols, head_, rows, row_h=0.40, head_h=0.40, size=10.5, first_bold=True, accent_col=None):
    rect(s, x, y, w, head_h, fill=DARK)
    cx = x
    for cw, h in zip(cols, head_):
        text(s, cx + 0.16, y + 0.09, cw - 0.24, head_h - 0.12, [[(h, {"bold": True, "size": 10, "color": WHITE})]])
        cx += cw
    yy = y + head_h
    for row in rows:
        rect(s, x, yy, w, row_h, fill=CARD)
        rect(s, x, yy + row_h - 0.01, w, 0.012, fill=LINE)
        cx = x
        for j, (cw, cell) in enumerate(zip(cols, row)):
            bold = first_bold and j == 0
            col = INK if bold else BODY
            if accent_col is not None and j == accent_col:
                col, bold = ACCENT_D, True
            text(s, cx + 0.16, yy + 0.08, cw - 0.24, row_h - 0.10,
                 [[(cell, {"bold": bold, "size": size, "color": col})]], line=1.25)
            cx += cw
        yy += row_h
    return yy


def bar_note(s, runs, y=6.40, h=0.42):
    rect(s, ML, y, CW, h, fill=TINT, radius=0.10)
    text(s, ML + 0.3, y + 0.06, CW - 0.6, h - 0.10, runs, size=11.5, color=BODY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def numcard(s, x, y, w, h, num, h1, h2, bar=ACCENT_L):
    rect(s, x, y, w, h, fill=CARD, line=LINE, radius=0.06)
    rect(s, x, y, w, 0.065, fill=bar)
    text(s, x + 0.24, y + 0.22, 0.8, 0.30, [[(num, {"bold": True, "size": 11, "color": ACCENT})]])
    text(s, x + 0.24, y + 0.54, w - 0.48, 0.34, [[(h1, {"bold": True, "size": 14, "color": INK})]])
    text(s, x + 0.24, y + 0.92, w - 0.48, h - 1.00, [[(h2, {"size": 10.5, "color": MUTED})]], line=1.40)


def part(s, kicker, title, sub, cur, total):
    head(s)
    text(s, ML, 2.55, CW, 0.34, [[(kicker, {"bold": True, "size": 11, "color": ACCENT, "space": 1.6})]])
    text(s, ML, 2.98, CW, 0.85, [[(title, {"bold": True, "size": 38, "color": INK})]])
    rect(s, ML, 3.98, 0.80, 0.045, fill=ACCENT)
    text(s, ML, 4.22, CW * 0.8, 0.4, [[(sub, {"size": 13.5, "color": MUTED})]])
    text(s, SW - ML - 2.0, 4.22, 2.0, 0.4,
         [[("%d / %d" % (cur, total), {"bold": True, "size": 12, "color": FAINT})]], align=PP_ALIGN.RIGHT)
    footer_rule(s)


# ================================================================
def build():
    prs = B.new_deck()
    S = lambda: B.blank(prs)
    n = lambda: B.CUR["n"]

    # ---------------------------------------------------------- 1. 표지
    s = S()
    text(s, ML, 0.44, 6, 0.3, [[("I D E A   ·   P R O M P T", {"bold": True, "size": 11, "color": ACCENT, "space": 1.3})]])
    text(s, ML, 1.86, CW, 1.0, [[("무엇을 만들지 모르겠고,", {"bold": True, "size": 34, "color": MUTED})]])
    text(s, ML, 2.70, CW, 1.1, [[("어떻게 시켜야 할지 모르겠을 때", {"bold": True, "size": 40, "color": INK})]])
    rect(s, ML, 3.84, 0.9, 0.05, fill=ACCENT)
    text(s, ML, 4.12, CW * 0.92, 0.5,
         [[("아이디어 재료를 늘리는 4가지 방법 · 2026년 기준 프롬프트 쓰는 법", {"size": 17, "color": BODY})]])
    w, gap = grid(2, 0.26)
    minicard(s, ML, 5.10, w, 1.26, "PART 1 · 아이디어",
             "아이디어가 없는 게 아니라 재료가 없는 것입니다", bar=ACCENT_L)
    minicard(s, ML + w + gap, 5.10, w, 1.26, "PART 2 · 프롬프트",
             "모델이 바뀌면서 잘 쓰는 법도 바뀌었습니다", bar=ACCENT)
    footer_rule(s)
    text(s, ML, FOOT_Y + 0.12, 10, 0.26,
         [[("참고 영상 2편의 내용을 정리했습니다", {"size": 9.5, "color": FAINT, "space": 0.8})]])
    note(n(), [
        "[이 덱의 용도]",
        "무엇을 만들지 정하지 못한 사람, 프롬프트를 어떻게 써야 할지 막막한 사람 모두에게 쓸 수 있습니다.",
        "",
        "[대본]",
        "\"만들고 싶은 건 있는데 뭘 만들지 모르겠다, 또는 시켜 봤는데 원하는 게 안 나온다.",
        "둘 다 흔한 일이고, 둘 다 방법이 있습니다. 오늘 두 가지를 같이 보겠습니다.\"",
    ])

    # ---------------------------------------------------------- 2. 오늘 두 가지
    s = S()
    head(s)
    title_block(s, "오늘 두 가지", "앞에서 재료를 모으고, 뒤에서 시키는 법을 바꿉니다")
    w, gap = grid(2, 0.26)
    wcard(s, ML, TOP, w, 3.00, ACCENT, "PART 1", "아이디어 재료 늘리기",
          ["남이 이미 만든 것 보기", "지금 뜨는 것 올라타기",
           "사람들이 손으로 하는 일 찾기", "불만 리뷰에서 기회 찾기"], size=11.5)
    wcard(s, ML + w + gap, TOP, w, 3.00, GREEN, "PART 2", "2026년 기준 프롬프트",
          ["예전 요령 중 지금은 필요 없는 것", "역할 · 근거 · 형식을 지정하기",
           "하지 말 것 대신 할 것 쓰기", "막혔을 때 꺼내는 한 줄"],
          label_color=GREEN_D, size=11.5)
    bar_note(s, [[("두 가지는 이어져 있습니다  —  ", {}),
                  ("재료가 있어야 요청할 내용이 생기고, 요청이 정확해야 결과가 나옵니다", {"bold": True, "color": ACCENT_D})]],
             y=5.46, h=0.56)
    footer_rule(s); page(s, "도입")
    note(n(), [
        "[시간] 2분",
        "[대본]",
        "\"앞부분은 재료 모으는 법, 뒷부분은 그 재료로 시키는 법입니다.",
        "순서가 중요합니다. 재료가 없으면 아무리 프롬프트를 잘 써도 쓸 내용이 없습니다.\"",
    ])

    # ---------------------------------------------------------- 3. PART 1
    s = S()
    part(s, "P A R T   1", "아이디어 찾기", "재료를 왕창 늘리는 네 가지 방법", 1, 2)
    note(n(), ["[전환]", "\"먼저 아이디어입니다.\""])

    # ---------------------------------------------------------- 4. 냉장고
    s = S()
    head(s)
    title_block(s, "아이디어가 없는 게 아닙니다", "재료가 두세 개밖에 없는 것입니다")
    cw_ = 6.40
    lcard(s, ML, TOP, cw_, 2.30, ACCENT, "흔한 출발점 · 내 불편에서 시작하기",
          ["쓸 수 있는 재료가 두세 개뿐입니다", "그래서 금방 막힙니다"], size=12)
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    lcard(s, rx, TOP, rw, 2.30, GREEN, "같은 상황 · 냉장고",
          ["재료가 두세 개면 만들 수 있는 요리가 없습니다",
           "요리 실력 문제가 아닙니다"], label_color=GREEN_D, size=12)
    bar_note(s, [[("아이디어가 안 떠오르는 건 창의력이 없어서가 아니라  ", {}),
                  ("재료가 부족해서입니다", {"bold": True, "color": ACCENT_D})]], y=5.10, h=0.70)
    footer_rule(s); page(s, "아이디어")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"대부분 '내 주변 불편함'에서 시작하죠. 좋은 방법이지만 문제가 하나 있습니다. 재료가 두세 개밖에 없다는 겁니다.",
        "냉장고에 재료가 두세 개 있으면 만들 수 있는 요리가 몇 개 없죠. 요리를 못해서가 아닙니다.",
        "그래서 오늘은 재료를 왕창 늘리는 방법을 봅니다.\"",
    ])

    # ---------------------------------------------------------- 5. 4가지
    s = S()
    head(s)
    title_block(s, "재료를 늘리는 네 가지", "모두 오늘 당장 혼자 할 수 있는 것들입니다")
    w, gap = grid(2, 0.26)
    items = [("01", "이미 나온 것 보기", "매일 새로 나오는 서비스를 훑습니다"),
             ("02", "지금 뜨는 것 보기", "커뮤니티에서 뜨는 주제에 올라탑니다"),
             ("03", "손으로 하는 일 찾기", "사람들이 돈 주고 자동화하려는 일을 봅니다"),
             ("04", "불만 리뷰 읽기", "별점 낮은 리뷰에 기회가 모여 있습니다")]
    for i, (num, h1, h2) in enumerate(items):
        x = ML + (i % 2) * (w + gap)
        y = TOP + (i // 2) * 1.62
        numcard(s, x, y, w, 1.46, num, h1, h2, bar=ACCENT if i in (0, 3) else ACCENT_L)
    bar_note(s, [[("네 개를 한꺼번에 하려 하면 하나도 안 합니다  —  ", {}),
                  ("가장 마음에 드는 하나만 고르세요", {"bold": True, "color": ACCENT_D})]], y=5.80, h=0.56)
    footer_rule(s); page(s, "아이디어")
    note(n(), [
        "[시간] 1분",
        "[대본]",
        "\"네 가지입니다. 이론이 아니라 지금 당장 혼자 할 수 있는 것들입니다.",
        "다만 한 번에 다 하려고 하면 결국 안 하게 됩니다. 하나만 고르세요.\"",
    ])

    # ---------------------------------------------------------- 6. 방법 1
    s = S()
    head(s)
    title_block(s, "① 이미 나온 것 보기", "세상에 나와 있는 서비스를 많이 보는 것만큼 빠른 방법이 없습니다")
    cols = [3.30, CW - 3.30 - 5.40, 5.40]
    table(s, ML, TOP, CW, cols, ["어디서", "무엇을 보나", "이럴 때 좋다"],
          [["프로덕트 헌트", "매일 새로 올라오는 해외 서비스", "요즘 어떤 걸 만드는지 감 잡기"],
           ["디스콰이엇 · 국내판", "국내 소규모 서비스", "작게 만든 것들의 범위 보기"],
           ["혁신의 숲", "국내 스타트업 · 매출 · 투자", "실제로 돈이 되는 분야 보기"],
           ["앱 스토어 인기 차트", "사람들이 실제로 쓰는 앱", "익숙한 영역에서 빈틈 찾기"],
           ["해커톤 갤러리", "또래가 만든 완성작", "6시간에 가능한 크기 가늠하기"]],
          row_h=0.52, size=11, accent_col=2)
    bar_note(s, [[("하루 ", {}), ("세 개만", {"bold": True, "color": ACCENT_D}),
                  (" 봐도 눈이 뜨입니다. 많이 보는 게 아니라 꾸준히 보는 게 핵심입니다", {})]], y=5.96, h=0.52)
    footer_rule(s); page(s, "아이디어")
    note(n(), [
        "[시간] 3분",
        "",
        "[대본]",
        "\"첫 번째는 레퍼런스입니다. 이미 나와 있는 걸 많이 보는 겁니다.",
        "프로덕트 헌트에는 매일 새 서비스가 올라옵니다. 하루에 세 개만 봐도 눈이 뜨입니다.",
        "국내 서비스를 보고 싶으면 디스콰이엇, 돈이 어디서 도는지 보고 싶으면 혁신의 숲을 보세요.",
        "또래가 뭘 만들었는지 보고 싶으면 해커톤 갤러리가 가장 현실적입니다. 6시간에 어디까지 되는지 바로 보입니다.\"",
    ])

    # ---------------------------------------------------------- 7. 자동 수집
    s = S()
    head(s)
    title_block(s, "매번 찾기 귀찮다면", "매일 아침 보고받게 만들어 두면 됩니다")
    cw_ = 7.40
    codebox(s, ML, TOP, cw_, 2.30, "PROMPT · 예약 작업",
            ["매일 오전 10시에 새로 올라온 서비스 TOP 5를 모아서",
             "한 장으로 정리해 줘.",
             "각각 무슨 서비스인지, 사람들이 뭐라고 반응했는지 한 줄씩."])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    lcard(s, rx, TOP, rw, 2.30, BLUE, "어디서 설정하나",
          ["클로드 데스크톱 → 예정된 작업", "새 작업 → 프롬프트 붙여넣기 → 실행"],
          label_color=BLUE_D, size=11)
    w, gap = grid(3, 0.24)
    for i, (h1, h2) in enumerate([("찾는 시간 0분", "아침에 결과만 읽습니다"),
                                  ("매일 쌓인다", "일주일이면 재료 35개"),
                                  ("주제를 좁혀도 된다", "\"1인 개발자용만\" 처럼")]):
        minicard(s, ML + i * (w + gap), TOP + 2.54, w, 1.24, h1, h2, bar=ACCENT_L)
    footer_rule(s); page(s, "아이디어")
    note(n(), [
        "[시간] 3분 (시연)",
        "",
        "[대본]",
        "\"매번 사이트를 찾아 들어가는 건 귀찮습니다. 그래서 예약 작업으로 걸어 둡니다.",
        "클로드 데스크톱의 예정된 작업에서 새 작업을 만들고, 이 프롬프트를 넣으면 매일 아침 정리된 결과를 받습니다.",
        "하루 5개면 일주일에 35개입니다. 재료가 없다는 말이 안 나옵니다.\"",
        "",
        "[운영]",
        "예약 작업을 쓰기 어려우면, 같은 프롬프트를 매일 아침 직접 한 번 보내는 것으로도 충분합니다.",
    ])

    # ---------------------------------------------------------- 8. 방법 2
    s = S()
    head(s)
    title_block(s, "② 지금 뜨는 것 보기", "이미 커진 파도에 올라타는 방법입니다")
    cw_ = 6.10
    lcard(s, ML, TOP, cw_, 2.40, ACCENT, "어디서 보나",
          ["구글 트렌드 · 네이버 데이터랩", "유튜브 인기 급상승 · 커뮤니티 인기글",
           "트렌드를 모아 주는 도구(MCP) 연결"], size=11.5)
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    codebox(s, rx, TOP, rw, 2.40, "PROMPT",
            ["지금 커뮤니티에서 뜨고 있는 주제 TOP 10이 뭐야?",
             "혼자서 작게 만들 수 있는 것 위주로 정리해 줘.",
             "",
             "→ 하나 고른 뒤: 그 안에서 제일 빨리 크는",
             "   세부 키워드 5개 뽑아 줘."])
    bar_note(s, [[("넓게 한 번, 좁게 한 번  —  ", {}),
                  ("주제를 고른 뒤 그 안에서 다시 좁히면 재료가 구체적으로 바뀝니다", {"bold": True, "color": ACCENT_D})]],
             y=5.42, h=0.56)
    footer_rule(s); page(s, "아이디어")
    note(n(), [
        "[시간] 2분",
        "[대본]",
        "\"두 번째는 트렌드입니다. 지금 뜨는 주제를 물어보고, 그중 하나를 골라 다시 좁혀 들어갑니다.",
        "넓게 한 번, 좁게 한 번. 이 두 번이면 키워드가 손에 잡힙니다.\"",
    ])

    # ---------------------------------------------------------- 9. 방법 3
    s = S()
    head(s)
    title_block(s, "③ 손으로 하는 일 찾기", "사람들이 돈을 주고라도 자동화하려는 일을 봅니다")
    w, gap = grid(3, 0.24)
    for i, (h1, h2) in enumerate([("외주 플랫폼에서", "\"자동화\"로 검색하면 의뢰 목록이 나옵니다"),
                                  ("주변에서", "엑셀로 매번 손으로 정리하는 사람을 찾습니다"),
                                  ("학교·동아리에서", "출석 · 집계 · 공지 돌리기처럼 반복되는 일")]):
        minicard(s, ML + i * (w + gap), TOP, w, 1.40, h1, h2, bar=ACCENT if i == 1 else ACCENT_L)
    lcard(s, ML, TOP + 1.64, CW, 1.86, GREEN, "실제로 올라오는 의뢰 예",
          ["쇼핑몰 자동 주문 · 티켓 예약 · 영상 제작 반복 작업 자동화",
           "무엇이 실제로 돈이 되는지가 같이 보입니다"], label_color=GREEN_D, size=11.5)
    bar_note(s, [[("\"이걸 누가 손으로 해?\" 싶은 일이  ", {}),
                  ("대부분 아직 손으로 처리되고 있습니다", {"bold": True, "color": ACCENT_D})]], y=5.74, h=0.56)
    footer_rule(s); page(s, "아이디어")
    note(n(), [
        "[시간] 2분",
        "[대본]",
        "\"세 번째는 사람들이 손으로 하는 일을 찾는 겁니다. 외주 플랫폼에서 '자동화'로 검색하면 의뢰가 쭉 나옵니다.",
        "재료가 쌓이는 동시에, 어떤 게 실제로 돈이 되는지도 같이 보입니다.",
        "꼭 플랫폼이 아니어도 됩니다. 주변에서 매번 손으로 정리하는 사람을 찾으면 그게 재료입니다.\"",
    ])

    # ---------------------------------------------------------- 10. 방법 4
    s = S()
    head(s)
    title_block(s, "④ 불만 리뷰 읽기", "별점이 가장 낮은 리뷰부터 정렬합니다")
    cw_ = 6.10
    lcard(s, ML, TOP, cw_, 2.40, RED, "읽는 법",
          ["유명한 서비스를 하나 고릅니다", "리뷰를 별점 낮은 순으로 정렬합니다",
           "반복되는 단어를 셉니다"], label_color=RED, size=11.5)
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    lcard(s, rx, TOP, rw, 2.40, ACCENT, "실제로 반복되는 말",
          ["\"너무 복잡하다\"  \"직관적이지 않다\"",
           "\"배울 게 많다\"  \"비싸다\"  \"느리다\""], size=11.5)
    bar_note(s, [[("공통 불만 하나를 해결한 더 쉬운 버전  —  ", {}),
                  ("그게 그대로 아이디어가 됩니다", {"bold": True, "color": ACCENT_D})]], y=5.42, h=0.56)
    footer_rule(s); page(s, "아이디어")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"네 번째는 불만 리뷰입니다. 유명한 서비스를 하나 골라 별점 낮은 순으로 정렬해 보세요.",
        "거기 반복되는 말이 있습니다. 복잡하다, 어렵다, 직관적이지 않다, 비싸다.",
        "이 공통 불만을 해결한 더 쉬운 버전이 곧 아이디어입니다. 완전히 새로운 걸 만들 필요가 없습니다.\"",
    ])

    # ---------------------------------------------------------- 11. 실습
    s = S()
    head(s)
    title_block(s, "10분 실습 · 재료 모으기", "지금 자리에서 재료 10개를 만듭니다")
    checklist(s, ML, TOP, CW,
              [("3분 · 불편 5개", "오늘 하루 짜증 났던 일을 그냥 적습니다"),
               ("3분 · 레퍼런스 3개", "비슷한 걸 이미 만든 서비스를 찾습니다"),
               ("2분 · 불만 리뷰 2개", "그 서비스의 낮은 별점 리뷰에서 한 줄씩"),
               ("2분 · 한 줄로 합치기", "누구의 · 어떤 불편을 · 어떻게 줄인다")],
              row_h=0.76, gap=0.12)
    bar_note(s, [[("고르는 기준  ", {"bold": True, "color": ACCENT_D}),
                  ("내가 이번 주에 쓸 것  ·  친구에게 보여 줄 수 있는 것  ·  6시간에 끝날 크기", {"bold": True})]],
             y=6.00, h=0.52)
    footer_rule(s); page(s, "아이디어")
    note(n(), [
        "[시간] 10분",
        "",
        "[운영]",
        "타이머를 띄우고 단계마다 끊어 줍니다. 완성도보다 개수가 목적입니다.",
        "마지막 한 줄은 \"누구의 어떤 불편을 어떻게 줄인다\" 형식으로만 쓰게 합니다.",
        "",
        "[대본]",
        "\"지금 10분 동안 재료를 만듭니다. 잘 쓰려고 하지 마세요. 개수가 목적입니다.\"",
    ])

    # ---------------------------------------------------------- 12. PART 2
    s = S()
    part(s, "P A R T   2", "프롬프트 쓰기", "모델이 바뀌면서 잘 쓰는 법도 바뀌었습니다", 2, 2)
    note(n(), ["[전환]", "\"재료를 찾았으면 이제 제대로 시킬 차례입니다.\""])

    # ---------------------------------------------------------- 13. 왜 다시
    s = S()
    head(s)
    title_block(s, "\"이제 대충 말해도 알아듣는다\"", "맞는 말이지만, 그래서 쓰는 법이 바뀌었습니다")
    w, gap = grid(2, 0.26)
    wcard(s, ML, TOP, w, 2.46, BLUE, "달라진 점", "모델이 스스로 하는 일이 늘었다",
          ["답하기 전에 생각하는 단계가 기본", "모르는 건 검색해서 찾아봄"], label_color=BLUE_D, size=11.5)
    wcard(s, ML + w + gap, TOP, w, 2.46, ACCENT, "그대로인 점", "원하는 걸 전하는 수단은 글뿐",
          ["무엇을 원하는지는 모델이 모름", "그래서 가이드가 새로 나옴"], size=11.5)
    bar_note(s, [[("예전에 배운 요령 중에는  ", {}),
                  ("지금은 시간만 더 쓰게 하는 것도 있습니다", {"bold": True, "color": ACCENT_D})]], y=5.42, h=0.56)
    footer_rule(s); page(s, "프롬프트")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"요즘은 대충 말해도 알아듣는다는 말, 많이 들으셨을 겁니다. 반은 맞습니다.",
        "모델이 답하기 전에 스스로 생각하고, 모르면 검색도 합니다.",
        "그런데도 모델을 만든 쪽에서 새 프롬프트 가이드를 따로 냅니다. 모델이 바뀌면 쓰는 법도 바뀌기 때문입니다.",
        "예전에 배운 요령 중에는 지금은 오히려 시간만 더 쓰게 하는 것도 있습니다.\"",
    ])

    # ---------------------------------------------------------- 14. 옛 요령 3개
    s = S()
    head(s)
    title_block(s, "예전 요령, 지금은 이렇게", "버리는 게 아니라 쓰는 자리를 바꿉니다")
    cols = [4.00, 3.60, CW - 4.00 - 3.60]
    table(s, ML, TOP, CW, cols, ["예전에 붙이던 말", "지금은", "대신 이렇게"],
          [["\"단계별로 생각해\"", "이미 생각하고 답함", "노력(effort) 설정을 조절한다"],
           ["\"너는 전문가야\"", "정답률은 그대로", "어떤 관점으로 볼지를 지정한다"],
           ["\"추측하지 마\"", "막연한 금지", "어디서 확인할지를 알려 준다"],
           ["\"예시를 붙여라\"", "지금도 유효", "말투와 형식을 맞추는 가장 확실한 방법"]],
          row_h=0.62, size=11, accent_col=2)
    bar_note(s, [[("공통점  ", {"bold": True, "color": ACCENT_D}),
                  ("하지 말 것을 말하는 대신, 할 일을 주는 쪽으로 바뀌었습니다", {"bold": True})]], y=5.84, h=0.52)
    footer_rule(s); page(s, "프롬프트")
    note(n(), [
        "[시간] 3분",
        "",
        "[대본]",
        "\"많이 들어 본 요령 네 가지입니다. 하나씩 보겠습니다.",
        "'단계별로 생각해'는 이미 생각하는 모델에게 생각하라고 한 번 더 붙이는 셈입니다. 책 펴고 앉았는데 옆에서 공부하라고 하는 거죠.",
        "'너는 전문가야'는 정답률을 올리진 않습니다. 대신 어떤 시각으로 봐 달라는 뜻으로 쓰면 좋습니다.",
        "'추측하지 마'는 막연합니다. 어디서 확인하라고 알려 주는 게 낫습니다.",
        "예시 붙이기는 지금도 가장 확실한 방법입니다.\"",
    ])

    # ---------------------------------------------------------- 15. effort
    s = S()
    head(s)
    title_block(s, "생각을 더 시키고 싶다면", "문장에 쓰지 말고 설정에서 조절합니다")
    cw_ = 5.00
    lcard(s, ML, TOP, cw_, 1.70, ACCENT, "어디 있나",
          ["보내기 버튼 옆 모델 이름을 누릅니다"], size=11.5)
    lcard(s, ML, TOP + 1.86, cw_, 1.70, ACCENT_L, "처음엔",
          ["기본값(중간)에서 시작합니다"], label_color=ACCENT_D, size=11.5)
    tx = ML + cw_ + 0.26
    tw = CW - cw_ - 0.26
    cols = [1.70, tw - 1.70]
    table(s, tx, TOP, tw, cols, ["언제", "어떻게"],
          [["낮춘다", "짧은 질문 · 문장 다듬기처럼 빨리 끝낼 일"],
           ["기본(중간)", "대부분의 작업 · 처음 시작할 때"],
           ["높인다", "여러 자료를 비교하거나 복잡한 계획을 세울 때"],
           ["주의", "높일수록 더 오래 기다리고 사용량도 빨리 씁니다"]],
          row_h=0.68, size=11, accent_col=1)
    bar_note(s, [[("매번 최대로 올려 둘 필요는 없습니다  —  ", {}),
                  ("기본값으로 두고, 필요할 때만 올리세요", {"bold": True, "color": ACCENT_D})]], y=5.68, h=0.56)
    footer_rule(s); page(s, "프롬프트")
    note(n(), [
        "[시간] 2분",
        "[대본]",
        "\"AI가 더 깊게 생각하길 바란다면 문장에 쓰지 말고 설정을 바꾸세요. 모델 이름 옆에 노력 설정이 있습니다.",
        "기본값으로 시작하고, 간단한 일은 낮추고, 복잡한 일은 올립니다. 다만 올릴수록 느려지고 사용량을 빨리 씁니다.\"",
    ])

    # ---------------------------------------------------------- 16. 관점 · 근거
    s = S()
    head(s)
    title_block(s, "역할 대신 관점, 금지 대신 근거", "같은 목적을 더 정확하게 전달하는 방법")
    cols = [1.70, 4.40, CW - 1.70 - 4.40]
    table(s, ML, TOP, CW, cols, ["상황", "이렇게 말고", "이렇게"],
          [["역할", "\"너는 세계 최고의 전문가야\"", "\"처음 읽는 사람 입장에서 막히는 곳을 찾아 줘\""],
           ["사실 확인", "\"절대 지어내지 마\"", "\"공식 문서를 확인하고 근거 링크와 함께 짧게\""],
           ["디자인", "\"흔한 AI 느낌은 피해 줘\"", "\"크림색 배경과 알약 버튼은 쓰지 마\""],
           ["말투", "\"자연스럽게 써 줘\"", "\"이 글의 말투와 길이만 참고해 줘\""]],
          row_h=0.66, size=11, accent_col=2)
    bar_note(s, [[("왼쪽은 ", {}), ("막연한 금지", {"bold": True, "color": ACCENT_D}),
                  (", 오른쪽은 ", {}), ("할 일을 주는 지시", {"bold": True, "color": ACCENT_D}),
                  ("입니다", {})]], y=5.60, h=0.56)
    footer_rule(s); page(s, "프롬프트")
    note(n(), [
        "[시간] 3분",
        "",
        "[대본]",
        "\"역할을 붙인다고 정답률이 오르진 않습니다. 하지만 어떤 입장에서 봐 달라는 뜻이라면 효과가 있습니다.",
        "'전문가야'라고 하고 넘기기보다, 내가 원하는 관점과 할 일을 같이 적으세요.",
        "'지어내지 마'도 마찬가지입니다. 하지 말라는 말 대신, 어디서 찾아 무엇을 근거로 보여 달라고 하세요.",
        "요즘 모델은 검색을 직접 할 수 있으니 확인할 곳을 정해 주면 됩니다.\"",
    ])

    # ---------------------------------------------------------- 17. 구체적 금지
    s = S()
    head(s)
    title_block(s, "피할 것은 구체적으로", "막연히 피하라고 하면 기본값이 다른 기본값으로 바뀔 뿐입니다")
    cw_ = 6.10
    lcard(s, ML, TOP, cw_, 1.80, RED, "막연한 요청",
          ["\"흔한 AI 느낌은 피해 줘\"", "\"좀 더 멋있게\""], label_color=RED, size=12)
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    lcard(s, rx, TOP, rw, 1.80, GREEN, "구체적인 요청",
          ["\"크림색 배경은 쓰지 마\"", "\"알약 모양 버튼은 쓰지 마\""], label_color=GREEN_D, size=12)
    lcard(s, ML, TOP + 2.04, CW, 1.76, ACCENT, "그리고 할 것도 같이",
          ["피할 것 하나를 적었으면, 원하는 것 하나도 같이 적습니다",
           "예: \"각진 버튼, 굵은 제목, 배경은 흰색\""], size=11.5)
    bar_note(s, [[("결과가 매번 다르게 나온다면  ", {}),
                  ("요청이 막연한 것이지 모델이 이상한 게 아닙니다", {"bold": True, "color": ACCENT_D})]], y=5.86, h=0.52)
    footer_rule(s); page(s, "프롬프트")
    note(n(), [
        "[시간] 2분",
        "[대본]",
        "\"'흔한 AI 느낌 피해 줘'라고 하면 기본 스타일 하나가 다른 기본 스타일로 바뀔 뿐입니다.",
        "무엇을 피해야 하는지 구체적으로 적으세요. 색, 모양, 길이처럼 눈에 보이는 것으로요.\"",
    ])

    # ---------------------------------------------------------- 18. 예시
    s = S()
    head(s)
    title_block(s, "예시는 지금도 가장 확실합니다", "말투와 형식을 맞추는 가장 믿을 만한 방법")
    cw_ = 7.40
    codebox(s, ML, TOP, cw_, 2.40, "PROMPT",
            ["아래 글의 말투와 길이를 참고해서 소개 문구를 써 줘.",
             "내용은 가져오지 말고 말투와 길이만 참고해 줘.",
             "",
             "여기부터 예시 —",
             "[마음에 드는 글 한두 개 붙여넣기]"])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    lcard(s, rx, TOP, rw, 2.40, GREEN, "왜 되나",
          ["\"자연스럽게\"는 사람마다 다릅니다", "예시는 그 기준을 눈으로 보여 줍니다",
           "\"여기부터 예시\"로 경계를 표시합니다"], label_color=GREEN_D, size=11)
    bar_note(s, [[("설명하기 어려운 건 설명하지 말고  ", {}),
                  ("보여 주세요", {"bold": True, "color": ACCENT_D})]], y=5.42, h=0.56)
    footer_rule(s); page(s, "프롬프트")
    note(n(), [
        "[시간] 2분",
        "[대본]",
        "\"'말하듯 써 줘'라고 해도 사람마다 떠올리는 말투가 다릅니다. 그럴 땐 예시를 보여 주는 게 가장 빠릅니다.",
        "대신 어디까지가 내 부탁이고 어디부터가 참고할 글인지 꼭 표시해 주세요.\"",
    ])

    # ---------------------------------------------------------- 19. 인터뷰
    s = S()
    head(s)
    title_block(s, "뭘 적어야 할지 모르겠을 때", "거꾸로, 나에게 물어봐 달라고 합니다")
    cw_ = 7.40
    codebox(s, ML, TOP, cw_, 2.46, "PROMPT",
            ["내가 뭘 정해야 하는지 아직 모르겠어.",
             "필요한 걸 하나씩 물어봐 줘.",
             "",
             "(질문에 답한 뒤)",
             "내가 답한 내용을 기반으로 프롬프트를 적어 줘."])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    lcard(s, rx, TOP, rw, 2.46, BLUE, "왜 이게 통하나",
          ["내가 뭘 모르는지 모를 때가 많습니다", "질문에 답하다 보면 방향이 보입니다",
           "답한 내용이 그대로 프롬프트가 됩니다"], label_color=BLUE_D, size=11)
    bar_note(s, [[("막혔을 때 꺼내는 한 줄  —  ", {}),
                  ("\"나를 인터뷰해 줘\"", {"bold": True, "color": ACCENT_D})]], y=5.48, h=0.56)
    footer_rule(s); page(s, "프롬프트")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"여기까지 듣고도 막상 뭘 적어야 할지 모를 때가 있습니다. 그럴 땐 거꾸로 하세요.",
        "'내가 뭘 정해야 할지 모르겠으니 하나씩 물어봐 줘.' 아는 것부터 답하다 보면 방향이 잡힙니다.",
        "끝나면 '답한 내용으로 프롬프트를 적어 줘'라고 하면 됩니다. 이건 공식 가이드에도 나오는 방법입니다.\"",
    ])

    # ---------------------------------------------------------- 20. 템플릿
    s = S()
    head(s)
    title_block(s, "한 장으로 정리한 프롬프트", "빈칸만 채우면 됩니다")
    codebox(s, ML, TOP, CW, 3.30, "TEMPLATE",
            ["[할 일]    무엇을 만들거나 고칠 건지 한 문장",
             "[맥락]    어디에 쓰는지 · 누가 보는지",
             "[관점]    어떤 입장에서 봐 줬으면 하는지",
             "[확인]    어디서 찾아 근거를 보여 줄지",
             "[형식]    어떤 모양으로 받고 싶은지 (길이 · 구조)",
             "[피할 것]  구체적으로 (색 · 모양 · 말투)",
             "[예시]    여기부터 예시 — 말투와 길이만 참고"])
    w, gap = grid(3, 0.24)
    for i, (h1, h2) in enumerate([("어디에 쓸 건지", "용도가 정해지면 형식이 따라옵니다"),
                                  ("뭘 보고 확인할지", "근거가 있으면 믿고 쓸 수 있습니다"),
                                  ("어떤 모양으로 받을지", "길이와 구조를 먼저 정합니다")]):
        minicard(s, ML + i * (w + gap), TOP + 3.54, w, 1.20, h1, h2, bar=ACCENT_L)
    footer_rule(s); page(s, "프롬프트")
    note(n(), [
        "[시간] 2분",
        "[대본]",
        "\"지금까지 본 걸 한 장으로 합치면 이렇게 됩니다. 매번 다 쓸 필요는 없습니다.",
        "다만 '잘해 줘'로 끝내지는 마세요. 어디에 쓸 건지, 뭘 보고 확인할지, 어떤 모양으로 받고 싶은지. 이 세 가지만 챙겨도 결과가 달라집니다.\"",
    ])

    # ---------------------------------------------------------- 21. 마무리
    s = S()
    head(s)
    title_block(s, "오늘 가져갈 것", "두 가지만 기억하면 됩니다")
    w, gap = grid(2, 0.26)
    wcard(s, ML, TOP, w, 2.40, ACCENT, "아이디어", "재료가 부족한 것뿐입니다",
          ["네 가지 중 하나만 골라 오늘 시작", "하루 세 개씩만 봐도 쌓입니다"], size=11.5)
    wcard(s, ML + w + gap, TOP, w, 2.40, GREEN, "프롬프트", "하지 말 것 대신 할 일을 줍니다",
          ["관점 · 근거 · 형식을 적는다", "막히면 \"나를 인터뷰해 줘\""],
          label_color=GREEN_D, size=11.5)
    rect(s, ML, 5.10, CW, 1.10, fill=DARK, radius=0.06)
    text(s, ML + 0.40, 5.26, CW - 0.8, 0.28,
         [[("오늘 할 일", {"bold": True, "size": 10, "color": ACCENT_L, "space": 1.0})]])
    text(s, ML + 0.40, 5.56, CW - 0.8, 0.44,
         [[("재료 수집 방법 하나 고르기  ·  프롬프트 한 줄을 오늘 바꿔 보기", {"bold": True, "size": 19, "color": WHITE})]])
    footer_rule(s); page(s, "마무리")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"정리하겠습니다. 아이디어가 없는 게 아니라 재료가 없는 겁니다. 네 가지 중 하나만 오늘 시작하세요.",
        "프롬프트는 하지 말 것을 말하는 대신 할 일을 주는 쪽으로 바뀌었습니다. 관점, 근거, 형식.",
        "그리고 막히면 언제든 '나를 인터뷰해 줘'로 시작하면 됩니다.\"",
    ])

    # ---------------------------------------------------------- 꼬리말 · 노트
    total = len(prs.slides)
    idx = {id(sl): k for k, sl in enumerate(prs.slides, 1)}
    for sl, sec in PAGES:
        k = idx[id(sl)]
        text(sl, SW - ML - 3.4, FOOT_Y + 0.12, 3.4, 0.26,
             [[("%s  %02d / %02d" % (sec, k, total), {"size": 9, "color": FAINT, "space": 0.8})]],
             align=PP_ALIGN.RIGHT)
    for k, sl in enumerate(prs.slides, start=1):
        if k in NOTES:
            sl.notes_slide.notes_text_frame.text = NOTES[k]
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, DECK)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    prs.save(out)
    return prs, out


if __name__ == "__main__":
    B.CUR["n"] = 0
    del B.WARN[:]
    prs, out = build()
    print("[특강] %s" % os.path.basename(out))
    print("  슬라이드 %d장 · 발표자 노트 %d개" % (len(prs.slides), len(NOTES)))
    if B.WARN:
        print("  [카드 넘침] %d건" % len(B.WARN))
        for w_ in B.WARN:
            print("    " + w_)
    probs = V3.lint(prs)
    if probs:
        print("  [글상자 검사] %d건" % len(probs))
        for pr in probs:
            print("    " + pr)
    if not B.WARN and not probs:
        print("  넘침 없음")
