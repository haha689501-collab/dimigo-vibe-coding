# -*- coding: utf-8 -*-
"""
4차시 v2 PPT 생성기 — 「내 아이디어를 바로 만든다」

실행:  python tools/make-lesson04-v2.py
결과:  디미고_바이브코딩_4차시_v2_바로만들기.pptx  (프로젝트 최상위, 새 파일)

- 교안: MD/디미고_바이브코딩_4차시_교안_v2_바로만들기.md
- 설명을 먼저 하지 않는다. 보여 주고 → 각자 만들게 하고 → 막혔을 때 가르친다.
- 슬라이드는 설명 자료가 아니라 실습 중에 띄워 두는 '미션판'이다. 11장.
- 색·여백·카드 부품은 1~4차시 덱 생성기의 것을 그대로 쓴다.
- 기존 PPT는 열지도 수정하지도 않는다. 새 파일만 만든다.
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

DECK = "디미고_바이브코딩_4차시_v2_바로만들기.pptx"

ML, CW, SW, SH, TOP, FOOT_Y = B.ML, B.CW, B.SW, B.SH, B.TOP, B.FOOT_Y
INK, BODY, MUTED, FAINT = B.INK, B.BODY, B.MUTED, B.FAINT
ACCENT, ACCENT_D, ACCENT_L = B.ACCENT, B.ACCENT_D, B.ACCENT_L
LINE, LINE2, CARD, SOFT, TINT = B.LINE, B.LINE2, B.CARD, B.SOFT, B.TINT
DARK, WHITE, BLUE, GREEN, PURPLE = B.DARK, B.WHITE, B.BLUE, B.GREEN, B.PURPLE
KR, MONO = B.KR, B.MONO
rect, text, header, footer_rule, title_block, banner = B.rect, B.text, B.header, B.footer_rule, B.title_block, B.banner
wcard, dcard, codebox, step_row = B.wcard, B.dcard, B.codebox, B.step_row

BLUE_D = RGBColor(0x3B, 0x53, 0x9B)
GREEN_D = RGBColor(0x2F, 0x7A, 0x55)
PURPLE_D = RGBColor(0x6B, 0x4A, 0x96)
RED = RGBColor(0xC2, 0x41, 0x41)
RED_SOFT = RGBColor(0xFB, 0xEC, 0xEC)
GREEN_SOFT = RGBColor(0xE8, 0xF5, 0xEE)
TOTAL_PARTS = 3


def grid(n, gap):
    _xs, w = V3.cols(n, gap)
    return w, gap


def lcard(s, x, y, w, h, topbar, label, lines, note=None, label_color=None, size=11.5, strong=False):
    """제목 줄 없이 라벨만 두는 카드"""
    wcard(s, x, y, w, h, topbar, label, None, lines, note=note,
          label_color=label_color, size=size, strong=strong)


def minicard(s, x, y, w, h, head, desc, bar=None):
    rect(s, x, y, w, h, fill=CARD, line=LINE, radius=0.06)
    if bar:
        rect(s, x, y, w, 0.065, fill=bar)
    text(s, x + 0.22, y + 0.20, w - 0.44, 0.34, [[(head, {"bold": True, "size": 13.5, "color": INK})]])
    text(s, x + 0.22, y + 0.56, w - 0.44, h - 0.62, [[(desc, {"size": 10.5, "color": MUTED})]], line=1.35)


def checklist(s, x, y, w, items, row_h=0.70, gap=0.10):
    """☐ 번호 · 제목 · 설명 한 줄씩 — 미션판"""
    for i, (head, desc) in enumerate(items):
        yy = y + i * (row_h + gap)
        rect(s, x, yy, w, row_h, fill=CARD, line=LINE, radius=0.06)
        rect(s, x + 0.22, yy + row_h / 2 - 0.15, 0.30, 0.30, fill=WHITE, line=ACCENT, lw=1.6, radius=0.12)
        text(s, x + 0.70, yy + 0.10, 2.3, 0.32, [[(head, {"bold": True, "size": 13.5, "color": INK})]])
        text(s, x + 0.70, yy + 0.40, w - 0.90, 0.30, [[(desc, {"size": 10.5, "color": MUTED})]])


def stuck_line(s, y=6.40):
    """막히면 — 세 단계를 한 줄로 (슬라이드 6 · 8 공통)"""
    rect(s, ML, y, CW, 0.42, fill=TINT, radius=0.10)
    text(s, ML + 0.3, y + 0.07, CW - 0.6, 0.3,
         [[("막히면   ", {"bold": True, "color": ACCENT_D}),
           ("① Ctrl + F5 새로고침   →   ② 무엇이 바뀌었나 보기   →   ③ 본 것과 바라는 것을 나눠서 다시 요청", {})]],
         size=11.5, color=BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


NOTES = {}


def note(i, lines):
    NOTES[i] = "\n".join(lines)


# ================================================================ 슬라이드
def build():
    prs = B.new_deck()
    S = lambda: B.blank(prs)
    n = lambda: B.CUR["n"]

    # ---------------------------------------------------------- 1. 표지
    s = S()
    text(s, ML, 0.44, 5, 0.3,
         [[("A I   ×   B U I L D", {"bold": True, "size": 11, "color": ACCENT, "space": 1.3})]])
    text(s, SW - ML - 7.4, 0.46, 7.4, 0.3,
         [[("한국디지털미디어고등학교 · 4차시", {"bold": True, "size": 10, "color": FAINT, "space": 1.0})]],
         align=PP_ALIGN.RIGHT)
    text(s, ML, 1.70, 2.2, 1.2, [[("04", {"bold": True, "size": 60, "color": ACCENT_L})]])
    text(s, ML, 2.84, CW, 1.0, [[("내 아이디어를 바로 만든다", {"bold": True, "size": 40, "color": INK})]])
    rect(s, ML, 3.94, 0.9, 0.05, fill=ACCENT)
    text(s, ML, 4.20, CW * 0.9, 0.5,
         [[("오늘 끝나면, 내 노트북에 실제로 돌아가는 작품이 하나 있습니다", {"size": 17, "color": BODY})]])
    w, gap = grid(3, 0.22)
    for i, (h1, h2) in enumerate([("보고", "선생님이 3분 만에 만드는 걸 본다"),
                                  ("만들고", "내 아이디어를 Claude Code로 만든다"),
                                  ("해 본다", "친구 작품을 직접 해 보고 서로 알려 준다")]):
        minicard(s, ML + i * (w + gap), 5.20, w, 1.20, h1, h2, bar=ACCENT if i == 1 else ACCENT_L)
    footer_rule(s)
    text(s, ML, FOOT_Y + 0.12, 8, 0.26,
         [[("코드는 AI가, 판단은 내가", {"size": 9.5, "color": FAINT, "space": 0.8})]])
    note(n(), [
        "● 여는 말 (30초)",
        "오늘은 설명을 길게 하지 않습니다. 바로 보여 주고, 바로 만들어요.",
        "90분 뒤에는 여러분 노트북에 여러분 아이디어로 만든 작품이 하나씩 있을 거예요.",
        "● 바로 다음 슬라이드로 넘어가 라이브 데모를 시작합니다.",
    ])

    # ---------------------------------------------------------- 2. 라이브 데모
    s = S()
    header(s)
    title_block(s, "먼저 보여 줄게요", "3분 안에 하나 만들어 보겠습니다")
    rect(s, ML, TOP, CW, 1.55, fill=DARK, radius=0.06)
    text(s, ML + 0.4, TOP + 0.30, CW - 0.8, 0.3,
         [[("여러분 차례", {"bold": True, "size": 10, "color": ACCENT_L, "space": 1.2})]])
    text(s, ML + 0.4, TOP + 0.66, CW - 0.8, 0.6,
         [[("만들고 싶은 거, 아무거나 말해 보세요", {"bold": True, "size": 26, "color": WHITE})]])
    step_row(s, 4.12, [
        ("①", "질문에 답한다", "Claude가 먼저 묻는 질문 3개에 답하기"),
        ("②", "계획을 읽는다", "우리가 말한 게 맞는지 한 줄씩 확인하고 승인"),
        ("③", "직접 해 본다", "파일을 열어서 끝까지 한 판"),
    ], h=1.50)
    banner(s, [[("코드는 한 줄도 쓰지 않습니다. 대신 ", {}),
                ("답하고, 읽고, 해 봅니다.", {"bold": True, "color": ACCENT_D}),
                (" 오늘 여러분도 이걸 합니다.", {})]],
           y=5.86, h=0.70)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 진행 (5분)",
        "학생 한 명에게 아이디어를 받습니다. 아무도 없으면 '급식에서 싫어하는 반찬 피하기 게임'.",
        "선생님 화면으로 전환 → 게시판의 첫 요청문(B-1) 첫 줄에 아이디어를 넣어 보냅니다.",
        "Claude가 질문 3개를 하면 반 전체에 물어서 답합니다. '화면에 뭐가 보여야 해?'",
        "계획이 나오면 소리 내어 읽고 승인합니다. '여기 이 줄, 우리가 말한 거 맞죠?'",
        "만들어지면 바로 열어서 한 판 해 봅니다.",
        "● 결과가 이상하게 나오면 오히려 좋습니다",
        "'봐요, AI도 틀려요. 이걸 고치는 게 오늘 할 일이에요.'",
        "● 마무리 말",
        "방금 제가 한 건 세 가지예요. 질문에 답했고, 계획을 읽었고, 직접 해 봤어요.",
    ])

    # ---------------------------------------------------------- 2-2. 대화창 vs Claude Code
    s = S()
    header(s)
    title_block(s, "오늘부터 도구가 바뀝니다", "1~3회기의 Claude 대화창  →  오늘부터 Claude 앱의 Code 탭")
    rows = [("질문에 답한다", "내 폴더의 파일을 직접 다룬다"),
            ("코드를 채팅창에 보여 준다", "파일을 직접 만들고 고친다"),
            ("한 번 만들어 보기에 좋다", "여러 회기에 걸쳐 키우는 작품에 좋다"),
            ("복사해서 파일에 붙여 넣어야 한다", "내 폴더 안에 바로 저장된다")]
    cw2 = (CW - 0.60) / 2
    rx = ML + cw2 + 0.60
    rect(s, ML, TOP, cw2, 0.46, fill=LINE2, radius=0.08)
    text(s, ML + 0.24, TOP + 0.08, cw2 - 0.48, 0.3, [[("Claude 대화창  ·  1~3회기", {"bold": True, "size": 12.5, "color": INK})]])
    rect(s, rx, TOP, cw2, 0.46, fill=ACCENT, radius=0.08)
    text(s, rx + 0.24, TOP + 0.08, cw2 - 0.48, 0.3, [[("Claude Code  ·  오늘부터", {"bold": True, "size": 12.5, "color": WHITE})]])
    for i, (a, b) in enumerate(rows):
        yy = TOP + 0.58 + i * 0.52
        rect(s, ML, yy, cw2, 0.44, fill=CARD, line=LINE, radius=0.08)
        text(s, ML + 0.24, yy + 0.07, cw2 - 0.48, 0.3, [[(a, {"size": 12, "color": BODY})]])
        text(s, ML + cw2 + 0.10, yy + 0.04, 0.40, 0.34, [[("→", {"bold": True, "size": 14, "color": FAINT})]], align=PP_ALIGN.CENTER)
        rect(s, rx, yy, cw2, 0.44, fill=CARD, line=ACCENT_L, radius=0.08)
        text(s, rx + 0.24, yy + 0.07, cw2 - 0.48, 0.3, [[(b, {"bold": True, "size": 12, "color": INK})]])
    w, gap = grid(2, 0.30)
    y = 5.02
    lcard(s, ML, y, w, 1.34, ACCENT, "우리가 쓰는 곳",
          ["Claude 앱 → Code 탭 → Local → 내 폴더 선택"], strong=True, size=11.5)
    lcard(s, ML + w + gap, y, w, 1.34, LINE2, "오늘은 쓰지 않는 곳",
          ["claude.ai 웹의 Code — 내 폴더 대신 GitHub 저장소를 요구"], size=11.5)
    rect(s, ML, 6.46, CW, 0.38, fill=TINT, radius=0.10)
    text(s, ML + 0.3, 6.50, CW - 0.6, 0.3,
         [[("대신 책임이 하나 생깁니다 — ", {}), ("승인하는 순간 내 파일이 진짜로 바뀝니다.", {"bold": True, "color": ACCENT_D})]],
         size=11.5, color=BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 이렇게 말해 보세요 (3분) — 방금 데모에서 본 것을 정리합니다",
        "1~3회기에는 대화창에서 만들었죠. 코드가 채팅창에 나오면 복사해서 붙여 넣었어요.",
        "방금 제 화면에서는 복사·붙여넣기가 없었어요. Claude가 제 폴더에 파일을 직접 만들었죠. 이게 Claude Code입니다.",
        "● 헷갈리기 쉬운 점",
        "claude.ai 웹사이트에도 Code 메뉴가 있지만, 그쪽은 GitHub 저장소를 고르라고 합니다. 오늘은 앱의 Code 탭 → Local 만 씁니다.",
        "화면에 'Select folder' 버튼이 보이면 맞게 온 것입니다.",
    ])

    # ---------------------------------------------------------- 3. 설치
    s = S()
    header(s)
    title_block(s, "설치는 네 단계입니다", "Claude 앱과 Git, 두 개만 설치합니다. 다 되면 손을 들어 주세요")
    step_row(s, TOP, [
        ("STEP 1", "설치 파일 두 개", "Claude 앱 · Git — 주소는 게시판에 (또는 USB)"),
        ("STEP 2", "받는 동안", "바탕화면에 myapp 폴더 만들기 · 만들고 싶은 것 한 문장 쓰기"),
        ("STEP 3", "설치 · 로그인", "Git은 Next만 누르기 · 앱은 수업 계정으로 로그인"),
        ("STEP 4", "폴더 열기", "Code → Local → Select folder → myapp"),
    ], h=1.70)
    w, gap = grid(2, 0.30)
    y = 4.24
    lcard(s, ML, y, w, 2.52, ACCENT_L, "막히면 이렇게",
          ["5분 안에 안 되면 짝 노트북에 합류한다",
           "학교 와이파이가 안 되면 휴대폰 핫스팟",
           "Node.js는 설치하지 않는다 — 앱 + Git이면 끝"],
          note="오늘은 둘이 하나를 만들어도 괜찮습니다.")
    lcard(s, ML + w + gap, y, w, 2.52, GREEN, "먼저 끝났다면",
          ["바로 도우미가 되어 준다",
           "대신 해 주지 말고 '어디까지 됐어?'부터 묻는다",
           "두 명까지만 돕고 자기 자리로 돌아온다"],
          label_color=GREEN_D)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 운영 (20분)",
        "25명이 동시에 내려받으면 학교 네트워크가 먼저 멈춥니다. USB를 먼저 돌리고 나머지를 분산하세요.",
        "STEP 2가 핵심입니다. 설치를 기다리는 동안 폴더 만들기와 아이디어 한 문장을 먼저 시킵니다. (다음 슬라이드)",
        "Git 설치 화면은 선택지가 많지만 바꿀 것이 없습니다. '전부 Next, 마지막에 Install'이라고만 말해 주세요.",
        "Git은 오늘 후반부 '세이브 포인트'에서 씁니다. 지금은 설치만 해 둡니다.",
        "인터넷 자료를 보고 Node.js를 미리 깔아 온 학생이 있어도 그대로 둡니다. 지우지 않아도 됩니다.",
        "● 18분 체크포인트",
        "'폴더 열린 사람?' 3분의 1이 안 되면 라운드 ②의 레벨업 카드를 빼고 그 시간을 여기에 씁니다.",
    ])

    # ---------------------------------------------------------- 4. 기다리는 동안
    s = S()
    header(s)
    title_block(s, "설치되는 동안 — 만들고 싶은 것 한 문장", "포스트잇이나 메모장에 적어 두세요. 첫 요청문의 첫 줄이 됩니다")
    lw_ = 4.35
    wcard(s, ML, TOP, lw_, 4.10, ACCENT, "한 문장 쓰는 법", "오늘은 '한 판'만",
          ["무엇을 하는지 한 문장으로",
           "1분 안에 한 판이 끝나는 크기로",
           "큰 아이디어라면 그중 한 장면만",
           "2회기 PRD.md가 있으면 myapp 폴더에 넣어 두기"],
          note="예) 떨어지는 싫은 반찬을 피하고, 좋아하는 반찬은 받는 게임", strong=True)
    gx = ML + lw_ + 0.30
    gw = CW - lw_ - 0.30
    text(s, gx, TOP - 0.02, gw, 0.3,
         [[("떠오르지 않으면 여기서 고르세요", {"bold": True, "size": 10, "color": ACCENT_D, "space": 1.0})]])
    ideas = [("급식 피하기", "싫은 반찬은 피하고 좋아하는 반찬은 받기"),
             ("지각 탈출 러너", "장애물을 점프로 넘으며 교실까지"),
             ("선생님 몰래 폰 보기", "돌아볼 때 멈추는 무궁화 꽃이 방식"),
             ("타자 디펜스", "내려오는 코딩 용어를 타이핑해서 막기"),
             ("반응속도 테스트", "화면이 초록색이 되는 순간 누르기"),
             ("시험 D-day 타이머", "남은 날짜와 오늘 공부 시간 재기")]
    cw2 = (gw - 0.20) / 2
    ch = 1.18
    for i, (h1, h2) in enumerate(ideas):
        cx = gx + (i % 2) * (cw2 + 0.20)
        cy = TOP + 0.38 + (i // 2) * (ch + 0.14)
        minicard(s, cx, cy, cw2, ch, h1, h2)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 설치를 기다리는 동안 띄워 둡니다",
        "아이디어 고르기에 오래 쓰지 않게 하세요. 2분 안에 못 정하면 목록에서 하나를 고릅니다.",
        "'롤 같은 거' 처럼 큰 아이디어 → '그중에 한 판, 1분짜리만. 나머지는 다음 시간에 붙이자.'",
    ])

    # ---------------------------------------------------------- 5. 규칙 세 개
    s = S()
    header(s)
    title_block(s, "규칙은 세 개뿐입니다", "이것만 지키면 오늘 오류의 대부분을 피합니다")
    w, gap = grid(3, 0.26)
    rules = [
        (ACCENT, ACCENT_D, "규칙 1", "HTML 파일 하나로",
         ["더블클릭하면 바로 열린다",
          "다음 시간에 휴대폰으로 올릴 수 있다",
          "React · npm이 나오면 설치 오류가 난다"]),
        (BLUE, BLUE_D, "규칙 2", "계획을 읽고 승인",
         ["승인하는 순간 내 파일이 바뀐다",
          "읽지 않고 승인하는 건 남이 쓴 글에 서명하는 것",
          "이상하면 거절하면 된다. 파일은 그대로"]),
        (GREEN, GREEN_D, "규칙 3", "한 번에 하나만",
         ["세 가지를 한꺼번에 시키면",
          "무엇 때문에 망가졌는지 알 수 없다",
          "하나 고치고, 해 보고, 다음 하나"]),
    ]
    for i, (bar, lc, label, head, lines) in enumerate(rules):
        wcard(s, ML + i * (w + gap), TOP, w, 3.00, bar, label, head, lines, label_color=lc, size=11.5)
    banner(s, [[("React나 npm 이야기가 나오면 → ", {}),
                ("\"외부 도구 없이 HTML 파일 하나로 다시 만들어 줘\"", {"bold": True, "color": ACCENT_D})]],
           y=5.62, h=0.80)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 이렇게 말해 보세요 (5분)",
        "규칙은 세 개뿐이에요. 이유도 하나씩 있어요.",
        "고치는 방법은 지금 설명하지 않을게요. 직접 막혀 보고 나서 같이 봅시다.",
        "● 참고",
        "기존 안의 '수정 네 칸', '수정 유형 ①~⑤'는 여기서 하지 않습니다. 슬라이드 7(손 멈춤)에서 실제 화면으로 가르칩니다.",
    ])

    # ---------------------------------------------------------- 5-2. 모드 네 가지
    s = S()
    header(s)
    title_block(s, "Claude에게 얼마나 맡길까 — 모드 네 가지", "전송 버튼 옆 모드 선택에서 고릅니다.  오늘은 수동(Manual)")
    modes = [
        ("MANUAL · 수동", "매번 물어본다", ACCENT, ACCENT_D, True,
         ["파일 고치기 — 물어봄", "명령 실행 — 물어봄", "처음 쓸 때 · 연습할 때"]),
        ("PLAN · 계획", "계획만 세운다", BLUE, BLUE_D, False,
         ["파일 고치기 — 안 함", "읽고 계획만 보여 줌", "무엇을 할지 먼저 볼 때"]),
        ("ACCEPT EDITS · 편집 허용", "파일은 바로 고친다", GREEN, GREEN_D, False,
         ["파일 고치기 — 묻지 않음", "명령 실행 — 물어봄", "손에 익은 뒤 (5회기~)"]),
        ("AUTO · 자동", "거의 묻지 않는다", RED, RED, False,
         ["파일 고치기 — 묻지 않음", "명령 실행 — 대부분 묻지 않음", "수업에서는 쓰지 않음"]),
    ]
    w, gap = grid(4, 0.20)
    for i, (label, head, bar, lc, rec, lines) in enumerate(modes):
        x = ML + i * (w + gap)
        wcard(s, x, TOP, w, 2.92, bar, label, head, lines, label_color=lc, size=11, strong=rec)
        if rec:
            rect(s, x + w - 1.02, TOP + 0.22, 0.86, 0.30, fill=ACCENT, radius=0.30)
            text(s, x + w - 1.02, TOP + 0.24, 0.86, 0.26, [[("오늘 추천", {"bold": True, "size": 9.5, "color": WHITE})]],
                 align=PP_ALIGN.CENTER)
    banner(s, [[("오늘은 ", {}), ("수동(Manual)", {"bold": True, "color": ACCENT_D}),
                (" — 확인 창이 떠야 '계획 읽고 승인 · 이상하면 거절'을 연습할 수 있습니다.", {})],
               [("계획만 먼저 보고 싶으면 계획 모드.  자동은 승인할 기회 없이 바로 바뀌므로 ", {}),
                ("수업에서 쓰지 않습니다.", {"bold": True, "color": ACCENT_D})]],
           y=5.44, h=0.98)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 이렇게 말해 보세요 (2분)",
        "Claude에게 얼마나 맡길지 정하는 스위치가 있어요. 전송 버튼 옆에 있습니다.",
        "오늘은 '수동'이에요. 파일을 고치기 전에, 명령을 실행하기 전에 매번 물어봅니다. 그 확인 창이 여러분이 판단하는 자리예요.",
        "'자동'으로 두면 확인 없이 바로 바뀝니다. 편하지만, 무엇이 바뀌었는지 모르고 지나가게 돼요. 수업에서는 쓰지 않습니다.",
        "● 다 같이 확인",
        "'모드가 Manual(수동)로 되어 있는 사람 손?' — 손을 들게 해서 전원 확인하고 넘어갑니다.",
        "● 참고",
        "편집 허용(Accept edits)은 고치는 요청이 손에 익은 5회기 이후, 드라이버가 써도 됩니다. 명령 실행은 여전히 물어봅니다.",
        "모드 이름과 위치는 앱 버전에 따라 조금 다를 수 있습니다. 수업 전에 선생님 앱에서 한 번 확인하세요.",
    ])

    # ---------------------------------------------------------- 6. 미션 ① 만들기 — 읽기 · 이해하기
    s = S()
    header(s)
    title_block(s, "미션 ①  만들기 전에, 먼저 읽게 하기", "20분 · 바로 \"만들어 줘\" 하지 않습니다.  읽기 → 이해하기 → 확인하기 → 만들기")
    cw_ = 7.05
    codebox(s, ML, TOP, cw_, 3.92, "요청 ① 읽고 묻게 하기 — 게시판에서 복사",
            ["아직 파일을 만들거나 고치지 마.",
             "먼저 이 폴더를 확인하고 알려 줘.",
             "",
             "1. 폴더 안에 어떤 파일이 있는지",
             "2. PRD.md가 있다면 — 만들려는 것, 사용자,",
             "   Must 기능, Later 기능을 짧게 정리",
             "3. PRD.md가 없다면 — 내 아이디어: [한 문장]",
             "4. 만들기 전에 나에게 물어볼 질문 3개",
             "   (화면에 보일 것 · 조작 방법 · 끝나는 조건)"])
    checklist(s, ML + cw_ + 0.26, TOP, CW - cw_ - 0.26, [
        ("1  읽기", "요청 ① 보내기 — 아직 아무것도 안 만든다"),
        ("2  이해하기", "Claude의 정리가 내 생각과 맞는지 읽기"),
        ("3  답하기", "질문 3개에 직접 답한다 = 오늘의 설계도"),
        ("4  확인하기", "요청 ② 계획 받기 → 읽고 승인 (다음 장)"),
        ("5  해 보기", "index.html 더블클릭 → 끝까지 한 판"),
    ], row_h=0.70, gap=0.105)
    stuck_line(s, y=6.40)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 실습 내내 이 슬라이드와 다음 장을 번갈아 띄웁니다 (미션 ① 전체 20분)",
        "핵심은 '바로 만들어 줘'를 하지 않는 것입니다. AI가 먼저 읽고, 이해한 것을 말하게 하고, 그게 맞는지 학생이 판단합니다.",
        "2회기 PRD.md를 폴더에 넣은 학생은 Claude가 PRD를 요약해 줍니다. 없는 학생은 아이디어 한 문장으로 시작합니다. 둘 다 괜찮습니다.",
        "질문에 '알아서 해 줘'라고 답하는 학생 → '그럼 Claude가 추측해. 네 게임인데 Claude 마음대로 돼도 괜찮아?'",
        "● 순회하면서 꼭 할 일",
        "'AI가 내 말을 엉뚱하게 알아들은' 학생 화면을 하나 골라 둡니다. 손 멈춤(말 바꾸기)에서 씁니다.",
        "미리 양해를 구하세요. '이 화면 잠깐 같이 봐도 될까? 정말 좋은 예라서.'",
    ])

    # ---------------------------------------------------------- 6-2. 미션 ① — 계획 먼저 · 확인하기
    s = S()
    header(s)
    title_block(s, "요청 ②  계획을 먼저 보여 줘", "질문에 다 답했다면 — 코드보다 계획을 먼저 받고, 읽고, 승인합니다")
    cw_ = 7.05
    codebox(s, ML, TOP, cw_, 3.92, "요청 ② 계획 먼저 — 게시판에서 복사",
            ["내 답(또는 PRD.md)을 기준으로 이번에 만들 것의",
             "계획을 보여 줘. 조건은 다음과 같아.",
             "- index.html 파일 하나로, 브라우저에서 바로 실행",
             "- 외부 라이브러리 · API · 로그인 · 서버 없이",
             "- Must 기능만 먼저, Later 기능은 아직 만들지 마",
             "- 휴대폰과 PC에서 모두 쓸 수 있게",
             "",
             "아직 코드는 만들지 말고 먼저 보여 줘.",
             "1. 만들 기능  2. 만드는 순서  3. 이번에 뺄 기능",
             "내가 승인하면 그때 만들어 줘."])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    wcard(s, rx, TOP, rw, 2.16, GREEN, "이러면 승인", "계획이 내 말과 같다",
          ["내가 답한 내용이 그대로 있다", "파일은 index.html 하나"], label_color=GREEN_D, size=11)
    wcard(s, rx, TOP + 2.26, rw, 1.82, RED, "이러면 거절", "말하지 않은 게 있다",
          ["로그인 · 서버 · 라이브러리 · 파일 여러 개"], label_color=RED, size=11)
    rect(s, ML, 6.50, CW, 0.38, fill=TINT, radius=0.10)
    text(s, ML + 0.3, 6.53, CW - 0.6, 0.32,
         [[("AI가 ", {}), ("무엇을 만들려는지 먼저 검토", {"bold": True, "color": ACCENT_D}),
           ("합니다.  결과를 기다리기만 하지 않습니다.", {})]],
         size=11.5, color=BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 이렇게 말해 보세요",
        "질문에 답했으면 이제 계획을 받습니다. 아직 코드는 안 만들어요. 계획을 읽고, 내가 말한 것과 같을 때만 승인합니다.",
        "● 순회 포인트",
        "계획을 읽지 않고 바로 승인하는 학생 → '방금 승인한 계획에 뭐라고 써 있었어? 빠진 거나 더해진 거 없어?'",
        "계획에 파일이 여러 개(style.css, script.js …)로 나뉘어 있으면 → 거절하고 'index.html 하나로' 다시 요청",
        "React · npm · Node 가 보이면 → 거절하고 '외부 도구 없이 HTML 파일 하나로'",
        "● 승인하면",
        "Claude가 만들기 시작합니다. 끝나면 index.html을 더블클릭해서 끝까지 해 보고, 이상하거나 아쉬운 점 하나를 적게 하세요.",
    ])

    # ---------------------------------------------------------- 7. 손 멈춤 · 말 바꾸기
    s = S()
    header(s)
    title_block(s, "AI 잘못일까, 말 잘못일까?", "손 멈춤 — 방금 친구 화면에서 본 것으로 5분만")
    rows = [("\"더 예쁘게 해 줘\"", "\"제목을 가운데로, 두 배 크게 해 줘\""),
            ("\"안 돼. 고쳐 줘\"", "\"스페이스바를 눌러도 점프를 안 해. 화면에 오류는 없어\""),
            ("\"버튼 잘 되게 해 줘\"", "\"누를 때마다 문구가 순서대로 바뀌게 해 줘\"")]
    bw = 3.70
    text(s, ML, TOP - 0.04, bw, 0.28, [[("이렇게 말하면", {"bold": True, "size": 10, "color": RED, "space": 1.0})]])
    text(s, ML + bw + 0.60, TOP - 0.04, 4, 0.28, [[("이렇게 바꾸기", {"bold": True, "size": 10, "color": GREEN_D, "space": 1.0})]])
    for i, (bad, good) in enumerate(rows):
        yy = TOP + 0.32 + i * 0.70
        rect(s, ML, yy, bw, 0.58, fill=RED_SOFT, radius=0.10)
        text(s, ML + 0.24, yy + 0.08, bw - 0.48, 0.46, [[(bad, {"size": 13, "color": RED})]], anchor=MSO_ANCHOR.MIDDLE)
        text(s, ML + bw + 0.08, yy + 0.12, 0.44, 0.4, [[("→", {"bold": True, "size": 16, "color": FAINT})]], align=PP_ALIGN.CENTER)
        gx = ML + bw + 0.60
        rect(s, gx, yy, CW - bw - 0.60, 0.58, fill=GREEN_SOFT, radius=0.10)
        text(s, gx + 0.24, yy + 0.08, CW - bw - 1.08, 0.46, [[(good, {"bold": True, "size": 13, "color": GREEN_D})]], anchor=MSO_ANCHOR.MIDDLE)
    codebox(s, ML, 4.72, CW, 1.64, "고치는 말 — 게시판에서 복사",
            ["지금: [본 것 — 예) 스페이스바를 눌러도 점프하지 않아]",
             "원하는 것: [예) 누르면 한 번 점프하게 해 줘]",
             "[고칠 부분]만 고치고, 다른 부분은 바꾸지 마. 먼저 계획을 보여 줘."])
    rect(s, ML, 6.46, CW, 0.38, fill=TINT, radius=0.10)
    text(s, ML + 0.3, 6.50, CW - 0.6, 0.3,
         [[("공통점 하나 — ", {}), ("본 것", {"bold": True, "color": ACCENT_D}), ("과 ", {}),
           ("바라는 것", {"bold": True, "color": ACCENT_D}), ("을 나눠서 말하고, 끝에 ", {}),
           ("\"다른 부분은 바꾸지 마\"", {"bold": True, "color": ACCENT_D})]],
         size=11.5, color=BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 진행 (5분)",
        "골라 둔 학생 화면을 띄웁니다. '○○이는 이렇게 말했고, AI는 이렇게 만들었어요.'",
        "반 전체에 묻습니다. 'AI 잘못일까요, 말 잘못일까요?'",
        "그다음에 이 슬라이드의 세 줄을 보여 줍니다. 반 친구 사례가 먼저, 규칙은 나중입니다.",
        "● 강조",
        "'다른 부분은 바꾸지 마' 한 줄이 멀쩡한 곳을 지켜 줍니다. 범위를 안 적으면 한 군데 고치다 다른 곳까지 바뀝니다.",
    ])

    # ---------------------------------------------------------- 7-2. 세이브 포인트 · Git과 커밋
    s = S()
    header(s)
    title_block(s, "고치기 전에, 세이브 포인트", "지금 잘 돌아가는 상태를 저장해 둡니다. 망가져도 여기로 돌아올 수 있게")
    w, gap = grid(3, 0.26)
    for i, (bar, lc, label, head, lines) in enumerate([
        (ACCENT, ACCENT_D, "GIT", "세이브 기능", ["내 폴더에 기록을 남겨 주는 프로그램"]),
        (BLUE, BLUE_D, "커밋", "세이브 한 번", ["지금 상태를 저장하고 이름표 붙이기"]),
        (GREEN, GREEN_D, "되돌리기", "세이브 불러오기", ["저장해 둔 그때 그대로 돌아가기"]),
    ]):
        wcard(s, ML + i * (w + gap), TOP, w, 1.82, bar, label, head, lines, label_color=lc, size=11)
    codebox(s, ML, 4.28, CW, 2.06, "Claude에게 이렇게 말하면 됩니다 — 명령어는 몰라도 됩니다",
            ["[지금 한 번]     이 폴더를 Git으로 기록하기 시작하고, 지금 상태를 커밋해 줘.",
             "                 메시지는 \"첫 화면 완성\", 작성자 이름은 [닉네임]으로.",
             "[잘 될 때마다]   지금 상태를 커밋해 줘. 메시지는 \"[무엇을 했는지]\".",
             "[기록 보기]      지금까지 커밋 기록을 보여 줘.",
             "[되돌리기]       먼저 지금 상태도 커밋하고, \"[메시지]\" 커밋 상태로 되돌려 줘."])
    rect(s, ML, 6.44, CW, 0.40, fill=TINT, radius=0.10)
    text(s, ML + 0.3, 6.48, CW - 0.6, 0.32,
         [[("\"다시 만들어 줘\"", {"bold": True}), ("는 새로 만들기 — 매번 다르게 나옵니다.   ", {}),
           ("커밋으로 되돌리기", {"bold": True, "color": ACCENT_D}), ("는 그때 그대로 — 인터넷 없이 내 컴퓨터에 남습니다.", {})]],
         size=11.5, color=BODY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 진행 (5분) — 손 멈춤 바로 다음, 미션 ② 전에",
        "이렇게 말해 보세요: '방금 만든 거, 마음에 드는 상태죠? 이제 고칠 건데, 고치다 망가지면 어떡하죠?'",
        "'다시 만들어 줘' 하면 새 게임이 나와요. 아까 그 게임은 다시 안 나옵니다. 그래서 고치기 전에 세이브합니다.",
        "● 같이 하기",
        "[지금 한 번] 요청문을 반 전체가 동시에 보냅니다. 계획을 읽고 승인 → '커밋했다'는 답이 오면 손 들기.",
        "작성자 이름은 실명 대신 닉네임으로. 나중에 GitHub에 올리면 공개될 수 있습니다.",
        "● 알아 둘 것",
        "기록은 myapp 폴더 안의 숨은 .git 폴더에 저장됩니다. 인터넷 · GitHub 계정이 필요 없습니다. .git 폴더는 지우지 않게 하세요.",
        "앱 화면에 커밋 버튼은 따로 없습니다. 말로 시키면 Claude가 Git을 대신 실행합니다.",
    ])

    # ---------------------------------------------------------- 8. 미션 ② 고치기
    s = S()
    header(s)
    title_block(s, "미션 ②  고치고, 세이브하고, 멈춰 보기", "15분 · 하나 고치고 → 해 보고 → 잘 되면 커밋 → 다음 하나")
    lw_ = 6.55
    checklist(s, ML, TOP, lw_, [
        ("고치기 ①", "눈에 보이는 것 하나 → 잘 되면 커밋"),
        ("고치기 ②", "동작 하나 → 잘 되면 커밋"),
        ("거절 한 번", "요청 안 한 게 계획에 있으면 거절 (없으면 \"로그인도 추가해 줘\" → 거절)"),
        ("버전 기록 보기", "\"지금까지 커밋 기록을 보여 줘\" — 내 세이브가 몇 개?"),
    ], row_h=0.80, gap=0.14)
    rx = ML + lw_ + 0.26
    codebox(s, rx, TOP, CW - lw_ - 0.26, 3.70, "멈추고 싶을 때 쓰는 말",
            ["그건 요청하지 않았어. 추가하지 마.",
             "지금 화면 그대로 둬.",
             "",
             "index.html만 고쳐 줘.",
             "다른 파일은 만들지 마.",
             "",
             "한 번에 하나만 해 줘.",
             "",
             "[망했을 때 — 세이브 불러오기]",
             "\"첫 화면 완성\" 커밋으로 되돌려 줘."])
    stuck_line(s, y=6.40)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 실습 내내 이 슬라이드를 띄워 둡니다 (15분)",
        "고치기가 잘 끝날 때마다 커밋하게 하세요. '잘 될 때 세이브'가 습관이 되면 6회기부터 되돌리기가 쉬워집니다.",
        "다 끝낸 학생은 레벨업 카드(다음 장)로 보냅니다.",
        "순회하면서 두 질문만 반복합니다. '무엇이 바뀌었어?' '원한 거랑 맞아?'",
        "계획을 읽지 않고 바로 승인하는 학생 → '방금 승인한 계획에 뭐라고 써 있었어?'",
        "세 가지를 한 번에 시키는 학생 → '하나만. 무엇이 망가뜨렸는지 모르게 돼.'",
        "● 거절 연습",
        "거절한 뒤 화면이 그대로인지 직접 확인하게 하세요. 승인하지 않으면 파일이 바뀌지 않는다는 걸 눈으로 보는 게 목적입니다.",
    ])

    # ---------------------------------------------------------- 9. 레벨업 카드
    s = S()
    header(s)
    title_block(s, "레벨업 카드", "미션 ②를 끝냈다면 하나씩 골라서 — 모두 데모데이 합격선에 들어 있는 항목입니다")
    cards = [
        ("LEVEL UP 1", ACCENT_L, "휴대폰으로도", ["\"휴대폰에서 화면을 터치해도 조작되게 해 줘\""]),
        ("LEVEL UP 2", RGBColor(0x8F, 0xC1, 0xF5), "점점 어렵게", ["\"시간이 지날수록 조금씩 빨라지게 해 줘\""]),
        ("LEVEL UP 3", RGBColor(0xF5, 0xC8, 0x6B), "최고 점수", ["\"최고 점수를 저장해서 새로고침해도 남게 해 줘\""]),
        ("LEVEL UP 4", RGBColor(0xF2, 0x9B, 0x9B), "손맛", ["\"맞았을 때 화면이 살짝 흔들리게 해 줘\""]),
        ("LEVEL UP 5", RGBColor(0x8F, 0xD9, 0xB6), "한 바퀴", ["\"시작 화면 → 플레이 → 결과 → 다시 하기가 되게 해 줘\""]),
    ]
    w, gap = grid(5, 0.18)
    for i, (label, col, head, lines) in enumerate(cards):
        dcard(s, ML + i * (w + gap), TOP, w, 2.45, label, col, head, lines, size=11)
    banner(s, [[("한 장씩 — 고치고, 해 보고, 커밋하고 다음 장. ", {}),
                ("오늘 해 둔 것은 데모데이까지 그대로 씁니다.", {"bold": True, "color": ACCENT_D})]],
           y=5.06, h=0.80)
    footer_rule(s)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 빨리 끝낸 학생용",
        "레벨업 카드는 순서가 없습니다. 내 작품에 어울리는 것부터 고르게 하세요.",
        "카드 내용이 데모데이 합격선(휴대폰 실행, 점점 어려워짐, 최고 점수 저장, 반응 2가지 이상, 한 바퀴)과 같다는 걸 짚어 주세요.",
        "● 시간이 부족한 반",
        "18분 체크포인트에서 설치가 늦었다면 이 슬라이드는 건너뜁니다.",
    ])

    # ---------------------------------------------------------- 10. 갤러리 워크
    s = S()
    header(s)
    title_block(s, "갤러리 워크 — 서로 해 보기", "10분 · 설명하지 말고, 그냥 해 보게 두세요")
    step_row(s, TOP, [
        ("1", "작품 열어 두기", "내 작품을 화면에 띄우고 노트북을 그대로 둔다"),
        ("2", "반은 돌아다니기", "4분 동안 최소 두 작품을 직접 해 본다. 그다음 교대"),
        ("3", "포스트잇 한 장", "해 본 작품 옆에 붙이고 간다"),
    ], h=1.55)
    w, gap = grid(2, 0.30)
    y = 4.10
    rect(s, ML, y, w, 1.60, fill=GREEN_SOFT, radius=0.08)
    text(s, ML + 0.35, y + 0.28, w - 0.7, 0.3, [[("포스트잇 위쪽", {"bold": True, "size": 10, "color": GREEN_D, "space": 1.0})]])
    text(s, ML + 0.35, y + 0.68, w - 0.7, 0.6, [[("재밌던 점 하나", {"bold": True, "size": 22, "color": INK})]])
    rect(s, ML + w + gap, y, w, 1.60, fill=RGBColor(0xFF, 0xF4, 0xE0), radius=0.08)
    text(s, ML + w + gap + 0.35, y + 0.28, w - 0.7, 0.3, [[("포스트잇 아래쪽", {"bold": True, "size": 10, "color": RGBColor(0xB4, 0x53, 0x09), "space": 1.0})]])
    text(s, ML + w + gap + 0.35, y + 0.68, w - 0.7, 0.6, [[("헷갈리거나 답답했던 점 하나", {"bold": True, "size": 22, "color": INK})]])
    banner(s, [[("\"이거 어떻게 해?\"라는 질문이 나오면, ", {}),
                ("그게 바로 다음에 고칠 점", {"bold": True, "color": ACCENT_D}), ("입니다.", {})]],
           y=5.92, h=0.66)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 운영 (10분)",
        "반은 앉아 있고 반은 돌아다닙니다. 4분 뒤 타이머가 울리면 교대합니다.",
        "앉아 있는 사람은 설명하지 않습니다. 친구가 헤매는 곳을 보기만 합니다.",
        "● 연결",
        "7회기 플레이테스트의 예고편입니다. 받은 포스트잇은 버리지 말고 사진으로 남겨 두게 하세요.",
        "다음 시간 조별 작품을 고를 때, 포스트잇 반응이 좋았던 작품이 후보가 됩니다.",
    ])

    # ---------------------------------------------------------- 11. 마무리
    s = S()
    header(s)
    title_block(s, "오늘 마무리", "5분 · 오늘이 우리 프로젝트의 1일 차입니다")
    w, gap = grid(3, 0.26)
    wcard(s, ML, TOP, w, 2.20, GREEN, "1  백업", "폴더를 지키기",
          ["myapp 폴더를 통째로 압축해서", "USB나 내 메일로 — 버전 기록도 함께"], label_color=GREEN_D)
    wcard(s, ML + (w + gap), TOP, w, 2.20, ACCENT, "2  한 줄 기록", "AI가 틀린 순간",
          ["아래 요청문으로 개발일지.md 첫 줄", "9회기 개발 일지가 여기서 시작"], strong=True)
    wcard(s, ML + 2 * (w + gap), TOP, w, 2.20, BLUE, "3  다음 시간", "우리 조 작품 정하기",
          ["오늘 만든 것 중 조 작품 고르기", "포스트잇 반응이 근거"], label_color=BLUE_D)
    codebox(s, ML, 4.62, CW, 1.76, "한 줄 기록 — 게시판에서 복사",
            ["개발일지.md 파일을 만들고, 오늘 날짜로 아래 한 줄을 적어 줘. 다른 파일은 바꾸지 마.",
             "- AI가 틀렸던 순간: [         ]  /  어떻게 고쳤나: [         ]",
             "다 적었으면 지금 상태를 커밋해 줘. 메시지는 \"4회기 마무리\"."])
    text(s, ML, 6.48, CW, 0.3,
         [[("노트북과 충전기를 꼭 가져오세요.  백업한 폴더도 함께.", {"bold": True, "size": 11.5, "color": ACCENT_D})]],
         align=PP_ALIGN.CENTER)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 마무리 (5분)",
        "백업 → 손을 들어 확인. 노트북을 안 가져오는 학생이 반드시 생깁니다. 백업이 있으면 그날 바로 이어서 할 수 있습니다.",
        "폴더를 통째로 압축하면 숨은 .git 폴더(버전 기록)도 함께 갑니다. 파일만 골라 보내면 기록이 빠집니다.",
        "백업 전에 마지막으로 한 번 커밋하게 하세요. '오늘 마무리' 같은 메시지로.",
        "한 줄 기록은 '틀린 게 없었다'면 '내 말을 AI가 다르게 알아들은 순간'으로 바꿔 쓰게 하세요.",
        "● 예고",
        "다음 시간에는 오늘 만든 것 중에서 우리 조 작품을 고릅니다. 포스트잇 반응을 근거로요.",
    ])

    # ---------------------------------------------------------- 노트 붙이기
    for idx, sl in enumerate(prs.slides, start=1):
        if idx in NOTES:
            sl.notes_slide.notes_text_frame.text = NOTES[idx]

    prs.save(os.path.join(ROOT, DECK))
    return prs


if __name__ == "__main__":
    B.CUR["n"] = 0
    del B.WARN[:]
    prs = build()
    print("[v2] %s" % DECK)
    print("  슬라이드 %d장 · 발표자 노트 %d개" % (len(prs.slides), len(NOTES)))
    if B.WARN:
        print("  [카드 넘침] %d건" % len(B.WARN))
        for w_ in B.WARN:
            print("    " + w_)
    probs = [x for x in V3.lint(prs) if "상자 1.20 in : 04" not in x]
    if probs:
        print("  [글상자 검사] %d건" % len(probs))
        for pr in probs:
            print("    " + pr)
    if not B.WARN and not probs:
        print("  넘침 없음")
