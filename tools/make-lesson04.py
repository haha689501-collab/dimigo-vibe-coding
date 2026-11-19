# -*- coding: utf-8 -*-
"""
4차시 PPT 생성기 — 「Claude Code로 새로 시작하기 : 만들고, 바뀐 것을 읽고, 고친다」

실행:  python tools/make-lesson04.py
결과:  디미고_바이브코딩_4차시_ClaudeCode로새로시작하기.pptx  (프로젝트 최상위, 새 파일)

- 3차시까지의 결과물은 쓰지 않습니다. 빈 폴더에서 새 프로젝트로 시작합니다.
- 학생들은 각자 노트북을 쓰고, 설치는 수업 전 숙제로 끝낸 상태를 전제로 합니다.
- 색·여백·카드 부품은 1~3차시 덱 생성기(tools/make-lesson03-v2.py)의 것을 그대로 씁니다.
- 기존 PPT는 열지도 수정하지도 않습니다. 새 파일만 만듭니다.
"""
import importlib.util
import os
import sys

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

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

# 두 가지 버전을 함께 만든다.
#   A 숙제형     : 설치를 수업 전에 끝내 온 경우 (실습 시간이 길다)
#   B 수업중설치형: 설치부터 90분 안에 끝내는 경우 (고치기를 두 번으로 줄인다)
VARIANTS = {
    "A": {"deck": "디미고_바이브코딩_4차시_A_설치는숙제.pptx", "homework": True},
    "B": {"deck": "디미고_바이브코딩_4차시_B_수업중설치.pptx", "homework": False},
}
HOMEWORK = True   # build() 안에서 바뀐다

ML, CW, SW, SH, TOP, FOOT_Y = B.ML, B.CW, B.SW, B.SH, B.TOP, B.FOOT_Y
INK, BODY, MUTED, FAINT = B.INK, B.BODY, B.MUTED, B.FAINT
ACCENT, ACCENT_D, ACCENT_L = B.ACCENT, B.ACCENT_D, B.ACCENT_L
LINE, LINE2, CARD, SOFT, TINT = B.LINE, B.LINE2, B.CARD, B.SOFT, B.TINT
DARK, WHITE, BLUE, GREEN, PURPLE = B.DARK, B.WHITE, B.BLUE, B.GREEN, B.PURPLE
KR, MONO = B.KR, B.MONO
rect, text, header, footer_rule, title_block, banner = B.rect, B.text, B.header, B.footer_rule, B.title_block, B.banner
wcard, dcard, codebox, step_row, part_slide = B.wcard, B.dcard, B.codebox, B.step_row, B.part_slide

BLUE_D = RGBColor(0x3B, 0x53, 0x9B)
GREEN_D = RGBColor(0x2F, 0x7A, 0x55)
PURPLE_D = RGBColor(0x6B, 0x4A, 0x96)
TOTAL_PARTS = 3


def lcard(s, x, y, w, h, topbar, label, lines, note=None, label_color=None, size=11.5, strong=False):
    """제목 줄 없이 라벨만 두는 카드 — 세로 공간이 빠듯할 때."""
    wcard(s, x, y, w, h, topbar, label, None, lines, note=note,
          label_color=label_color, size=size, strong=strong)


def grid(n, gap):
    _xs, w = V3.cols(n, gap)
    return w, gap


def minicard(s, x, y, w, h, head, desc, bar=None):
    rect(s, x, y, w, h, fill=CARD, line=LINE, radius=0.06)
    if bar:
        rect(s, x, y, w, 0.065, fill=bar)
    text(s, x + 0.24, y + 0.24, w - 0.48, 0.34, [[(head, {"bold": True, "size": 14, "color": INK})]])
    text(s, x + 0.24, y + 0.64, w - 0.48, h - 0.74, [[(desc, {"size": 11, "color": MUTED})]], line=1.4)


NOTES = {}


def note(i, lines):
    NOTES[i] = "\n".join(lines)


# ================================================================ 슬라이드
def build(deck_name):
    hw = HOMEWORK                      # True = 설치를 숙제로 끝내 온 경우
    fix_n = "세 번" if hw else "두 번"   # 오늘 고치는 횟수
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
    text(s, ML, 2.84, CW, 1.0, [[("Claude Code로 새로 시작하기", {"bold": True, "size": 40, "color": INK})]])
    rect(s, ML, 3.94, 0.9, 0.05, fill=ACCENT)
    text(s, ML, 4.20, CW * 0.86, 0.5,
         [[("만들고, 바뀐 것을 읽고, 내가 고친다", {"size": 17, "color": BODY})]])
    w, gap = grid(3, 0.22)
    for i, (h1, h2) in enumerate([("빈 폴더에서", "오늘 새 프로젝트를 시작한다"),
                                  ("한 번에 만들고", "말로 시키면 화면이 나온다"),
                                  (fix_n + " 고친다", "바뀐 줄을 읽고 내가 승인한다")]):
        minicard(s, ML + i * (w + gap), 5.20, w, 1.20, h1, h2, bar=ACCENT if i == 2 else ACCENT_L)
    footer_rule(s)
    text(s, ML, FOOT_Y + 0.12, 8, 0.26,
         [[("오늘은 '만드는 법'보다 '고치는 법'을 배웁니다", {"size": 9.5, "color": FAINT, "space": 0.8})]])
    note(n(), [
        "● 여는 말",
        "오늘부터 도구가 바뀝니다. 대화창이 아니라 Claude Code에서, 내 노트북의 폴더에 직접 만듭니다.",
        "지난 시간까지 만든 파일은 쓰지 않습니다. 빈 폴더에서 새로 시작합니다.",
        "● 강조",
        "오늘 배우는 건 '만드는 법'이 아니라 '나온 결과를 고치는 법'입니다.",
    ])

    # ---------------------------------------------------------- 2. 오늘부터 새로 시작
    s = S()
    header(s)
    title_block(s, "오늘부터 새로 시작합니다", "지금까지는 기본기, 오늘부터는 도구")
    w, gap = grid(2, 0.30)
    wcard(s, ML, TOP, w, 3.00, LINE2, "1 ~ 3차시",
          "바이브코딩의 기본기",
          ["말로 시켜서 결과를 얻는 법을 익혔다",
           "무엇을 만들지 먼저 정하는 법을 익혔다",
           "결과가 기준에 맞는지 확인하는 법을 익혔다"],
          note="그때 만든 파일은 오늘 쓰지 않습니다.")
    wcard(s, ML + w + gap, TOP, w, 3.00, ACCENT, "4차시부터",
          "Claude Code로 진짜 만들기",
          ["빈 폴더에서 새 프로젝트를 시작한다",
           "AI가 내 노트북의 파일을 직접 만들고 고친다",
           "회기마다 같은 프로젝트를 이어서 키운다"],
          strong=True)
    banner(s, [[("오늘 만든 폴더를 5차시, 6차시에 그대로 이어서 씁니다. ", {}),
                ("오늘이 우리 프로젝트의 1일 차입니다.", {"bold": True, "color": ACCENT_D})]],
           y=5.66, h=0.80)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 이렇게 말해 보세요",
        "지난 세 번은 기본기를 익히는 시간이었어요. 말로 시키는 법, 먼저 정하는 법, 확인하는 법.",
        "오늘부터는 도구를 바꿔서 진짜 프로젝트를 시작합니다. 지난 파일은 안 써요. 새로 시작해요.",
        "● 중요",
        "오늘 만든 폴더는 앞으로 계속 씁니다. 백업을 꼭 안내하세요.",
    ])

    # ---------------------------------------------------------- 3. 왜 도구를 바꾸나
    s = S()
    header(s)
    title_block(s, "왜 도구를 바꿀까요", "대화창에서 할 수 있는 일과, 할 수 없는 일")
    w, gap = grid(2, 0.30)
    wcard(s, ML, TOP, w, 3.05, LINE2, "지금까지 · 대화창",
          "복사해서 붙여 넣기",
          ["코드를 복사해 메모장에 붙여 넣고 저장했다",
           "고칠 때마다 이 과정을 다시 반복했다",
           "파일을 여러 개로 나누기 어려웠다",
           "무엇을 언제 바꿨는지 기록이 없다"])
    wcard(s, ML + w + gap, TOP, w, 3.05, ACCENT, "오늘부터 · CLAUDE CODE",
          "내 폴더를 직접 만든다",
          ["AI가 내 노트북의 파일을 직접 만들고 고친다",
           "복사·붙여넣기가 없다",
           "파일이 여러 개여도 함께 다룬다",
           "무엇이 몇 줄 바뀌었는지 보여 준다"],
          strong=True)
    banner(s, [[("대화창에서는 마음에 안 들면 안 쓰면 그만이었습니다. 이제는 승인하는 순간 내 파일이 바뀝니다.", {})],
               [("그래서 오늘 배울 세 가지 — ", {}),
                ("계획을 본다 · 바뀐 줄을 본다 · 내가 승인한다", {"bold": True, "color": ACCENT_D})]],
           y=5.72, h=0.95)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 이렇게 말해 보세요",
        "대화창은 편했지만 고칠 때마다 복사하고 붙여 넣었죠. 그게 사라집니다.",
        "대신 책임이 하나 생겨요. 승인하는 순간 내 파일이 진짜로 바뀝니다.",
    ])

    # ---------------------------------------------------------- 4. PART 1
    s = S()
    if hw:
        part_slide(s, "P A R T   1", "출발선 맞추기", "앱 점검 · 빈 폴더 열기 — 8분", 1, TOTAL_PARTS)
        note(n(), [
            "● 진행 안내",
            "설치는 숙제였습니다. 이 블록은 설치하는 시간이 아니라 출발선을 맞추는 시간입니다.",
            "확인 사진을 올린 학생을 먼저 도우미로 지정하고, 안 된 학생에게 붙여 주세요.",
        ])
    else:
        part_slide(s, "P A R T   1", "설치하고 폴더 열기", "설치 · 로그인 · 빈 폴더 — 20분", 1, TOTAL_PARTS)
        note(n(), [
            "● 진행 안내",
            "설치는 '같이, 한 단계씩' 입니다. 각자 진도를 내게 두면 절반이 길을 잃습니다.",
            "단계마다 손을 들게 하고, 먼저 끝낸 학생은 즉시 도우미로 돌리세요.",
        ])

    # ---------------------------------------------------------- 5. 출발선 / 설치
    s = S()
    header(s)
    if hw:
        title_block(s, "출발선 — 네 가지만 확인", "설치해 온 사람은 확인만, 안 된 사람은 손을 들어 주세요")
        step_row(s, TOP, [
            ("확인 1", "로그인", "앱을 열면 내 이름이 보인다"),
            ("확인 2", "Code 화면", "Local / Cloud 선택이 보인다"),
            ("확인 3", "빈 폴더", "바탕화면에 만든 myapp 폴더를 열었다"),
            ("확인 4", "첫 대화", "\"이 폴더는 지금 비어 있나요?\" 에 답이 온다"),
        ], h=1.60)
        w, gap = grid(2, 0.30)
        y = 4.18
        lcard(s, ML, y, w, 2.45, ACCENT_L, "막히면 이렇게",
              ["수업에서 쓰던 그 계정으로 로그인한다",
               "혼자 붙잡고 있지 말고 짝 노트북에 합류한다",
               "학교 와이파이가 안 되면 휴대폰 핫스팟"],
              note="macOS는 처음 실행할 때 권한 허용 창이 한 번 더 뜹니다.")
        lcard(s, ML + w + gap, y, w, 2.45, GREEN, "먼저 끝났다면",
              ["도우미가 되어 준다",
               "대신 해 주지 말고 '어디까지 됐어?'부터 묻는다",
               "두 명까지만 돕고 자기 자리로 돌아온다"],
              label_color=GREEN_D)
        note(n(), [
            "● 운영",
            "확인 네 가지를 화면에 띄우고 손 든 학생 수를 세며 넘어갑니다.",
            "14분 체크포인트 — 3분의 1이 안 되면 뒤의 거절 연습을 버리고 수정 실습에 시간을 더 줍니다.",
        ])
    else:
        title_block(s, "설치는 네 단계입니다", "한 단계씩 같이 갑니다. 다 되면 손을 들어 주세요")
        step_row(s, TOP, [
            ("STEP 1", "설치 파일 받기", "USB 또는 공유 폴더에서 내 것(Windows · macOS)"),
            ("STEP 2", "받는 동안", "claude.ai 로그인 · 바탕화면에 myapp 폴더 만들기"),
            ("STEP 3", "설치 · 로그인", "앱을 실행하고 수업에서 쓰던 계정으로 로그인"),
            ("STEP 4", "폴더 열기", "Code → Local → Select folder → myapp"),
        ], h=1.70)
        w, gap = grid(2, 0.30)
        y = 4.24
        lcard(s, ML, y, w, 2.52, ACCENT_L, "막히면 이렇게",
              ["5분 안에 안 되면 짝 노트북에 합류한다",
               "학교 와이파이가 안 되면 휴대폰 핫스팟",
               "macOS는 '열기'와 권한 허용을 한 번 더 눌러야 한다"],
              note="한 명 때문에 반 전체가 기다리지 않게 합니다.")
        lcard(s, ML + w + gap, y, w, 2.52, GREEN, "먼저 끝났다면",
              ["바로 도우미가 되어 준다",
               "대신 해 주지 말고 '어디까지 됐어?'부터 묻는다",
               "두 명까지만 돕고 자기 자리로 돌아온다"],
              label_color=GREEN_D)
        note(n(), [
            "● 운영",
            "25명이 동시에 내려받으면 학교 네트워크가 먼저 멈춥니다. USB를 먼저 돌리고 나머지를 분산하세요.",
            "STEP 2가 핵심입니다. 설치를 기다리는 동안 로그인과 폴더 만들기를 먼저 시켜 빈 시간을 없앱니다.",
            "18분 체크포인트 — 3분의 1이 안 되면 거절 연습을 버리고 이 블록에 시간을 더 씁니다.",
        ])
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)

    # ---------------------------------------------------------- 6. 화면 구분
    s = S()
    header(s)
    title_block(s, "제대로 왔는지 확인하는 법", "화면이 두 가지라 헷갈립니다. 한 군데만 보면 됩니다")
    w, gap = grid(2, 0.30)
    wcard(s, ML, TOP, w, 3.00, ACCENT, "우리가 쓸 화면",
          "데스크톱 앱 · Local",
          ["환경 선택에 Local / Cloud 가 보인다",
           "Local을 고르면 Select folder 버튼이 있다",
           "내 노트북의 파일을 그대로 만든다",
           "GitHub 계정이 필요 없다"],
          strong=True)
    wcard(s, ML + w + gap, TOP, w, 3.00, LINE2, "오늘은 쓰지 않는 화면",
          "브라우저 claude.ai/code",
          ["주소창에 claude.ai 가 보인다",
           "GitHub 저장소를 고르라고 한다",
           "내 폴더를 고를 수 없다"],
          note="저장소와 배포는 다음 회기에 배웁니다.")
    banner(s, [[("화면에 ", {}), ("Select folder", {"bold": True, "color": ACCENT_D, "font": MONO}),
                (" 버튼이 있으면 데스크톱 앱, GitHub 저장소를 고르라고 하면 브라우저입니다.", {})]],
           y=5.66, h=0.80)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 이렇게 말해 보세요",
        "화면이 두 가지예요. 우리가 쓸 건 Select folder가 있는 쪽입니다.",
        "브라우저 쪽은 GitHub 저장소를 묻는데, 그건 다음 시간에 배울 내용이에요.",
    ])

    # ---------------------------------------------------------- 7. 오늘 만들 것
    s = S()
    header(s)
    title_block(s, "오늘 만들 것", "작게 시작합니다. 오늘의 목적은 '고치는 연습'입니다")
    w, gap = grid(2, 0.34)
    wcard(s, ML, TOP, w, 3.00, ACCENT, "공통 과제",
          "나를 소개하는 한 장 페이지",
          ["이름과 한 줄 소개",
           "좋아하는 것 세 가지",
           "버튼 하나 — 누르면 무언가 바뀐다"],
          note="내용은 각자 다르게, 구조는 같게. 그래야 서로 도울 수 있습니다.", strong=True)
    lcard(s, ML + w + gap, TOP, w, 3.00, BLUE, "다른 걸 만들고 싶다면",
          ["간단 점수판 — 올리기 · 내리기 · 처음으로",
           "타이머 — 시작 · 멈춤 · 다시",
           "오늘의 메모 — 쓰고 지우기"],
          note="무엇을 고르든 '한 파일, 버튼 하나'를 지킵니다.", label_color=BLUE_D)
    banner(s, [[("오늘은 크게 만들지 않습니다. ", {}),
                ("작게 만들고 " + fix_n + " 고치는 것", {"bold": True, "color": ACCENT_D}),
                ("이 목표입니다.", {})]],
           y=5.62, h=0.80)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 이렇게 말해 보세요",
        "오늘은 큰 걸 만들지 않아요. 작게 만들고, 세 번 고치는 게 목표예요.",
        "고치는 연습이 오늘의 진짜 수업이니까요.",
        "● 운영",
        "고민하는 학생이 있으면 공통 과제로 바로 보내세요. 주제 고르기에 5분을 쓰면 안 됩니다.",
    ])

    # ---------------------------------------------------------- 8. PART 2
    s = S()
    part_slide(s, "P A R T   2", "만들기", "첫 요청 한 번으로 화면이 나옵니다", 2, TOTAL_PARTS)
    note(n(), [
        "● 진행 안내",
        "여기서 첫 성공을 빠르게 만들어 줍니다. 화면이 떠야 뒤의 수정 연습이 살아납니다.",
    ])

    # ---------------------------------------------------------- 9. 화면 네 곳
    s = S()
    header(s)
    title_block(s, "화면에서 볼 곳은 네 군데", "나머지는 오늘 몰라도 됩니다")
    step_row(s, TOP, [
        ("①", "열린 폴더", "여기가 작업 대상. 다른 폴더는 건드리지 않는다"),
        ("②", "모드", "계획(Plan)으로 시작. 승인 전에는 만들지 않는다"),
        ("③", "바뀐 줄", "초록은 새로 생긴 줄, 빨강은 지워진 줄"),
        ("④", "승인 / 거절", "마음에 들지 않으면 거절할 수 있다"),
    ], h=1.60)
    w, gap = grid(2, 0.30)
    y = 4.18
    lcard(s, ML, y, w, 2.45, ACCENT, "계획 모드를 쓰는 이유",
          ["무엇을 어떻게 만들지 먼저 말하게 한다",
           "내가 승인하기 전에는 파일을 건드리지 않는다"],
          note="계획을 읽지 않고 승인하는 것은 읽지 않고 서명하는 것과 같습니다.", strong=True)
    lcard(s, ML + w + gap, y, w, 2.45, BLUE, "계획에서 볼 것 세 가지",
          ["내가 말하지 않은 것이 들어 있지 않은가",
           "만들겠다는 파일이 하나가 맞는가",
           "한 번에 너무 많이 하려 들지 않는가"],
          label_color=BLUE_D)
    footer_rule(s)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 시연",
        "네 곳을 마우스로 하나씩 짚어 주세요. 말보다 화면이 빠릅니다.",
        "● 강조",
        "계획을 읽지 않고 승인하는 학생이 가장 많습니다. '내가 말 안 한 게 들어 있나?'를 반복해 물어 주세요.",
    ])

    # ---------------------------------------------------------- 10. 첫 요청문
    s = S()
    header(s)
    title_block(s, "첫 요청문 — 한 번에 작게", "계획 모드로 입력하고, 계획을 읽은 뒤 승인합니다")
    codebox(s, ML, TOP, CW, 2.55, "그대로 복사해서 쓰세요",
            ["나를 소개하는 한 장짜리 웹페이지를 index.html 로 만들어 주세요.",
             "- 이름, 한 줄 소개, 좋아하는 것 세 가지가 보입니다.",
             "- 버튼 하나를 넣고, 누르면 응원 문구가 하나 바뀌게 해 주세요.",
             "- 한 개의 HTML 파일로만 만들고, 외부 라이브러리나 인터넷 연결은 쓰지 마세요.",
             "- 휴대폰 화면에서도 보기 좋아야 합니다.",
             "먼저 계획을 알려 주고, 제가 승인하면 만들어 주세요."])
    step_row(s, 5.16, [
        ("①", "계획을 읽는다", "무엇을 만들겠다는지 소리 내어 읽기"),
        ("②", "승인한다", "이상하면 승인하지 않고 다시 요청"),
        ("③", "브라우저로 연다", "폴더의 index.html 을 더블클릭해서 확인"),
    ], h=1.45)
    footer_rule(s)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 운영",
        "요청문은 게시판에 올려 두고 복사해 쓰게 합니다. 타이핑에 시간을 쓰지 않게 합니다.",
        "다른 주제를 고른 학생은 첫 줄만 바꾸면 됩니다.",
        "● 중요",
        "화면이 떴으면 성공입니다. 예쁘지 않아도 괜찮다고 먼저 말해 주세요. 고치는 건 다음 블록입니다.",
    ])

    # ---------------------------------------------------------- 11. 결과 확인하는 법
    s = S()
    header(s)
    title_block(s, "결과는 이렇게 확인합니다", "'나왔다'가 아니라 '무엇이 생겼나'를 봅니다")
    w, gap = grid(3, 0.26)
    for i, (bar, lc, label, head, lines) in enumerate([
        (ACCENT, ACCENT_D, "① 파일", "무엇이 생겼나",
         ["내 폴더에 index.html 이 생겼는지 본다", "말하지 않은 파일이 같이 생기지 않았는지 본다"]),
        (BLUE, BLUE_D, "② 화면", "실제로 열어 본다",
         ["index.html 을 더블클릭해 브라우저로 연다", "고친 뒤에는 Ctrl+F5 로 새로고침한다"]),
        (GREEN, GREEN_D, "③ 동작", "직접 눌러 본다",
         ["버튼을 눌러 본다", "휴대폰 크기로 창을 좁혀 본다"]),
    ]):
        wcard(s, ML + i * (w + gap), TOP, w, 2.80, bar, label, head, lines, label_color=lc, size=11)
    banner(s, [[("보기만 하지 말고 ", {}), ("직접 눌러 보세요.", {"bold": True, "color": ACCENT_D}),
                (" 눌러 보기 전까지는 됐는지 알 수 없습니다.", {})]],
           y=5.46, h=0.80)
    footer_rule(s)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 운영",
        "첫 결과가 나오면 바로 다음 요청으로 넘어가려는 학생이 많습니다. 꼭 눌러 보게 하세요.",
        "여기서 발견한 것이 뒤의 수정 거리가 됩니다.",
    ])

    # ---------------------------------------------------------- 12. PART 3
    s = S()
    part_slide(s, "P A R T   3", "고치기", "말로 시키면 뚝딱 나옵니다. 문제는 그다음입니다", 3, TOTAL_PARTS)
    note(n(), [
        "● 오늘의 핵심 블록",
        "여기서부터가 4차시의 진짜 내용입니다.",
        "결과를 받는 법이 아니라, 받은 결과를 고치는 법입니다.",
    ])

    # ---------------------------------------------------------- 13. 뚝딱 나왔는데
    s = S()
    header(s)
    title_block(s, "말로 시키면 뚝딱 나옵니다", "그런데 그다음에 할 일이 진짜 실력입니다")
    w, gap = grid(2, 0.30)
    wcard(s, ML, TOP, w, 2.90, LINE2, "흔한 모습",
          "나왔으니까 됐다",
          ["화면이 뜬 것만 보고 넘어간다",
           "어디가 어떻게 바뀌었는지 모른다",
           "문제가 생기면 처음부터 다시 만든다"])
    wcard(s, ML + w + gap, TOP, w, 2.90, ACCENT, "우리가 할 일",
          "읽고, 고르고, 다시 시킨다",
          ["무엇이 바뀌었는지 줄 단위로 읽는다",
           "내가 원한 것과 맞는지 판단한다",
           "틀렸으면 무엇이 다른지 말해서 다시 요청한다"], strong=True)
    banner(s, [[("AI가 코드를 만들어 줄수록, 사람이 할 일은 ", {}),
                ("판단하고 설명하는 쪽", {"bold": True, "color": ACCENT_D}), ("으로 옮겨갑니다.", {})]],
           y=5.46, h=0.80)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 이렇게 말해 보세요",
        "말로 시키면 뚝딱 나오죠. 그런데 그게 내가 원한 게 맞는지는 누가 판단하죠? 여러분이에요.",
        "오늘 " + fix_n + " 고칠 거예요. 고치는 말이 정확할수록 적게 고쳐도 됩니다.",
    ])

    # ---------------------------------------------------------- 14. 수정 요청 네 칸
    s = S()
    header(s)
    title_block(s, "수정 요청은 네 칸으로", "이 형식이 오늘 가장 중요합니다")
    step_row(s, TOP, [
        ("문제", "무슨 일이 일어났나", "본 사실을 적는다. '이상하다'가 아니라 '버튼이 안 눌린다'"),
        ("근거", "원래 무엇을 원했나", "내가 요청했던 내용을 그대로 다시 적는다"),
        ("결과", "어떻게 되어야 하나", "고친 뒤 기대하는 동작을 적는다"),
        ("범위", "어디까지 고치나", "바꿀 부분과 바꾸지 말 것을 함께 적는다"),
    ], h=1.80)
    codebox(s, ML, 4.38, CW, 2.10, "예시 — 계획 모드로 입력",
            ["문제: 버튼을 눌러도 응원 문구가 바뀌지 않습니다.",
             "근거: 처음 요청에서 '버튼을 누르면 문구가 하나 바뀐다'고 했습니다.",
             "원하는 결과: 누를 때마다 준비된 문구 중 하나가 번갈아 보이게 해 주세요.",
             "수정 범위: 버튼 동작만 고치고, 화면 배치와 색은 바꾸지 마세요."])
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 강조",
        "'범위'가 가장 중요합니다. 범위를 안 적으면 한 군데 고치다 멀쩡한 곳까지 바뀝니다.",
        "● 순회 지도",
        "'더 예쁘게 해 줘' 같은 막연한 요청을 하는 학생에게 네 칸을 채워 다시 쓰게 하세요.",
    ])

    # ---------------------------------------------------------- 15. 수정 유형 ①②③
    s = S()
    header(s)
    title_block(s, "자주 쓰게 될 수정 ① ② ③", "말을 바꾸면 결과가 달라집니다")
    w, gap = grid(3, 0.26)
    for i, (bar, lc, label, head, lines) in enumerate([
        (ACCENT, ACCENT_D, "① 화면 배치", "어디에 무엇이 보이나",
         ["✗ 더 예쁘게 해 줘", "✓ 이름을 맨 위 가운데로 옮기고, 글씨를 더 크게 해 주세요."]),
        (BLUE, BLUE_D, "② 문구와 색", "눈에 보이는 것",
         ["✗ 색 좀 바꿔 줘", "✓ 배경을 연한 회색으로, 버튼은 남색으로 바꿔 주세요."]),
        (GREEN, GREEN_D, "③ 동작", "눌렀을 때 일어나는 일",
         ["✗ 버튼 잘 되게 해 줘", "✓ 버튼을 누를 때마다 문구가 순서대로 바뀌게 해 주세요."]),
    ]):
        wcard(s, ML + i * (w + gap), TOP, w, 2.95, bar, label, head, lines, label_color=lc, size=11)
    banner(s, [[("공통점 — ", {}),
                ("무엇을 보았는지(사실)와 어떻게 되길 바라는지(결과)를 나눠서 말한다.",
                 {"bold": True, "color": ACCENT_D})]],
           y=5.62, h=0.80)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 운영",
        "세 가지를 다 가르치려 하지 말고, 학생이 막힌 유형을 짚어 이 장으로 돌아오세요.",
        "✗와 ✓를 소리 내어 읽게 하면 차이를 금방 느낍니다.",
    ])

    # ---------------------------------------------------------- 16. 수정 유형 ④⑤
    s = S()
    header(s)
    title_block(s, "자주 쓰게 될 수정 ④ ⑤", "안 될 때와, 되돌리고 싶을 때")
    w, gap = grid(2, 0.30)
    wcard(s, ML, TOP, w, 2.95, PURPLE, "④ 오류", "안 되는 것을 고칠 때",
          ["✗ 안 돼요. 고쳐 주세요",
           "✓ 버튼을 눌러도 아무 일이 일어나지 않습니다",
           "✓ 화면에 빨간 글씨는 없습니다"],
          note="무엇을 눌렀고 무엇이 안 일어났는지를 말하면 훨씬 빨리 찾습니다.",
          label_color=PURPLE_D)
    wcard(s, ML + w + gap, TOP, w, 2.95, ACCENT_L, "⑤ 되돌리기", "마음에 들지 않을 때",
          ["✓ 승인하기 전이라면 그냥 거절하면 된다",
           "✓ 이미 바뀌었다면 \"방금 바꾼 것을 되돌려 주세요\"",
           "✓ 마음에 든 상태는 폴더째 복사해 두면 안전하다"],
          note="되돌리기는 다음 회기에 git으로 제대로 배웁니다.")
    banner(s, [[("거절은 실패가 아닙니다. ", {}),
                ("도구를 다루는 사람이 하는 일입니다.", {"bold": True, "color": ACCENT_D})]],
           y=5.62, h=0.80)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 강조",
        "오류를 말할 때는 '안 돼요'가 아니라 '무엇을 눌렀더니 무엇이 안 일어났다'로 말하게 합니다.",
        "되돌리기를 한 번 해 보면 학생들이 훨씬 과감하게 시도합니다.",
    ])

    # ---------------------------------------------------------- 17. 이상할 때 순서
    s = S()
    header(s)
    title_block(s, "결과가 이상할 때, 이 순서로", "당황하지 말고 위에서부터 확인합니다")
    step_row(s, TOP, [
        ("1단계", "새로고침했나", "Ctrl+F5. 고쳤는데 화면이 그대로인 경우가 가장 많다"),
        ("2단계", "무엇이 바뀌었나", "바뀐 줄을 본다. 내가 말한 파일이 맞는지 확인"),
        ("3단계", "원한 것과 맞나", "직접 눌러 보고 처음 요청과 비교한다"),
        ("4단계", "다시 요청", "네 칸(문제·근거·결과·범위)으로 한 번 더"),
    ], h=1.60)
    w, gap = grid(2, 0.30)
    y = 4.18
    lcard(s, ML, y, w, 2.60, ACCENT, "이럴 때는 멈추세요",
          ["말하지 않은 파일까지 만들거나 고치려 한다",
           "요청하지 않은 기능이 계획에 들어 있다",
           "한 번에 여러 가지를 바꾸려 한다"], strong=True)
    lcard(s, ML + w + gap, y, w, 2.60, ACCENT_L, "거절할 때 하는 말",
          ["\"로그인은 요청하지 않았습니다. 추가하지 마세요.\"",
           "\"index.html 만 고쳐 주세요. 다른 파일은 만들지 마세요.\"",
           "\"한 번에 하나만 고쳐 주세요. 먼저 ○○만.\""], size=11)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 자주 나는 상황",
        "'고쳤다는데 화면이 그대로예요' — 새로고침(Ctrl+F5)을 안 한 경우가 대부분입니다.",
        "AI가 엉뚱한 파일을 만들려 하면 반 전체에 보여 주세요. 오늘 최고의 교재입니다.",
    ])

    # ---------------------------------------------------------- 18. CLAUDE.md
    s = S()
    header(s)
    if hw:
        title_block(s, "마지막 — 오늘 정한 것을 파일로 남깁니다", "다음 시간에 AI가 처음부터 다시 묻지 않도록")
    else:
        title_block(s, "시간이 남으면 — 규칙을 파일로", "여기까지 못 해도 괜찮습니다. 다음 시간 첫 블록에서 함께 만듭니다")
    w, gap = grid(2, 0.30)
    lcard(s, ML, TOP, w, 2.05, LINE2, "없으면",
          ["다음 시간에 AI는 우리 프로젝트를 모른다",
           "같은 설명을 매번 다시 해야 한다"])
    lcard(s, ML + w + gap, TOP, w, 2.05, ACCENT, "있으면",
          ["폴더를 열 때마다 자동으로 읽는다",
           "하지 말 것을 미리 막아 두고, 요청문이 짧아진다"], strong=True)
    codebox(s, ML, 4.58, CW, 2.22, "계획 모드로 입력",
            ["오늘 작업한 내용을 바탕으로 CLAUDE.md 를 만들어 주세요.",
             "- 이 프로젝트는 한 개의 index.html 로 된 웹페이지다.",
             "- 외부 라이브러리, 서버, 로그인은 쓰지 않는다.",
             "- 고치기 전에 계획을 말하고, 고친 뒤 바꾼 내용을 3줄로 알려 준다.",
             "CLAUDE.md 파일만 만들고 다른 파일은 바꾸지 마세요."])
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 이렇게 말해 보세요",
        "새 대화를 열면 AI는 우리 프로젝트를 몰라요. 2차시에 배운 그 문제죠.",
        "그래서 오늘 정한 규칙을 파일로 남깁니다. 다음 시간에 AI가 먼저 읽고 시작해요.",
        "● 운영",
        ("만들어진 CLAUDE.md 를 꼭 직접 열어서 읽게 하세요."
         if hw else
         "설치에 시간을 쓴 날이라 여기까지 못 가는 반이 많습니다. 빨리 끝낸 학생의 심화 과제로 주고, 5차시 첫 블록에서 함께 만드세요."),
    ])

    # ---------------------------------------------------------- 19. 실습 순서
    s = S()
    header(s)
    title_block(s, "오늘 실습 순서", "여기까지 하면 오늘은 성공입니다")
    if hw:
        items = [
            ("실습 ①   8분", "출발선 맞추기", "앱 점검 → 빈 폴더 myapp 열기 → 첫 대화"),
            ("실습 ②  14분", "한 번에 만들기", "첫 요청문 → 계획 읽기 → 승인 → 브라우저로 열어 보기"),
            ("실습 ③  24분", "세 번 고치기", "배치 한 번 · 색이나 문구 한 번 · 동작 한 번"),
            ("실습 ④   8분", "일부러 거절해 보기", "\"로그인 기능도 추가해 줘\" → 계획을 읽고 거절하기"),
            ("실습 ⑤   8분", "규칙 남기기", "CLAUDE.md 만들고 폴더 백업하기"),
        ]
    else:
        items = [
            ("실습 ①  20분", "설치 · 로그인 · 폴더", "설치 파일 → 로그인 → myapp 열기 → 첫 대화"),
            ("실습 ②  14분", "한 번에 만들기", "첫 요청문 → 계획 읽기 → 승인 → 브라우저로 열어 보기"),
            ("실습 ③  20분", "두 번 고치기", "눈에 보이는 것 한 번 · 동작 한 번"),
            ("실습 ④   8분", "일부러 거절해 보기", "\"로그인 기능도 추가해 줘\" → 계획을 읽고 거절하기"),
            ("실습 ⑤   4분", "폴더 백업", "myapp 폴더를 USB나 메일로 보내기"),
        ]
    for i, (num, head, desc) in enumerate(items):
        yy = TOP + i * 0.78
        rect(s, ML, yy, CW, 0.68, fill=CARD, line=LINE, radius=0.05)
        rect(s, ML, yy, 0.055, 0.68, fill=ACCENT if i < 3 else ACCENT_L)
        text(s, ML + 0.30, yy + 0.12, 1.9, 0.3,
             [[(num, {"bold": True, "size": 11, "color": ACCENT_D, "space": 0.6})]])
        text(s, ML + 2.30, yy + 0.10, 4.3, 0.34, [[(head, {"bold": True, "size": 14, "color": INK})]])
        text(s, ML + 6.70, yy + 0.14, CW - 6.95, 0.44, [[(desc, {"size": 11, "color": MUTED})]], line=1.35)
    banner(s, [[("처음이라면 ", {}), ("실습 ③의 첫 번째 수정까지만 해도 성공", {"bold": True, "color": ACCENT_D}),
                ("입니다. 빨리 끝낸 사람은 한 번 더 고쳐 보세요.", {})]], y=6.26, h=0.62)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 운영",
        "타이머를 띄워 두고 블록마다 시작·끝을 분명히 알립니다.",
        "실습 ④(거절 연습)는 시간이 없으면 다음 시간 도입부로 옮겨도 됩니다.",
    ])

    # ---------------------------------------------------------- 20. 마무리
    s = S()
    header(s)
    title_block(s, "오늘 기억할 세 단어", "오늘이 우리 프로젝트의 1일 차입니다")
    w, gap = grid(3, 0.26)
    for i, (h1, h2) in enumerate([("계획", "만들기 전에 무엇을 할지 먼저 듣는다"),
                                  ("바뀐 줄", "무엇이 어떻게 바뀌었는지 눈으로 읽는다"),
                                  ("승인", "맞다고 판단했을 때만 내가 승인한다")]):
        x = ML + i * (w + gap)
        rect(s, x, TOP, w, 1.55, fill=CARD, line=ACCENT_L, lw=1.4, radius=0.06)
        rect(s, x, TOP, w, 0.075, fill=ACCENT)
        text(s, x + 0.26, TOP + 0.36, w - 0.52, 0.5, [[(h1, {"bold": True, "size": 23, "color": INK})]])
        text(s, x + 0.26, TOP + 0.92, w - 0.52, 0.5, [[(h2, {"size": 11.5, "color": MUTED})]], line=1.4)
    w2, gap2 = grid(2, 0.30)
    y = 4.20
    left_lines = ["내 노트북의 myapp 폴더",
                  "index.html — 오늘 만들고 " + fix_n + " 고친 내 페이지"]
    left_lines.append("CLAUDE.md — 우리 프로젝트 규칙" if hw else "(빨리 끝낸 사람) CLAUDE.md — 프로젝트 규칙")
    lcard(s, ML, y, w2, 2.55, GREEN, "오늘 남은 것", left_lines,
          note="폴더를 USB나 메일로 꼭 백업하세요. 다음 시간에 이어서 씁니다.", label_color=GREEN_D)
    next_lines = ["오늘 만든 것을 인터넷에 올린다",
                  "주소가 생기고, 휴대폰으로 열어 본다"]
    if not hw:
        next_lines.insert(0, "CLAUDE.md 로 프로젝트 규칙 남기기")
    lcard(s, ML + w2 + gap2, y, w2, 2.55, ACCENT, "다음 시간", next_lines,
          note="노트북과 충전기를 꼭 가져오세요.", strong=True)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 마무리",
        "두세 명에게 30초씩 — 오늘 고친 것 한 가지와 거절한 것 한 가지.",
        "폴더 백업을 꼭 확인하세요. 다음 시간 진행이 여기에 달려 있습니다.",
    ])

    # ---------------------------------------------------------- 노트 붙이기
    for idx, sl in enumerate(prs.slides, start=1):
        if idx in NOTES:
            sl.notes_slide.notes_text_frame.text = NOTES[idx]

    prs.save(os.path.join(ROOT, deck_name))
    return prs


if __name__ == "__main__":
    for key in ("A", "B"):
        cfg = VARIANTS[key]
        HOMEWORK = cfg["homework"]
        globals()["HOMEWORK"] = HOMEWORK
        B.CUR["n"] = 0
        del B.WARN[:]
        NOTES.clear()
        prs = build(cfg["deck"])
        print("")
        print("[%s] %s" % (key, cfg["deck"]))
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
