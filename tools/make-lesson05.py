# -*- coding: utf-8 -*-
"""
5차시 PPT 생성기 — 「인터랙티브 웹 프로젝트 2 : 개선하고 배포하기」

실행:  python tools/make-lesson05.py
결과:  디미고_바이브코딩_5차시_개선과배포.pptx   (프로젝트 최상위, 새 파일)

- 4차시 덱(인터랙티브 웹 프로젝트)의 디자인·어조를 그대로 잇습니다.
- 카드·배너·코드상자 부품은 3·4차시 생성기의 것을 재사용합니다.
- 기존 PPT는 열지도 수정하지도 않습니다. 새 파일만 만듭니다.
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

DECK = "디미고_바이브코딩_5차시_개선과배포.pptx"

ML, CW, SW, SH, TOP, FOOT_Y = B.ML, B.CW, B.SW, B.SH, B.TOP, B.FOOT_Y
INK, BODY, MUTED, FAINT = B.INK, B.BODY, B.MUTED, B.FAINT
ACCENT, ACCENT_D, ACCENT_L = B.ACCENT, B.ACCENT_D, B.ACCENT_L
LINE, LINE2, CARD, SOFT, TINT = B.LINE, B.LINE2, B.CARD, B.SOFT, B.TINT
DARK, WHITE, BLUE, GREEN, PURPLE = B.DARK, B.WHITE, B.BLUE, B.GREEN, B.PURPLE
KR, MONO = B.KR, B.MONO
rect, text, header, footer_rule, title_block, banner = B.rect, B.text, B.header, B.footer_rule, B.title_block, B.banner
wcard, dcard, codebox, step_row = B.wcard, B.dcard, B.codebox, B.step_row
grid, lcard, minicard, checklist = L4.grid, L4.lcard, L4.minicard, L4.checklist
BLUE_D, GREEN_D, PURPLE_D, RED = L4.BLUE_D, L4.GREEN_D, L4.PURPLE_D, L4.RED

TOTAL_PARTS = 3
NOTES = {}


def note(i, lines):
    NOTES[i] = "\n".join(lines)


def table_rows(s, y, rows, col1=3.5, row_h=0.62, gap=0.08, head=None):
    """두 칸짜리 표 — 증상/조치 같은 대응표"""
    yy = y
    if head:
        text(s, ML + 0.26, yy, col1 - 0.3, 0.26,
             [[(head[0], {"bold": True, "size": 10, "color": ACCENT_D, "space": .8})]])
        text(s, ML + col1 + 0.26, yy, CW - col1 - 0.5, 0.26,
             [[(head[1], {"bold": True, "size": 10, "color": ACCENT_D, "space": .8})]])
        yy += 0.34
    for a, b in rows:
        rect(s, ML, yy, CW, row_h, fill=CARD, line=LINE, radius=0.05)
        rect(s, ML, yy, 0.05, row_h, fill=ACCENT_L)
        text(s, ML + 0.26, yy + row_h / 2 - 0.15, col1 - 0.4, 0.3,
             [[(a, {"bold": True, "size": 12.5, "color": INK})]], anchor=MSO_ANCHOR.MIDDLE)
        text(s, ML + col1 + 0.26, yy + row_h / 2 - 0.15, CW - col1 - 0.5, 0.3,
             [[(b, {"size": 11.5, "color": BODY})]], anchor=MSO_ANCHOR.MIDDLE)
        yy += row_h + gap
    return yy


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
         [[("한국디지털미디어고등학교 · 5차시", {"bold": True, "size": 10, "color": FAINT, "space": 1.0})]],
         align=PP_ALIGN.RIGHT)
    text(s, ML, 1.70, 2.2, 1.2, [[("05", {"bold": True, "size": 60, "color": ACCENT_L})]])
    text(s, ML, 2.84, CW, 1.0, [[("인터랙티브 웹 프로젝트 2", {"bold": True, "size": 40, "color": INK})]])
    rect(s, ML, 3.94, 0.9, 0.05, fill=ACCENT)
    text(s, ML, 4.20, CW * 0.9, 0.5,
         [[("구현을 완성하고, 웹에 배포해 다른 기기에서 실행하는 과정", {"size": 17, "color": BODY})]])
    w, gap = grid(3, 0.22)
    for i, (h1, h2) in enumerate([("개선", "테스트에서 발견한 문제를 수정한다"),
                                  ("기록", "프로젝트 규칙과 변경 내용을 남긴다"),
                                  ("배포", "공개 주소로 접속해 동작을 확인한다")]):
        minicard(s, ML + i * (w + gap), 5.20, w, 1.20, h1, h2, bar=ACCENT if i == 2 else ACCENT_L)
    footer_rule(s)
    text(s, ML, FOOT_Y + 0.12, 8, 0.26,
         [[("로컬 실행에서 공개 URL까지", {"size": 9.5, "color": FAINT, "space": 0.8})]])
    note(n(), [
        "● 여는 말 (1분)",
        "지난 시간에는 내 노트북에서만 실행되는 작품을 만들었습니다.",
        "오늘은 그 작품을 완성도 있게 다듬고, 인터넷 주소로 공개해 다른 기기에서 실행합니다.",
        "● 도착점",
        "수업이 끝나면 각자 공개 URL 하나와 그 주소로 접속한 기록이 남습니다.",
    ])

    # ---------------------------------------------------------- 2. 시작 조건과 도착점
    s = S()
    header(s)
    title_block(s, "시작 조건과 도착점", "지난 차시 결과물을 확인한 뒤 작업을 이어 갑니다")
    w, gap = grid(2, 0.30)
    wcard(s, ML, TOP, w, 3.00, LINE2, "시작 조건 · 4차시 결과물",
          "지금 가지고 있는 것",
          ["index.html — 핵심 기능이 작동하는 단일 파일",
           "커밋 기록 — 되돌릴 수 있는 버전",
           "개발일지 — 구현 내용과 남은 과제"],
          note="세 가지가 없으면 먼저 복구합니다. 백업 폴더나 예비 파일을 사용합니다.")
    wcard(s, ML + w + gap, TOP, w, 3.00, ACCENT, "도착점 · 5차시 결과물",
          "수업 후 남는 것",
          ["개선된 index.html — 사용성 문제 수정 반영",
           "CLAUDE.md — 프로젝트 규칙 문서",
           "공개 URL — 다른 기기에서 접속 가능"], strong=True)
    banner(s, [[("오늘의 기준 — ", {}),
                ("다른 사람이 자신의 휴대폰에서 주소만으로 실행할 수 있는 상태", {"bold": True, "color": ACCENT_D})]],
           y=5.62, h=0.78)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 확인 (2분)",
        "시작 조건 세 가지를 각자 확인하게 합니다. 손을 들어 확인하면 빠릅니다.",
        "파일이 없는 학생은 예비 파일로 진행하고, 백업 습관을 다시 안내합니다.",
    ])

    # ---------------------------------------------------------- 3. 작업 흐름
    s = S()
    header(s)
    title_block(s, "오늘의 작업 흐름", "개선과 배포를 분리해 각각 검증합니다")
    step_row(s, TOP, [
        ("①", "기능 개선", "테스트에서 발견한 문제를 수정하고 커밋"),
        ("②", "규칙 정리", "프로젝트 규칙을 CLAUDE.md로 고정"),
        ("③", "웹 배포", "저장소에 올리고 공개 주소 생성"),
        ("④", "배포 검증", "다른 기기와 브라우저에서 동작 확인"),
    ], h=1.70)
    w, gap = grid(2, 0.30)
    y = 4.28
    lcard(s, ML, y, w, 2.40, ACCENT, "시간 배분",
          ["개선 20분 · 규칙 정리 5분",
           "배포 20분 · 검증 10분",
           "교차 테스트 10분 · 기록 10분"], strong=True)
    lcard(s, ML + w + gap, y, w, 2.40, BLUE, "판단 기준",
          ["개선은 발견한 문제 중 하나만 선택한다",
           "배포 전에 로컬에서 정상 동작을 확인한다",
           "배포 후에는 반드시 다른 기기에서 확인한다"],
          label_color=BLUE_D)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 안내 (2분)",
        "개선과 배포를 섞지 않습니다. 개선이 끝나고 동작을 확인한 상태에서 배포로 넘어갑니다.",
        "배포가 막히는 학생이 생기므로, 개선 단계는 시간을 지켜 끊습니다.",
    ])

    # ---------------------------------------------------------- 4. 개선 대상 선정
    s = S()
    header(s)
    title_block(s, "개선 대상 선정", "발견한 문제를 분류한 뒤 하나를 선택합니다")
    w, gap = grid(3, 0.26)
    for i, (bar, lc, label, head, lines) in enumerate([
        (RED, RED, "지금 고친다", "실행을 막는 문제",
         ["핵심 기능이 작동하지 않음", "특정 입력에서 화면이 멈춤"]),
        (ACCENT, ACCENT_D, "다음에 고친다", "사용에 불편한 문제",
         ["조작 순서가 직관적이지 않음", "결과 표시가 늦거나 불명확함"]),
        (LINE2, MUTED, "이번에는 제외", "범위 밖 요청",
         ["새로운 기능 추가", "전체 디자인 교체"]),
    ]):
        wcard(s, ML + i * (w + gap), TOP, w, 2.55, bar, label, head, lines, label_color=lc, size=11)
    banner(s, [[("선택 기준 — ", {}),
                ("사용자가 목표를 끝까지 수행하지 못하게 막는 문제를 먼저 처리합니다.", {"bold": True, "color": ACCENT_D})]],
           y=5.20, h=0.80)
    footer_rule(s)
    B.page(s, 1, TOTAL_PARTS)
    note(n(), [
        "● 안내 (3분)",
        "지난 차시 교차 테스트에서 받은 피드백을 꺼내 분류하게 합니다.",
        "'새 기능 추가'는 오늘 범위가 아닙니다. 개발일지의 다음 작업으로 옮깁니다.",
    ])

    # ---------------------------------------------------------- 5. 미션 ① 기능 개선
    s = S()
    header(s)
    title_block(s, "미션 ① 기능 개선", "20분 · 선택한 문제 하나를 수정하고 결과를 커밋합니다")
    codebox(s, ML, TOP, CW, 2.35, "수정 요청문 · 프로젝트에 맞게 수정",
            ["발생 조건: [어떤 입력 또는 조작에서 발생하는지]",
             "현재 동작: [실제로 나타나는 결과]",
             "기대 동작: [수정 후 기대하는 결과]",
             "변경 범위: [수정할 영역]만 수정하고 다른 기능과 레이아웃은 유지해 줘.",
             "수정 전에 변경 계획을 먼저 알려 줘."])
    checklist(s, ML, 4.92, CW, [
        ("계획 검토", "변경 대상 파일과 수정 범위가 요청과 일치하는지 확인"),
        ("동작 확인", "브라우저에서 새로고침 후 기대 동작과 비교"),
        ("버전 기록", "정상 동작을 확인한 시점에서 커밋"),
    ], row_h=0.60, gap=0.08)
    footer_rule(s)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 진행 (20분)",
        "요청문은 게시판에서 복사해 사용하게 합니다. 네 항목을 모두 채우도록 안내합니다.",
        "'계획을 먼저 알려 줘'를 유지하게 합니다. 승인 전에는 파일이 변경되지 않습니다.",
        "● 순회 포인트",
        "막연한 요청을 하는 학생에게는 발생 조건과 현재 동작을 먼저 적게 합니다.",
        "커밋 없이 다음 수정으로 넘어가지 않도록 중간에 한 번 확인합니다.",
    ])

    # ---------------------------------------------------------- 6. CLAUDE.md
    s = S()
    header(s)
    title_block(s, "프로젝트 규칙 문서 · CLAUDE.md", "반복해서 전달하던 조건을 파일로 고정합니다")
    w, gap = grid(2, 0.30)
    lcard(s, ML, TOP, w, 1.95, LINE2, "문서가 없을 때",
          ["요청마다 실행 환경과 제약을 다시 설명",
           "대화가 바뀌면 기존 조건이 유지되지 않음"])
    lcard(s, ML + w + gap, TOP, w, 1.95, ACCENT, "문서가 있을 때",
          ["폴더를 열 때 규칙을 먼저 참조",
           "요청문이 짧아지고 결과의 일관성이 높아짐"], strong=True)
    codebox(s, ML, 4.40, CW, 2.45, "작성 요청문",
            ["이 폴더의 작업 내용을 바탕으로 CLAUDE.md 를 작성해 줘.",
             "- 단일 index.html 로 동작하는 웹 프로젝트이다.",
             "- 외부 라이브러리, 서버, 로그인, API 키는 사용하지 않는다.",
             "- 수정 전에 변경 계획을 제시하고, 승인 후 수정한다.",
             "- 모바일 화면에서도 핵심 기능을 사용할 수 있어야 한다.",
             "CLAUDE.md 파일만 생성하고 다른 파일은 변경하지 마."])
    footer_rule(s)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 진행 (5분)",
        "생성된 CLAUDE.md 를 직접 열어 내용을 확인하게 합니다.",
        "조별 작업이 시작되는 다음 차시부터 이 문서가 팀 규칙의 기준이 됩니다.",
    ])

    # ---------------------------------------------------------- 7. 배포의 의미
    s = S()
    header(s)
    title_block(s, "배포란 무엇인가", "실행 위치가 내 컴퓨터에서 웹 서버로 바뀝니다")
    w, gap = grid(2, 0.30)
    wcard(s, ML, TOP, w, 2.70, LINE2, "로컬 실행",
          "내 컴퓨터에서만 실행",
          ["파일을 가진 사람만 실행 가능",
           "다른 기기에서 확인하려면 파일 전달 필요",
           "휴대폰에서 바로 열기 어려움"])
    wcard(s, ML + w + gap, TOP, w, 2.70, ACCENT, "웹 배포",
          "주소로 누구나 실행",
          ["URL 하나로 접속하면 즉시 실행",
           "PC·휴대폰 등 기기 제약 없음",
           "수정한 내용을 다시 올려 갱신"], strong=True)
    banner(s, [[("배포 조건 — ", {}),
                ("단일 index.html · 외부 설치 없이 실행 · 상대 경로 사용", {"bold": True, "color": ACCENT_D}),
                (". 4차시 구현 원칙을 그대로 지키면 배포 가능합니다.", {})]],
           y=5.34, h=0.80)
    footer_rule(s)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 안내 (3분)",
        "배포는 새로운 기술이 아니라 '파일을 공개된 위치에 올리는 작업'이라고 설명합니다.",
        "4차시의 단일 파일 원칙을 지켰다면 추가 작업 없이 그대로 배포됩니다.",
    ])

    # ---------------------------------------------------------- 8. 공개 전 점검
    s = S()
    header(s)
    title_block(s, "공개 전 점검", "공개되는 순간 코드와 파일을 누구나 확인할 수 있습니다")
    checklist(s, ML, TOP, CW, [
        ("파일 이름", "첫 화면 파일 이름이 index.html 인지 확인"),
        ("개인정보", "이름·연락처·학번 등 공개하지 않을 정보 제거"),
        ("API 키", "키나 비밀번호를 코드에 포함하지 않음 — 공개 시 유출"),
        ("이미지·소리", "직접 만들었거나 사용 조건을 확인한 자료만 포함"),
        ("경로", "파일 경로를 상대 경로로 작성 — 대소문자 구분 확인"),
    ], row_h=0.62, gap=0.08)
    banner(s, [[("웹에 공개된 코드는 브라우저 개발자 도구로 전부 열람할 수 있습니다. ", {}),
                ("공개해도 되는 내용만 포함합니다.", {"bold": True, "color": ACCENT_D})]],
           y=6.00, h=0.72)
    footer_rule(s)
    B.page(s, 2, TOTAL_PARTS)
    note(n(), [
        "● 안내 (3분)",
        "API 키 항목은 반드시 짚습니다. 공개 저장소의 키는 자동으로 수집됩니다.",
        "이미지는 직접 만든 것 또는 사용 조건을 확인한 것만 포함하도록 합니다.",
    ])

    # ---------------------------------------------------------- 9. 저장소 준비
    s = S()
    header(s)
    title_block(s, "저장소 준비", "배포에 사용할 공개 저장소를 생성합니다")
    step_row(s, TOP, [
        ("STEP 1", "로그인 확인", "GitHub 로그인 상태 확인 — 미가입자만 가입"),
        ("STEP 2", "저장소 생성", "New repository · 이름은 영문 소문자"),
        ("STEP 3", "공개 설정", "Public 선택 — 무료 배포 조건"),
        ("STEP 4", "파일 업로드", "index.html 등 프로젝트 파일 업로드"),
    ], h=1.78)
    w, gap = grid(2, 0.30)
    y = 4.32
    lcard(s, ML, y, w, 2.45, ACCENT_L, "이름 규칙",
          ["영문 소문자와 하이픈만 사용",
           "공백과 한글은 사용하지 않음",
           "예: myapp · catch-game · quiz-app"],
          note="미가입 학생은 가입과 이메일 인증을 먼저 끝냅니다.")
    lcard(s, ML + w + gap, y, w, 2.45, BLUE, "업로드 대상",
          ["index.html — 필수",
           "이미지·소리 파일 — 사용한 경우",
           "CLAUDE.md · 개발일지 — 기록용(선택)"],
          label_color=BLUE_D)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 안내 (4분)",
        "대부분 가입이 끝난 상태이므로 로그인 확인부터 시작합니다.",
        "미가입 학생 1~2명은 도우미를 붙여 가입과 이메일 인증까지 끝내게 합니다.",
        "저장소 이름은 주소에 그대로 들어갑니다. 영문 소문자로 짧게 정하게 합니다.",
    ])

    # ---------------------------------------------------------- 10. 미션 ② 배포
    s = S()
    header(s)
    title_block(s, "미션 ② 배포", "20분 · 업로드한 파일을 공개 주소로 연결합니다")
    step_row(s, TOP, [
        ("①", "Settings", "저장소 상단의 Settings 선택"),
        ("②", "Pages", "왼쪽 메뉴에서 Pages 선택"),
        ("③", "Branch 지정", "main · /(root) 선택 후 Save"),
        ("④", "주소 확인", "1~2분 후 생성된 주소로 접속"),
    ], h=1.70)
    codebox(s, ML, 4.28, CW, 1.52, "생성되는 주소 형식",
            ["https://[계정이름].github.io/[저장소이름]/",
             "예시:  https://student01.github.io/myapp/"])
    L4.stuck_line(s, y=6.04)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 진행 (20분)",
        "주소는 바로 열리지 않습니다. 1~2분 기다린 뒤 새로고침하도록 미리 안내합니다.",
        "먼저 성공한 학생을 도우미로 지정하고, 배포에 성공한 주소를 게시판에 올리게 합니다.",
        "가입이 늦어진 학생은 가입이 끝나는 대로 이 단계에 합류시킵니다.",
        "● 자주 발생",
        "404 — 파일 이름이 index.html 이 아니거나 Branch 설정이 저장되지 않은 경우입니다.",
    ])

    # ---------------------------------------------------------- 11. 배포 검증
    s = S()
    header(s)
    title_block(s, "배포 검증", "배포 완료 여부는 다른 기기에서 확인합니다")
    w, gap = grid(3, 0.26)
    for i, (bar, lc, label, head, lines) in enumerate([
        (ACCENT, ACCENT_D, "① PC", "브라우저 확인",
         ["주소 직접 입력해 실행", "핵심 기능 전체 동작 확인"]),
        (BLUE, BLUE_D, "② 모바일", "휴대폰 확인",
         ["화면 구성이 깨지지 않는지 확인", "터치 조작으로 핵심 기능 사용"]),
        (GREEN, GREEN_D, "③ 공유", "주소 전달",
         ["QR 코드 생성해 공유", "게시판에 주소 등록"]),
    ]):
        wcard(s, ML + i * (w + gap), TOP, w, 2.45, bar, label, head, lines, label_color=lc, size=11)
    banner(s, [[("휴대폰 확인은 필수입니다. ", {}),
                ("PC에서 정상이어도 모바일에서 레이아웃이 깨지는 경우가 많습니다.", {"bold": True, "color": ACCENT_D})]],
           y=5.10, h=0.78)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 진행 (10분)",
        "QR 코드는 브라우저 주소창 공유 기능이나 QR 생성 사이트를 사용합니다.",
        "모바일에서 발견한 문제는 개발일지의 다음 작업으로 기록하게 합니다.",
    ])

    # ---------------------------------------------------------- 12. 재배포 흐름
    s = S()
    header(s)
    title_block(s, "수정 후 재배포", "배포는 한 번으로 끝나지 않습니다")
    step_row(s, TOP, [
        ("①", "로컬 수정", "Claude Code로 수정하고 브라우저에서 확인"),
        ("②", "커밋", "정상 동작 시점을 버전으로 기록"),
        ("③", "업로드", "변경된 파일을 저장소에 다시 올림"),
        ("④", "반영 확인", "1~2분 후 강제 새로고침으로 확인"),
    ], h=1.70)
    w, gap = grid(2, 0.30)
    y = 4.28
    lcard(s, ML, y, w, 2.40, ACCENT, "커밋과 배포의 차이",
          ["커밋 — 내 컴퓨터에 남기는 버전 기록",
           "배포 — 공개 주소에 반영되는 결과물",
           "커밋했다고 배포가 갱신되지는 않음"], strong=True)
    lcard(s, ML + w + gap, y, w, 2.40, PURPLE, "반영이 안 될 때",
          ["강제 새로고침 — Ctrl + F5",
           "1~2분 경과 후 다시 확인",
           "업로드한 파일 이름과 위치 재확인"],
          label_color=PURPLE_D)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 안내 (3분)",
        "커밋과 배포를 구분하지 못하면 '고쳤는데 그대로'라는 질문이 반복됩니다.",
        "다음 차시부터 수정할 때마다 이 흐름을 반복한다고 예고합니다.",
    ])

    # ---------------------------------------------------------- 13. 교차 테스트
    s = S()
    header(s)
    title_block(s, "교차 테스트", "10분 · 주소만 전달하고 설명 없이 사용하게 합니다")
    step_row(s, TOP, [
        ("①", "주소 전달", "설명 없이 URL 또는 QR만 공유"),
        ("②", "관찰", "사용 순서와 멈추는 지점을 기록"),
        ("③", "교대", "역할을 바꿔 동일하게 진행"),
    ], h=1.62)
    w, gap = grid(2, 0.30)
    y = 4.20
    lcard(s, ML, y, w, 2.45, ACCENT, "기록 항목",
          ["첫 화면에서 수행한 행동",
           "목표 달성까지 걸린 시간",
           "중단되거나 되돌아간 지점"], strong=True)
    lcard(s, ML + w + gap, y, w, 2.45, GREEN, "기록 방식",
          ["관찰한 사실만 기록한다",
           "평가나 조언은 기록하지 않는다",
           "예: '등록 버튼을 두 번 눌렀다'"],
          label_color=GREEN_D)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 진행 (10분)",
        "제작자는 설명하지 않고 관찰만 합니다. 질문을 받아도 답하지 않도록 안내합니다.",
        "관찰 기록은 다음 차시 조 작품 선정의 근거 자료가 됩니다.",
    ])

    # ---------------------------------------------------------- 14. 기록과 다음 단계
    s = S()
    header(s)
    title_block(s, "기록과 다음 단계", "10분 · 결과물과 주소를 정리하고 다음 차시를 준비합니다")
    checklist(s, ML, TOP, CW, [
        ("개발일지 갱신", "구현한 기능 · 공개 주소 · 교차 테스트에서 발견한 문제"),
        ("다음 작업 정리", "오늘 처리하지 못한 항목을 목록으로 기록"),
        ("주소 등록", "게시판에 공개 주소 등록 — 다음 차시 참고 자료"),
        ("백업", "프로젝트 폴더 전체를 압축해 보관 — .git 폴더 포함"),
    ], row_h=0.62, gap=0.08)
    banner(s, [[("다음 차시 — ", {}),
                ("개인 작품을 공유하고 조별 대표작을 선정합니다.", {"bold": True, "color": ACCENT_D}),
                (" 오늘의 주소와 관찰 기록이 선정 근거가 됩니다.", {})]],
           y=5.74, h=0.78)
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 마무리 (10분)",
        "공개 주소를 게시판에 올렸는지 확인합니다. 다음 차시 진행이 이 목록에서 시작됩니다.",
        "백업은 폴더 전체 압축입니다. .git 폴더가 포함되어야 버전 기록이 유지됩니다.",
    ])

    # ---------------------------------------------------------- 15. 참고: 배포 문제 대응
    s = S()
    header(s)
    title_block(s, "참고: 배포 문제 대응", "증상별로 확인 순서를 정리했습니다")
    table_rows(s, TOP, [
        ("404 오류", "파일 이름이 index.html 인지, Pages의 Branch 설정이 저장되었는지 확인"),
        ("흰 화면", "브라우저 개발자 도구(F12)의 오류 메시지 확인 — 경로 오류가 대부분"),
        ("이미지 누락", "파일 이름의 대소문자와 상대 경로 확인 — 서버는 대소문자를 구분"),
        ("수정 미반영", "1~2분 경과 후 Ctrl + F5 — 업로드 완료 여부 확인"),
        ("모바일 깨짐", "화면 폭 기준 레이아웃 수정 요청 — 다음 차시 작업으로 기록"),
    ], col1=2.6, row_h=0.64, gap=0.09, head=("증상", "확인 순서"))
    footer_rule(s)
    B.page(s, 3, TOTAL_PARTS)
    note(n(), [
        "● 활용",
        "미션 ② 진행 중 화면에 띄워 두면 질문이 줄어듭니다.",
        "해결되지 않는 경우 증상과 확인한 단계를 기록하고 짝 노트북에서 공동 실습합니다.",
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
    print(DECK)
    print("  슬라이드 %d장 · 발표자 노트 %d개" % (len(prs.slides), len(NOTES)))
    if B.WARN:
        print("  [카드 넘침] %d건" % len(B.WARN))
        for w_ in B.WARN:
            print("    " + w_)
    probs = [x for x in V3.lint(prs) if "상자 1.20 in : 05" not in x]
    if probs:
        print("  [글상자 검사] %d건" % len(probs))
        for pr in probs:
            print("    " + pr)
    if not B.WARN and not probs:
        print("  넘침 없음")
