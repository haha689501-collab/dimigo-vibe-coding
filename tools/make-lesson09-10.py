# -*- coding: utf-8 -*-
"""
9·10차시 통합 PPT 생성기 — 「출시 후보 · 작품 쇼케이스」

실행:  python tools/make-lesson09-10.py
결과:  디미고_바이브코딩_9-10차시_출시후보와쇼케이스.pptx   (프로젝트 최상위, 새 파일)

- 9차시: 우선순위 · 블루팀 패치 · UI 완성 · 크레딧 → 기능 동결 · 기기 점검 · TEST.md · README · v1.0-rc
- 10차시: 최종 배포 점검 · v1.0 릴리스 · URL · QR · 발표 · 데모 대비 · 쇼케이스 · 상호 평가 · 회고
- 저장소 초대 협업 없음: 통합 담당이 버전 폴더(v1.0-rc → v1.0)로 관리, 배포는 Netlify Drop.
- 7·8차시 덱과 같은 디자인. 90분 구성 · 다음 차시 예고 슬라이드는 넣지 않습니다.
"""
import importlib.util
import os
import sys

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
L6 = _load("l6", "make-lesson06.py")

DECK = "디미고_바이브코딩_9-10차시_출시후보와쇼케이스.pptx"

ML, CW, SW, SH, TOP, FOOT_Y = B.ML, B.CW, B.SW, B.SH, B.TOP, B.FOOT_Y
INK, BODY, MUTED, FAINT = B.INK, B.BODY, B.MUTED, B.FAINT
ACCENT, ACCENT_D, ACCENT_L = B.ACCENT, B.ACCENT_D, B.ACCENT_L
LINE, CARD, TINT, DARK, WHITE = B.LINE, B.CARD, B.TINT, B.DARK, B.WHITE
BLUE, GREEN, PURPLE = B.BLUE, B.GREEN, B.PURPLE
rect, text, header, footer_rule, title_block = B.rect, B.text, B.header, B.footer_rule, B.title_block
wcard, codebox, step_row = B.wcard, B.codebox, B.step_row
grid, lcard, minicard, checklist = L4.grid, L4.lcard, L4.minicard, L4.checklist
BLUE_D, GREEN_D, PURPLE_D, RED = L4.BLUE_D, L4.GREEN_D, L4.PURPLE_D, L4.RED
table, bar_note = L6.table, L6.bar_note

NOTES = {}
PAGES = []


def note(i, lines):
    NOTES[i] = "\n".join(lines)


def page(s, section):
    PAGES.append((s, section))


def retro(s, items):
    w, gap = grid(3, 0.26)
    for i, (num, h1, h2) in enumerate(items):
        x = ML + i * (w + gap)
        rect(s, x, TOP, w, 1.30, fill=CARD, line=LINE, radius=0.05)
        text(s, x + 0.24, TOP + 0.20, 0.5, 0.3, [[(num, {"bold": True, "size": 12, "color": ACCENT})]])
        text(s, x + 0.24, TOP + 0.50, w - 0.48, 0.34, [[(h1, {"bold": True, "size": 14, "color": INK})]])
        text(s, x + 0.24, TOP + 0.88, w - 0.48, 0.30, [[(h2, {"size": 10.5, "color": MUTED})]])


def build():
    prs = B.new_deck()
    S = lambda: B.blank(prs)
    n = lambda: B.CUR["n"]

    # ------------------------------------------------------------ 표지
    s = S()
    text(s, ML, 0.44, 5, 0.3, [[("A I   ×   B U I L D", {"bold": True, "size": 11, "color": ACCENT, "space": 1.3})]])
    text(s, SW - ML - 7.4, 0.46, 7.4, 0.3,
         [[("한국디지털미디어고등학교 · 9 · 10차시", {"bold": True, "size": 10, "color": FAINT, "space": 1.0})]],
         align=PP_ALIGN.RIGHT)
    text(s, ML, 1.70, 3.4, 1.2, [[("09 · 10", {"bold": True, "size": 60, "color": ACCENT_L})]])
    text(s, ML, 2.84, CW, 1.0, [[("출시 후보 · 작품 쇼케이스", {"bold": True, "size": 40, "color": INK})]])
    rect(s, ML, 3.94, 0.9, 0.05, fill=ACCENT)
    text(s, ML, 4.20, CW * 0.9, 0.5,
         [[("검증 결과를 패치해 출시 후보를 만들고, 작품과 개발 과정을 공개합니다", {"size": 17, "color": BODY})]])
    w, gap = grid(3, 0.22)
    for i, (h1, h2) in enumerate([("패치", "우선순위 · 블루팀 · UI 완성"),
                                  ("출시 후보", "기능 동결 · TEST.md · v1.0-rc"),
                                  ("쇼케이스", "배포 · 발표 · 상호 평가")]):
        minicard(s, ML + i * (w + gap), 5.20, w, 1.20, h1, h2, bar=ACCENT if i == 1 else ACCENT_L)
    footer_rule(s)
    text(s, ML, FOOT_Y + 0.12, 8, 0.26,
         [[("검증된 작품에서, 공개하는 작품으로", {"size": 9.5, "color": FAINT, "space": 0.8})]])
    note(n(), [
        "[구성] 9차시 = PART 1 패치 · UI + PART 2 출시 후보 · 10차시 = PART 3 배포와 쇼케이스",
        "",
        "[대본]",
        "\"8차시에 찾은 문제 목록이 오늘의 출발점입니다. 9차시에 고치고 기능을 멈춘 뒤 출시 후보를 만들고,",
        "10차시에는 최종 배포하고 작품과 개발 과정을 발표합니다.\"",
    ])

    # ============================================================ PART 1 · 패치와 UI 완성 (9차시)
    s = S()
    B.part_slide(s, "P A R T   1", "패치 · UI 완성", "9차시 전반 · 검증 결과를 반영합니다", 1, 3)
    note(n(), ["[시간] 전환",
               "[대본] \"찾은 문제를 전부 고칠 시간은 없습니다. 먼저 고를 것을 정합니다.\""])

    # ------------------------------------------------------------ 우선순위
    s = S()
    header(s)
    title_block(s, "우선순위 결정", "관찰 · 보고 결과를 세 묶음으로 나눕니다")
    w, gap = grid(3, 0.26)
    wcard(s, ML, TOP, w, 2.80, RED, "NOW · 이번 차시", "높음 · S1 · 포기 지점",
          ["코드 실행 취약점", "진행 불가 버그", "다수가 막힌 화면"], label_color=RED, size=11, strong=True)
    wcard(s, ML + w + gap, TOP, w, 2.80, ACCENT, "LATER · v1.1", "중간 · S2 · 머뭇거림",
          ["데이터 변조", "핵심 기능 오동작", "안내 문구 부족"], label_color=ACCENT_D, size=11)
    wcard(s, ML + 2 * (w + gap), TOP, w, 2.80, GREEN, "WON'T · 기록만", "구조적 한계 · 범위 밖",
          ["클라이언트 점수 신뢰", "서버가 필요한 기능", "README에 한계 명시"], label_color=GREEN_D, size=11)
    codebox(s, ML, 5.24, CW, 1.14, "PRD 결정 기록에 추가",
            ["## 결정 기록 (9차시)  ·  항목 · 출처(관찰 / 레드팀 / AI 감사) · 분류(NOW / LATER / WON'T) · 이유"])
    footer_rule(s); page(s, "패치")
    note(n(), [
        "[시간] 5분",
        "",
        "[대본]",
        "\"고칠 것을 세 묶음으로 나눕니다. NOW는 오늘 반드시, LATER는 출시 후 v1.1, WON'T는 고치지 않지만 기록합니다.",
        "점수 조작은 서버 없이 완전히 막을 수 없습니다. 이건 WON'T로 분류하고 README에 한계로 적는 게 정직한 대응입니다.\"",
    ])

    # ------------------------------------------------------------ 블루팀 패치
    s = S()
    header(s)
    title_block(s, "블루팀 패치", "증상이 아니라 원인 패턴을 제거합니다")
    table(s, ML, TOP, CW, [2.4, 4.4, 5.07],
          ["취약점", "대응", "확인"],
          [["XSS", "textContent · 요소 생성으로 출력", "보고서 페이로드 재입력"],
           ["입력 경계", "길이 · 형식 검증 · 공백 정리", "빈 값 · 200자 · 특수문자"],
           ["저장 변조", "스키마 검증 · 범위 확인 · 기본값 복구", "손상 값 주입 후 새로고침"],
           ["상태 노출", "모듈 스코프 · 전역 최소화", "Object.keys(window)"],
           ["비밀값", "키 제거 · 즉시 폐기 · 재발급", "버전 파일 · 배포본까지 검색"]],
          row_h=0.44)
    codebox(s, ML, 5.10, CW, 1.36, "PROMPT",
            ["보고서 #2의 원인 패턴을 프로젝트 전체에서 찾아 같은 방식으로 고쳐 줘.",
             "고친 위치 목록과 재현 절차로 다시 확인하는 방법도 보여 줘."])
    footer_rule(s); page(s, "패치")
    note(n(), [
        "[시간] 15분",
        "",
        "[대본]",
        "\"보고된 한 곳만 고치면 같은 패턴이 다른 곳에 남아 있습니다. innerHTML 하나가 문제였다면 프로젝트 전체의 innerHTML을 봐야 합니다.",
        "비밀값은 코드에서 지우는 것으로 끝나지 않습니다. 예전 버전 파일과 이미 배포된 주소에 남아 있으니 키 자체를 폐기하고 새로 발급하세요.",
        "패치가 배포되면 보고한 조의 레드팀이 배포 주소에서 재현 절차로 다시 확인합니다.\"",
    ])

    # ------------------------------------------------------------ UI 완성
    s = S()
    header(s)
    title_block(s, "UI 완성", "DESIGN.md 기준으로 정리하고 대표 연출 하나를 완성합니다")
    w, gap = grid(3, 0.26)
    wcard(s, ML, TOP, w, 2.60, ACCENT, "정리", "토큰 위반 제거",
          ["간격 · 정렬 · 글자 위계", "강조색 한 화면 1곳"], label_color=ACCENT_D, size=11)
    wcard(s, ML + w + gap, TOP, w, 2.60, BLUE, "안내", "관찰 결과 반영",
          ["시작 · 조작 안내 노출", "결과 화면 재시작 경로"], label_color=BLUE_D, size=11)
    wcard(s, ML + 2 * (w + gap), TOP, w, 2.60, PURPLE, "대표 연출", "기억에 남는 한 장면",
          ["득점 · 전환 · 결과 중 1곳", "reduced-motion 대응"], label_color=PURPLE_D, size=11)
    codebox(s, ML, 5.00, CW, 1.36, "PROMPT",
            ["DESIGN.md를 기준으로 index.html을 검사해 토큰 밖의 값을 정리해 줘. 기능 코드는 바꾸지 마.",
             "그다음 결과 화면에 대표 연출 하나를 제안 3개로 보여 주고, 내가 고른 것만 구현해 줘."])
    footer_rule(s); page(s, "패치")
    note(n(), [
        "[시간] 15분",
        "",
        "[대본]",
        "\"5차시에 만든 DESIGN.md를 다시 씁니다. 조원 여러 명이 코드를 고치면서 토큰 밖의 값이 다시 생겼을 겁니다.",
        "관찰 기록에서 '어디 눌러야 해요?'가 나왔다면 연출보다 안내가 먼저입니다.",
        "대표 연출은 하나만. 여러 개를 넣으면 기억에 남는 게 없습니다. 움직임을 줄이는 설정(prefers-reduced-motion)도 함께 처리하세요.\"",
    ])

    # ------------------------------------------------------------ 크레딧 · Before/After
    s = S()
    header(s)
    title_block(s, "크레딧 · 전후 비교", "출처를 남기고 변화를 증거로 보여 줍니다")
    cw_ = 6.20
    codebox(s, ML, TOP, cw_, 3.92, "README.md · 크레딧",
            ["## Credits",
             "- 효과음: [이름] · [출처 URL] · CC0",
             "- 배경음: [이름] · [출처 URL] · CC BY 4.0",
             "- 글꼴: Pretendard · SIL OFL",
             "- 이미지: Claude 생성 (프롬프트 기록)",
             "- 코드 지원: Claude Code",
             "",
             "## Known limits",
             "- 점수는 브라우저에서 계산 · 조작 가능"])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    table(s, rx, TOP, rw, [1.75, rw - 1.75], ["Before / After", "증거"],
          [["화면", "같은 장면 스크린샷 2장"],
           ["사용성", "첫 판 완료 시간 · 포기 수"],
           ["보안", "보고 건수 · 패치 기록"],
           ["품질", "남은 S1 · S2 버그 수"]],
          row_h=0.50)
    bar_note(s, [[("라이선스를 확인할 수 없는 소재는 ", {}), ("교체하거나 제거", {"bold": True, "color": ACCENT_D})]])
    page(s, "패치")
    note(n(), [
        "[시간] 5분",
        "",
        "[대본]",
        "\"쓴 소리 · 이미지 · 글꼴은 전부 출처와 라이선스를 적습니다. 확인이 안 되면 바꾸세요.",
        "Known limits는 약점이 아니라 설계 판단을 보여 주는 항목입니다. 오늘 WON'T로 분류한 것을 여기 적습니다.",
        "전후 비교는 스크린샷만이 아니라 숫자로도 남깁니다. 첫 판 완료 시간이 줄었는지, 열린 S1 버그가 0개인지.\"",
    ])

    # ============================================================ PART 2 · 출시 후보 (9차시)
    s = S()
    B.part_slide(s, "P A R T   2", "출시 후보", "9차시 후반 · 기능을 멈추고 품질을 확정합니다", 2, 3)
    note(n(), ["[시간] 전환",
               "[대본] \"지금부터 새 기능은 없습니다. 만든 것을 확실하게 만드는 시간입니다.\""])

    # ------------------------------------------------------------ 기능 동결
    s = S()
    header(s)
    title_block(s, "기능 동결", "출시 후보 전까지 통합본에 들어갈 수 있는 변경을 제한합니다")
    w, gap = grid(2, 0.30)
    wcard(s, ML, TOP, w, 2.60, GREEN, "ALLOW · 통합 가능", "수정 · 정리 · 문서",
          ["버그 · 취약점 수정", "문구 · 토큰 위반 정리", "TEST.md · README.md"], label_color=GREEN_D, size=11)
    wcard(s, ML + w + gap, TOP, w, 2.60, RED, "BLOCK · 통합 금지", "추가 · 구조 변경",
          ["새 기능 · 새 화면", "구조 변경 · 파일 이동", "새 라이브러리 · 외부 API"], label_color=RED, size=11)
    codebox(s, ML, 5.10, CW, 1.30, "CLAUDE.md에 추가",
            ["## 기능 동결 (9차시~)  ·  새 기능 · 구조 변경 금지  ·  수정은 BUGS.md 번호가 있는 버그만  ·  변경 파일 목록을 먼저 보여 준다"])
    footer_rule(s); page(s, "출시 후보")
    note(n(), [
        "[시간] 3분",
        "",
        "[대본]",
        "\"실제 출시 직전에 회사들이 하는 일입니다. Feature freeze, 기능 동결.",
        "새 기능은 아무리 작아도 새 버그를 데려옵니다. 지금부터 통합 담당은 고치는 변경만 받습니다.",
        "이 규칙을 CLAUDE.md에 넣어 두면 AI도 같이 지킵니다.\"",
    ])

    # ------------------------------------------------------------ 기기 점검
    s = S()
    header(s)
    title_block(s, "기기 · 환경 점검", "배포 주소로 실제 기기에서 확인합니다")
    table(s, ML, TOP, CW, [2.9, 2.24, 2.24, 2.24, 2.25],
          ["환경", "시작", "조작", "결과", "저장"],
          [["PC · Chrome", "", "", "", ""],
           ["PC · 다른 브라우저", "", "", "", ""],
           ["휴대폰 · 세로", "", "", "", ""],
           ["휴대폰 · 가로", "", "", "", ""],
           ["창 축소 · 확대 200%", "", "", "", ""]],
          row_h=0.50)
    bar_note(s, [[("실제 휴대폰 1대 이상 필수 — ", {}),
                  ("터치 · 화면 크기 · 성능은 DevTools 기기 모드와 다릅니다", {"bold": True, "color": ACCENT_D})]], y=5.92)
    footer_rule(s); page(s, "출시 후보")
    note(n(), [
        "[시간] 8분",
        "",
        "[대본]",
        "\"칸마다 통과 · 실패를 적습니다. 실패하면 바로 BUGS.md에 기록합니다.",
        "휴대폰은 반드시 실제 기기로 확인하세요. 터치 이벤트와 화면 높이가 개발자 도구와 다르게 동작하는 경우가 많습니다.\"",
        "",
        "[운영] 배포 주소를 QR로 띄우면 휴대폰 접속이 빠릅니다 (Chrome 주소창 → 공유 → QR 코드).",
    ])

    # ------------------------------------------------------------ TEST.md
    s = S()
    header(s)
    title_block(s, "TEST.md", "조 PRD의 완료 기준을 전부 판정합니다")
    cw_ = 6.80
    codebox(s, ML, TOP, cw_, 3.92, "TEST.md · 3차시 완료 확인 형식",
            ["| 기준 | 조건 | 예상 | 실제 | 판정 |",
             "|---|---|---|---|---|",
             "| AC-1 | 빈 닉네임 저장 | 안내 표시 | 안내 표시 | 통과 |",
             "| AC-2 | 새로고침 | 기록 유지 | 유지 | 통과 |",
             "| SEC-1 | <b> 입력 | 글자 그대로 | 그대로 | 통과 |",
             "| DEV-3 | 휴대폰 가로 | 버튼 노출 | 잘림 | 실패 #21 |"])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    codebox(s, rx, TOP, rw, 2.40, "PROMPT",
            ["조 PRD의 완료 기준 · 보안",
             "요구사항 · 기기 점검표로",
             "TEST.md 표를 만들어 줘.",
             "판정 칸은 비워 둬."])
    lcard(s, rx, TOP + 2.56, rw, 1.36, GREEN, "판정은 사람이",
          ["직접 실행 후 기록"], label_color=GREEN_D, size=11)
    bar_note(s, [[("실패 항목은 숨기지 않고 ", {}), ("BUGS.md 번호와 함께 기록", {"bold": True, "color": ACCENT_D})]])
    page(s, "출시 후보")
    note(n(), [
        "[시간] 10분",
        "",
        "[대본]",
        "\"3차시에 쓴 완료 확인 형식 그대로입니다. 조건 · 예상 · 실제 · 판정.",
        "표는 AI가 만들어도 되지만 판정은 반드시 사람이 직접 실행해서 적습니다. AI에게 판정을 맡기면 확인하지 않은 통과가 생깁니다.",
        "레드팀 보고서의 재현 절차도 SEC 항목으로 넣으세요. 패치가 유지되는지 확인하는 회귀 테스트가 됩니다.\"",
    ])

    # ------------------------------------------------------------ README
    s = S()
    header(s)
    title_block(s, "README.md", "처음 보는 사람이 실행하고 평가할 수 있는 문서")
    cw_ = 6.20
    codebox(s, ML, TOP, cw_, 3.92, "README.md 구조",
            ["# 작품 이름 — 한 문장 소개",
             "## 플레이  배포 주소 · QR",
             "## 조작법  키 · 터치",
             "## 팀  조 이름 · 역할 (실명 제외)",
             "## 개발 과정  PRD · 기능 분담 · 검증 요약",
             "## Credits  소재 · 출처 · 라이선스",
             "## Known limits  설계상 한계",
             "## Security  검증 범위 · 대응"])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    lcard(s, rx, TOP, rw, 1.86, ACCENT, "공개 범위",
          ["실명 대신 조 이름 · 역할", "학교 · 학번 · 사진 제외"], size=11)
    lcard(s, rx, TOP + 2.02, rw, 1.90, BLUE, "Security 항목",
          ["레드팀 보고 · 패치 기록", "남은 위험은 Known limits"], label_color=BLUE_D, size=11)
    bar_note(s, [[("README는 ", {}), ("작품의 첫 화면", {"bold": True, "color": ACCENT_D}),
                  (" — 작품 폴더와 함께 제출", {})]])
    page(s, "출시 후보")
    note(n(), [
        "[시간] 7분",
        "",
        "[대본]",
        "\"README는 작품을 처음 보는 사람이 가장 먼저 읽는 문서입니다. 10차시 평가도 README에서 시작합니다.",
        "Security 항목에는 어떤 검증을 받았고 무엇을 고쳤는지, 무엇은 구조상 남아 있는지를 적습니다. 숨기는 것보다 신뢰를 줍니다.\"",
    ])

    # ------------------------------------------------------------ 출시 후보 태그
    s = S()
    header(s)
    title_block(s, "출시 후보 · v1.0-rc", "조건을 통과한 통합본을 버전으로 고정합니다")
    cw_ = 6.55
    checklist(s, ML, TOP, cw_, [
        ("S1 버그", "남은 S1 버그 0개"),
        ("취약점", "높음 등급 미해결 0개"),
        ("TEST.md", "MUST · SEC 항목 전부 통과"),
        ("기기", "PC · 휴대폰 한 바퀴 동작"),
        ("문서", "README · Credits · Known limits"),
    ], row_h=0.70, gap=0.105)
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    codebox(s, rx, TOP, rw, 3.92, "PROMPT · 통합 담당",
            ["왼쪽 조건을 index.html ·",
             "TEST.md · BUGS.md 기준으로",
             "확인해 줘. 아직 복사하지 마.",
             "",
             "모두 통과하면 작품 폴더를",
             "v1.0-rc 폴더로 복사하고,",
             "통합한 기능 목록으로",
             "변경 기록 초안을 써 줘."])
    bar_note(s, [[("rc = Release Candidate — ", {}), ("10차시에 그대로 v1.0이 될 수 있는 상태", {"bold": True, "color": ACCENT_D})]])
    page(s, "출시 후보")
    note(n(), [
        "[시간] 5분",
        "",
        "[대본]",
        "\"다섯 조건을 다 통과하면 지금 작품 폴더를 v1.0-rc 폴더로 복사해 고정합니다. 이 폴더는 이후 수정하지 않습니다.",
        "10차시에 문제가 없으면 이 폴더가 그대로 v1.0이 됩니다. 문제가 있으면 고치고 rc2 폴더를 만듭니다.\"",
        "",
        "[운영] Git을 쓰는 조는 같은 시점에 커밋하고 v1.0-rc 태그를 붙여도 됩니다. 필수는 아닙니다.",
    ])

    # ------------------------------------------------------------ 9차시 미션 체크
    s = S()
    header(s)
    title_block(s, "9차시 미션 체크", "조별로 확인하고 출시 후보 주소를 제출합니다")
    checklist(s, ML, TOP, 6.55, [
        ("패치", "NOW 항목 통합 · 보고 조 재확인"),
        ("UI", "토큰 위반 0 · 안내 보강 · 대표 연출 1개"),
        ("동결", "CLAUDE.md 동결 규칙 · 새 기능 0"),
        ("검증", "TEST.md 판정 · 기기 점검표"),
        ("rc", "v1.0-rc 폴더 · 변경 기록"),
    ], row_h=0.70, gap=0.105)
    rx = ML + 6.55 + 0.26
    rw = CW - 6.55 - 0.26
    wcard(s, rx, TOP, rw, 1.90, ACCENT, "제출", "반 게시판",
          ["배포 주소 · v1.0-rc 폴더"], size=11, strong=True)
    wcard(s, rx, TOP + 2.04, rw, 1.88, GREEN, "10차시 준비", None,
          ["발표 5분 구성 초안", "6차시 티저 영상 확인"], label_color=GREEN_D, size=11)
    bar_note(s, [[("10차시에는 ", {}), ("코드 수정 없이 배포 · 발표", {"bold": True, "color": ACCENT_D}), ("만 합니다", {})]])
    page(s, "출시 후보")
    note(n(), [
        "[시간] 3분",
        "[대본] \"10차시는 고치는 시간이 아닙니다. 오늘 만든 rc가 그대로 발표됩니다. 발표 구성 초안과 6차시에 만든 영상을 확인해 두세요.\"",
    ])

    # ------------------------------------------------------------ 9차시 회고
    s = S()
    header(s)
    title_block(s, "회고 · 출시 후보", "출시 직전의 판단을 기록합니다")
    retro(s, [("①", "동결 후 참은 기능", "왜 지금이 아닌가"),
              ("②", "패치에서 찾은 같은 패턴", "보고된 곳 말고 어디에"),
              ("③", "TEST.md 실패 항목", "고칠 것 · 한계로 둘 것")])
    footer_rule(s); page(s, "출시 후보")
    note(n(), ["[시간] 3분"])

    # ============================================================ PART 3 · 배포와 쇼케이스 (10차시)
    s = S()
    B.part_slide(s, "P A R T   3", "배포 · 쇼케이스", "10차시 · 완성 결과와 발전 과정을 보여 줍니다", 3, 3)
    note(n(), ["[시간] 전환",
               "[대본] \"오늘은 만든 것을 공개하는 날입니다. 작품만이 아니라 어떻게 만들었는지를 보여 줍니다.\""])

    # ------------------------------------------------------------ 최종 배포 점검
    s = S()
    header(s)
    title_block(s, "최종 배포 점검", "공개 전에 주소 · 기록 · 정보를 확인합니다")
    table(s, ML, TOP, CW, [2.4, 4.9, 4.57],
          ["항목", "확인", "방법"],
          [["접속", "로그인 안 된 환경에서 첫 화면", "시크릿 창 · 휴대폰"],
           ["오류", "한 바퀴 동안 콘솔 오류 0", "DevTools Console"],
           ["경로", "이미지 · 소리 · 스크립트 로드", "Network 404 확인"],
           ["비밀값", "배포 폴더 모든 파일의 키 · 토큰", "폴더 전체 검색"],
           ["개인정보", "실명 · 사진 · 주석 속 메모", "README · 화면 · 주석"],
           ["라이선스", "Credits와 실제 소재 일치", "파일 목록 대조"]],
          row_h=0.50)
    bar_note(s, [[("배포하면 ", {}), ("화면이 아니라 폴더 전체", {"bold": True, "color": ACCENT_D}), ("가 공개됩니다 — 필요한 파일만", {})]],
             y=5.92)
    footer_rule(s); page(s, "쇼케이스")
    note(n(), [
        "[시간] 8분",
        "",
        "[대본]",
        "\"Netlify에 폴더를 올리면 화면에 안 보이는 파일도 주소만 알면 누구나 열 수 있습니다. feature 파일, 백업 파일, 메모까지 확인하세요.",
        "Claude에게: '배포할 폴더의 모든 파일에서 키 · 토큰 · 실명으로 보이는 문자열을 찾아 목록만 보여 줘.'",
        "키가 나오면 코드에서 지우는 것으로 끝이 아닙니다. 키를 폐기하고 새로 발급합니다.\"",
    ])

    # ------------------------------------------------------------ v1.0 · URL · QR
    s = S()
    header(s)
    title_block(s, "v1.0 · URL · QR", "rc를 정식 버전으로 확정하고 접속 경로를 만듭니다")
    step_row(s, TOP, [("1", "rc 재확인", "점검표 통과"),
                      ("2", "v1.0 확정", "rc 폴더 → v1.0"),
                      ("3", "배포", "v1.0 폴더 업로드"),
                      ("4", "URL", "README · 게시판"),
                      ("5", "QR", "부스 카드 출력")], h=1.40)
    w, gap = grid(2, 0.30)
    lcard(s, ML, TOP + 1.62, w, 2.00, ACCENT, "QR 만들기",
          ["Chrome 주소창 → 공유 → QR 코드", "휴대폰으로 읽어 접속 확인"], size=11)
    lcard(s, ML + w + gap, TOP + 1.62, w, 2.00, BLUE, "부스 카드",
          ["작품 이름 · 한 문장 · QR", "조 이름 · 조작법 한 줄"], label_color=BLUE_D, size=11)
    bar_note(s, [[("rc 이후 수정이 있었다면 ", {}), ("rc2 → 재점검 → v1.0", {"bold": True, "color": ACCENT_D})]], y=5.92)
    footer_rule(s); page(s, "쇼케이스")
    note(n(), [
        "[시간] 7분",
        "",
        "[대본]",
        "\"9차시에 만든 rc에 문제가 없으면 그 폴더 이름을 v1.0으로 바꾸고 배포합니다. 변경 기록은 rc 때 쓴 초안을 다듬습니다.",
        "QR은 브라우저에서 바로 만들 수 있습니다. 반드시 휴대폰으로 읽어서 실제로 열리는지 확인하세요.\"",
    ])

    # ------------------------------------------------------------ 발표 구성
    s = S()
    header(s)
    title_block(s, "발표 구성 · 5분", "작품보다 판단의 과정을 보여 줍니다")
    step_row(s, TOP, [("0:00", "티저", "6차시 15초 영상"),
                      ("0:15", "문제 · 사용자", "조 PRD 한 문장"),
                      ("0:45", "라이브 데모", "한 바퀴 · 2분"),
                      ("2:45", "개발 과정", "결정 · 검증 · 패치"),
                      ("4:15", "한계 · 배운 것", "Known limits")], h=1.40)
    table(s, ML, TOP + 1.62, CW, [2.6, 9.27], ["개발 과정 증거", "보여 줄 것"],
          [["결정", "대표작 선정 점수표 · 강점 흡수 · NON-GOAL"],
           ["협업", "기능 분담표 · 통합 순서 · 개발일지"],
           ["검증", "레드팀 보고 → 패치 기록 · Before / After"]],
          row_h=0.46)
    bar_note(s, [[("\"열심히 했다\" 대신 ", {}), ("기록으로 증명", {"bold": True, "color": ACCENT_D})]], y=5.92)
    footer_rule(s); page(s, "쇼케이스")
    note(n(), [
        "[시간] 3분 (설명) · 조당 발표 5분 + 질문 2분",
        "",
        "[대본]",
        "\"6차시에 만든 15초 영상으로 시작합니다. 그다음 문제와 사용자를 한 문장으로, 바로 데모.",
        "개발 과정은 말이 아니라 기록으로 보여 줍니다. 대표작을 고른 점수표, 누가 어떤 기능을 만들었는지, 레드팀 보고와 패치.",
        "마지막에 한계를 직접 말하세요. 무엇을 알고 남겨 두었는지가 설계 수준입니다.\"",
    ])

    # ------------------------------------------------------------ 데모 대비
    s = S()
    header(s)
    title_block(s, "데모 대비", "라이브 데모는 실패를 전제로 준비합니다")
    w, gap = grid(4, 0.18)
    for i, (head, bar, lc, lines) in enumerate([
            ("백업 영상", ACCENT, ACCENT_D, ["한 바퀴 화면 녹화", "1분 이내"]),
            ("로컬 사본", BLUE, BLUE_D, ["v1.0 폴더 · 노트북", "네트워크 없이 실행"]),
            ("시나리오", GREEN, GREEN_D, ["조작 순서 고정", "저장 데이터 초기화"]),
            ("역할", PURPLE, PURPLE_D, ["진행 · 조작 · 시간", "질문 응답 담당"])]):
        lcard(s, ML + i * (w + gap), TOP, w, 2.10, bar, head, lines, label_color=lc, size=10.5)
    codebox(s, ML, TOP + 2.30, CW, 1.30, "녹화",
            ["Windows: Win + Alt + R (Xbox Game Bar)  ·  휴대폰: 화면 녹화  ·  데모 전 화면 확대 · 알림 끄기"])
    bar_note(s, [[("데모가 멈추면 ", {}), ("10초 안에 백업 영상으로 전환", {"bold": True, "color": ACCENT_D})]], y=5.92)
    footer_rule(s); page(s, "쇼케이스")
    note(n(), [
        "[시간] 준비 10분",
        "",
        "[대본]",
        "\"라이브 데모는 꼭 한 번은 실패합니다. 네트워크, 화면 연결, 저장된 기록 때문에요.",
        "백업 영상과 로컬 사본을 준비하고, 멈추면 고치려 하지 말고 바로 영상으로 넘어가세요.\"",
    ])

    # ------------------------------------------------------------ 쇼케이스 운영
    s = S()
    header(s)
    title_block(s, "쇼케이스 운영", "발표 후 부스를 돌며 서로의 작품을 플레이합니다")
    step_row(s, TOP, [("1", "조별 발표", "5분 + 질문 2분"),
                      ("2", "부스 설치", "QR 카드 · 기기 1대"),
                      ("3", "순환 플레이", "조당 5분 · 시계 방향"),
                      ("4", "피드백 카드", "좋았던 점 · 제안")], h=1.40)
    w, gap = grid(2, 0.30)
    lcard(s, ML, TOP + 1.62, w, 2.00, ACCENT, "부스 담당",
          ["조원 1명 상주 · 교대", "설명은 첫 화면까지만"], size=11)
    lcard(s, ML + w + gap, TOP + 1.62, w, 2.00, GREEN, "피드백 카드",
          ["장면 · 이유를 한 줄로", "제안은 근거와 함께"], label_color=GREEN_D, size=11)
    bar_note(s, [[("받은 피드백은 ", {}), ("LATER 목록에 추가 — v1.1의 시작", {"bold": True, "color": ACCENT_D})]], y=5.92)
    footer_rule(s); page(s, "쇼케이스")
    note(n(), [
        "[시간] 발표 이후 20분",
        "",
        "[대본]",
        "\"부스 담당은 설명을 최소로 합니다. 8차시 사용자 테스트와 같습니다. 첫 화면만 보고 쓸 수 있어야 합니다.",
        "피드백 카드는 조 PRD의 LATER 목록으로 옮기세요. 수업은 끝나도 작품은 v1.1로 계속될 수 있습니다.\"",
    ])

    # ------------------------------------------------------------ 상호 평가
    s = S()
    header(s)
    title_block(s, "상호 평가", "작품 기준으로 판정 가능한 항목만 평가합니다")
    table(s, ML, TOP, CW, [2.2, 6.6, 3.07],
          ["기준", "판정 질문", "근거"],
          [["완결성", "설명 없이 시작 → 결과 → 재시작이 되는가", "직접 플레이"],
           ["사용성", "첫 화면에서 할 일을 알 수 있는가", "직접 플레이"],
           ["디자인", "같은 역할이 같은 모양인가 · 대표 연출", "화면"],
           ["보안", "검증 범위 · 패치 · 한계가 공개되어 있는가", "README"],
           ["과정", "결정 · 협업 · 검증이 기록으로 남았는가", "개발 기록"]],
          row_h=0.50)
    bar_note(s, [[("각 1~3점 · ", {}), ("점수마다 근거 한 줄", {"bold": True, "color": ACCENT_D}),
                  (" · 자기 조는 평가하지 않음", {})]], y=5.92)
    footer_rule(s); page(s, "쇼케이스")
    note(n(), [
        "[시간] 쇼케이스 중 진행",
        "",
        "[대본]",
        "\"'재미있다'는 평가할 수 없습니다. 판정 가능한 질문만 씁니다. 2차시 완성 기준과 같은 원리입니다.",
        "보안과 과정은 화면이 아니라 README와 개발 기록을 보고 판단합니다.\"",
    ])

    # ------------------------------------------------------------ 프로젝트 여정
    s = S()
    header(s)
    title_block(s, "10회기 정리", "프롬프트 한 줄에서 검증된 출시까지")
    step_row(s, TOP, [("1", "프롬프트", "의도 · 제약"),
                      ("2·3", "PRD · MVP", "결정을 파일로"),
                      ("4", "Claude Code", "계획 · 커밋"),
                      ("5", "디자인 시스템", "토큰 · 규칙")], h=1.40)
    step_row(s, TOP + 1.60, [("6", "발표 · 모션", "PRD 검증 · 영상"),
                             ("7·8", "조 작품 · 검증", "완결 · 레드팀"),
                             ("9·10", "출시 · 공개", "동결 · 릴리스")], h=1.40)
    bar_note(s, [[("AI가 코드를 쓰는 동안 ", {}),
                  ("사람은 결정하고, 기록하고, 검증했다", {"bold": True, "color": ACCENT_D})]], y=5.56, h=0.62)
    footer_rule(s); page(s, "쇼케이스")
    note(n(), [
        "[시간] 3분",
        "",
        "[대본]",
        "\"처음에는 프롬프트 한 줄이었습니다. 그다음 결정을 PRD로 남겼고, Claude Code로 파일을 다뤘고, 디자인 규칙을 만들었고,",
        "조 작품으로 합치고, 다른 조의 공격을 받아 고치고, 출시했습니다.",
        "코드는 AI가 많이 썼습니다. 하지만 무엇을 만들지, 무엇을 뺄지, 무엇이 안전한지는 여러분이 정했습니다.\"",
    ])

    # ------------------------------------------------------------ 프로젝트 회고
    s = S()
    header(s)
    title_block(s, "회고 · 프로젝트", "다음 프로젝트에 가져갈 것을 남깁니다")
    retro(s, [("①", "AI에게 맡긴 것 · 내가 정한 것", "경계를 어디에 그었나"),
              ("②", "가장 효과가 컸던 문서", "PRD · DESIGN · CLAUDE · TEST"),
              ("③", "다음에도 쓸 규칙 한 줄", "내 CLAUDE.md 첫 줄")])
    bar_note(s, [[("작품 폴더와 배포 주소는 포트폴리오 — ", {}), ("공개 범위를 확인하고 유지", {"bold": True, "color": ACCENT_D})]],
             y=4.00, h=0.62)
    footer_rule(s); page(s, "쇼케이스")
    note(n(), [
        "[시간] 5분",
        "[대본] \"세 번째 질문의 답을 실제로 적어 두세요. 다음 프로젝트를 시작할 때 CLAUDE.md 첫 줄이 됩니다.\"",
    ])

    # ------------------------------------------------------------ 꼬리말 · 노트
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
    prs.save(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, DECK))
    return prs


if __name__ == "__main__":
    B.CUR["n"] = 0
    del B.WARN[:]
    prs = build()
    print("[9·10차시] %s" % DECK)
    print("  슬라이드 %d장 · 발표자 노트 %d개" % (len(prs.slides), len(NOTES)))
    if B.WARN:
        print("  [카드 넘침] %d건" % len(B.WARN))
        for w_ in B.WARN:
            print("    " + w_)
    probs = [x for x in V3.lint(prs) if ": 09" not in x]
    if probs:
        print("  [글상자 검사] %d건" % len(probs))
        for pr in probs:
            print("    " + pr)
    if not B.WARN and not probs:
        print("  넘침 없음")
