# -*- coding: utf-8 -*-
"""
6차시 PPT 생성기 — 「미니프로젝트 발표 · GitHub 협업」

실행:  python tools/make-lesson06.py
결과:  디미고_바이브코딩_6차시_발표와GitHub협업.pptx   (프로젝트 최상위, 새 파일)

- 5차시 「디자인 시스템 v2」 덱과 같은 디자인(머리말 · 제목 · 카드 · 남색 표 · 꼬리말 구역 표시)을 씁니다.
- 발표자 노트는 [시간] · [대본] · [운영] 형식입니다.
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

DECK = "디미고_바이브코딩_6차시_발표와GitHub협업.pptx"

ML, CW, SW, SH, TOP, FOOT_Y = B.ML, B.CW, B.SW, B.SH, B.TOP, B.FOOT_Y
INK, BODY, MUTED, FAINT = B.INK, B.BODY, B.MUTED, B.FAINT
ACCENT, ACCENT_D, ACCENT_L = B.ACCENT, B.ACCENT_D, B.ACCENT_L
LINE, LINE2, CARD, SOFT, TINT = B.LINE, B.LINE2, B.CARD, B.SOFT, B.TINT
DARK, WHITE, BLUE, GREEN, PURPLE = B.DARK, B.WHITE, B.BLUE, B.GREEN, B.PURPLE
DARK_TX, DARK_MUT = B.DARK_TX, B.DARK_MUT
KR, MONO = B.KR, B.MONO
rect, text, header, footer_rule, title_block, banner = B.rect, B.text, B.header, B.footer_rule, B.title_block, B.banner
wcard, dcard, codebox, step_row = B.wcard, B.dcard, B.codebox, B.step_row
grid, lcard, minicard, checklist = L4.grid, L4.lcard, L4.minicard, L4.checklist
BLUE_D, GREEN_D, PURPLE_D, RED = L4.BLUE_D, L4.GREEN_D, L4.PURPLE_D, L4.RED
RED_SOFT, GREEN_SOFT = L4.RED_SOFT, L4.GREEN_SOFT

NOTES = {}
PAGES = []          # (slide, 구역 이름) — 마지막에 "구역  NN / 전체" 를 붙인다


def note(i, lines):
    NOTES[i] = "\n".join(lines)


def page(s, section):
    PAGES.append((s, section))


def table(s, x, y, w, cols, head, rows, row_h=0.40, head_h=0.40, size=10.5, first_bold=True, accent_col=None):
    """5차시 스타일 표 — 남색 머리글 + 흰 행 + 얇은 구분선"""
    rect(s, x, y, w, head_h, fill=DARK)
    cx = x
    for j, (cw, h) in enumerate(zip(cols, head)):
        text(s, cx + 0.16, y + 0.09, cw - 0.24, head_h - 0.12, [[(h, {"bold": True, "size": 10, "color": WHITE})]])
        cx += cw
    yy = y + head_h
    for i, row in enumerate(rows):
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
    """아래쪽 한 줄 강조 띠"""
    rect(s, ML, y, CW, h, fill=TINT, radius=0.10)
    text(s, ML + 0.3, y + 0.06, CW - 0.6, h - 0.10, runs, size=11.5, color=BODY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ================================================================ 슬라이드
def build():
    prs = B.new_deck()
    S = lambda: B.blank(prs)
    n = lambda: B.CUR["n"]

    # ------------------------------------------------------------ 1. 표지
    s = S()
    text(s, ML, 0.44, 5, 0.3, [[("A I   ×   B U I L D", {"bold": True, "size": 11, "color": ACCENT, "space": 1.3})]])
    text(s, SW - ML - 7.4, 0.46, 7.4, 0.3,
         [[("한국디지털미디어고등학교 · 6차시", {"bold": True, "size": 10, "color": FAINT, "space": 1.0})]],
         align=PP_ALIGN.RIGHT)
    text(s, ML, 1.70, 2.2, 1.2, [[("06", {"bold": True, "size": 60, "color": ACCENT_L})]])
    text(s, ML, 2.84, CW, 1.0, [[("미니프로젝트 발표 · GitHub 협업", {"bold": True, "size": 40, "color": INK})]])
    rect(s, ML, 3.94, 0.9, 0.05, fill=ACCENT)
    text(s, ML, 4.20, CW * 0.9, 0.5,
         [[("설계를 공유하고, 하나의 저장소에서 함께 만드는 팀 개발 프로세스", {"size": 17, "color": BODY})]])
    w, gap = grid(3, 0.22)
    for i, (h1, h2) in enumerate([("발표", "PRD로 팀의 설계를 공유한다"),
                                  ("리뷰", "범위 · 기술 · 보안 관점으로 검증한다"),
                                  ("협업", "브랜치 · PR · 리뷰 · 머지로 통합한다")]):
        minicard(s, ML + i * (w + gap), 5.20, w, 1.20, h1, h2, bar=ACCENT if i == 2 else ACCENT_L)
    footer_rule(s)
    text(s, ML, FOOT_Y + 0.12, 8, 0.26,
         [[("혼자 만드는 코드에서, 팀이 함께 관리하는 코드로", {"size": 9.5, "color": FAINT, "space": 0.8})]])
    note(n(), [
        "[90분 운영 개요]",
        "도입 5분 → 조별 PRD 발표 · 피드백 30분 → GitHub 협업 개념 10분 → 저장소 세팅 10분 → 협업 실습(이슈 · 브랜치 · PR · 리뷰 · 충돌) 25분 → 배포 · 회고 10분",
        "",
        "[대본]",
        "\"지난 시간에 조별로 미니프로젝트 주제를 정하고 PRD까지 썼죠. 오늘은 두 가지를 합니다.",
        "첫째, 그 PRD를 발표하고 다른 조에게 검증을 받습니다.",
        "둘째, 이제부터는 혼자가 아니라 팀이 같은 코드를 고칩니다. 그래서 GitHub에서 협업하는 방법을 배웁니다.",
        "오늘 끝나면 조마다 저장소가 하나 있고, 모든 조원이 PR을 한 번 이상 머지한 상태가 됩니다.\"",
    ])

    # ------------------------------------------------------------ 2. 90분 구성
    s = S()
    header(s)
    title_block(s, "90분 구성", "발표로 설계를 검증하고, 저장소로 협업을 시작합니다")
    blocks = [("0–5분", "도입", "목표 · 진행 방식", LINE2),
              ("5–35분", "발표", "조별 PRD · 피드백", ACCENT),
              ("35–45분", "개념", "GitHub Flow · 권한", ACCENT_L),
              ("45–55분", "세팅", "저장소 · 보호 · 클론", ACCENT_L),
              ("55–80분", "협업 실습", "이슈 · PR · 충돌", GREEN),
              ("80–90분", "배포 · 회고", "Pages · 체크", LINE2)]
    w, gap = grid(6, 0.12)
    for i, (t, h1, h2, bar) in enumerate(blocks):
        x = ML + i * (w + gap)
        rect(s, x, TOP, w, 1.30, fill=CARD, line=LINE, radius=0.05)
        rect(s, x, TOP, w, 0.065, fill=bar)
        text(s, x + 0.18, TOP + 0.20, w - 0.3, 0.26, [[(t, {"bold": True, "size": 9.5, "color": ACCENT_D})]])
        text(s, x + 0.18, TOP + 0.48, w - 0.3, 0.34, [[(h1, {"bold": True, "size": 14, "color": INK})]])
        text(s, x + 0.18, TOP + 0.86, w - 0.3, 0.30, [[(h2, {"size": 10, "color": MUTED})]])
    w2, gap2 = grid(2, 0.30)
    wcard(s, ML, 3.92, w2, 2.62, ACCENT, "오늘의 산출물", "팀 저장소 + 머지된 PR",
          ["피드백을 반영한 PRD.md · 이슈 목록",
           "조원 1인당 PR 1개 이상 머지",
           "GitHub Pages 배포 주소 1개"], strong=True)
    wcard(s, ML + w2 + gap2, 3.92, w2, 2.62, GREEN, "완료 조건", "main은 PR로만 바뀐다",
          ["main에 직접 push한 커밋 0개",
           "모든 PR에 리뷰 승인 1개 이상",
           "충돌 1회 이상 직접 해결"], label_color=GREEN_D)
    footer_rule(s); page(s, "도입")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"오늘 흐름입니다. 앞 30분은 발표, 뒤 50분은 GitHub 협업 실습입니다.",
        "산출물은 세 가지예요. 피드백을 반영한 PRD, 팀 저장소의 머지된 PR, 그리고 배포 주소.",
        "완료 조건에서 가장 중요한 건 첫 줄입니다. main 브랜치는 PR로만 바뀝니다. 직접 push는 0개.\"",
        "",
        "[운영]",
        "발표 순서와 타이머를 칠판에 미리 적어 둡니다. 조당 6분(발표 3분 · 피드백 2분 · 전환 1분).",
    ])

    # ------------------------------------------------------------ 3. PART · 발표
    s = S()
    B.part_slide(s, "P A R T   1", "미니프로젝트 발표", "PRD로 설계를 공유하고 검증받습니다", 1, 3)

    note(n(), [
        "[시간] 전환",
        "[운영] 첫 발표 조를 미리 정해 두고 바로 시작합니다. 발표 자료는 PRD.md 화면 그대로도 충분합니다.",
    ])

    # ------------------------------------------------------------ 4. 발표 형식
    s = S()
    header(s)
    title_block(s, "발표 형식", "조당 3분 · PRD의 결정 사항만 전달합니다")
    items = [("01", "한 문장 정의", "무엇을, 누구를 위해"),
             ("02", "문제와 사용자", "해결할 불편 · 대상 사용자"),
             ("03", "MUST 기능", "이번 버전 3~5개"),
             ("04", "화면 흐름", "시작 → 핵심 동작 → 결과"),
             ("05", "역할 분담", "조원별 담당 기능"),
             ("06", "기술 위험", "막힐 것 같은 지점 1개")]
    w, gap = grid(3, 0.22)
    for i, (num, h1, h2) in enumerate(items):
        x = ML + (i % 3) * (w + gap)
        y = TOP + (i // 3) * 1.42
        rect(s, x, y, w, 1.26, fill=CARD, line=LINE, radius=0.05)
        text(s, x + 0.24, y + 0.20, 0.6, 0.3, [[(num, {"bold": True, "size": 11, "color": ACCENT})]])
        text(s, x + 0.24, y + 0.50, w - 0.48, 0.34, [[(h1, {"bold": True, "size": 15, "color": INK})]])
        text(s, x + 0.24, y + 0.88, w - 0.48, 0.30, [[(h2, {"size": 11, "color": MUTED})]])
    bar_note(s, [[("발표 3분 · 피드백 2분 · 전환 1분  —  ", {}),
                  ("기능 나열보다 \"무엇을 뺐는지\"가 설계의 수준입니다", {"bold": True, "color": ACCENT_D})]],
             y=5.36, h=0.70)
    footer_rule(s); page(s, "발표")
    note(n(), [
        "[시간] 1분 (설명) → 이후 발표 진행",
        "",
        "[대본]",
        "\"발표는 3분입니다. PRD를 다 읽지 말고, 결정한 것만 말해 주세요. 여섯 가지 순서입니다.",
        "특히 MUST 기능은 3~5개로 말하세요. 그리고 '이번에 뺀 것'을 하나 이상 말하면 좋습니다. 무엇을 뺐는지가 설계 수준을 보여 줍니다.",
        "마지막 기술 위험은 솔직하게. 서버가 필요한지, 실시간 동기화가 필요한지, 외부 API가 필요한지.\"",
        "",
        "[운영]",
        "타이머를 화면에 띄웁니다. 3분이 지나면 바로 피드백으로 넘깁니다.",
        "발표 화면: PRD.md를 VS Code · 메모장 · Claude 화면 어느 것으로 띄워도 됩니다.",
    ])

    # ------------------------------------------------------------ 5. 피드백 방식
    s = S()
    header(s)
    title_block(s, "피드백 방식", "청중 조는 역할을 나눠 PRD를 검증합니다")
    w, gap = grid(3, 0.26)
    roles = [("범위 리뷰어", ACCENT, ACCENT_D, "MUST가 3분짜리 데모에 맞는가",
              ["MUST가 5개를 넘지 않는가", "빼도 되는 기능은 무엇인가", "완료 조건이 판정 가능한가"]),
             ("기술 리뷰어", BLUE, BLUE_D, "만들 수 있는 구조인가",
              ["서버 · 실시간 동기화가 필요한가", "브라우저만으로 가능한가", "가장 오래 걸릴 기능은 무엇인가"]),
             ("보안 리뷰어", RED, RED, "어디가 공격 표면인가",
              ["입력값이 화면에 그대로 출력되는가", "키 · 개인정보가 코드에 들어가는가", "점수 · 기록을 조작할 수 있는가"])]
    for i, (label, bar, lc, head, lines) in enumerate(roles):
        wcard(s, ML + i * (w + gap), TOP, w, 2.70, bar, label, head, lines, label_color=lc, size=11)
    codebox(s, ML, 5.20, CW, 1.12, "피드백 문장 형식",
            ["[역할] · [근거가 된 PRD 문장] → [제안]      예) 보안 · \"닉네임을 랭킹에 표시\" → 입력 길이 제한과 HTML 이스케이프 필요"])
    footer_rule(s); page(s, "발표")
    note(n(), [
        "[시간] 2분 (설명) · 이후 조마다 피드백 2분",
        "",
        "[대본]",
        "\"발표를 듣는 조는 세 역할로 나눕니다. 조 안에서 범위 · 기술 · 보안 리뷰어를 한 명 이상씩 정하세요.",
        "범위 리뷰어는 '너무 크다'를 찾습니다. 기술 리뷰어는 '만들 수 있나'를 봅니다.",
        "보안 리뷰어가 오늘 핵심입니다. 여러분 전공이죠. 입력값이 그대로 화면에 찍히면 XSS, 점수를 클라이언트에서 계산하면 조작 가능, 키가 코드에 있으면 유출.",
        "피드백은 반드시 PRD의 어떤 문장 때문인지 근거를 붙여 주세요.\"",
        "",
        "[운영]",
        "조마다 피드백 2개(서로 다른 역할)만 받습니다. 시간이 남으면 교사가 보안 관점 하나를 추가합니다.",
        "발표 조의 기록 담당은 받은 피드백을 그대로 적어 둡니다 → 다음 슬라이드에서 PRD에 반영.",
    ])

    # ------------------------------------------------------------ 6. 피드백 반영
    s = S()
    header(s)
    title_block(s, "피드백 반영", "받은 피드백을 PRD에 반영하고 결정을 기록합니다")
    cw_ = 7.05
    codebox(s, ML, TOP, cw_, 3.92, "PROMPT",
            ["PRD.md에 아래 피드백을 반영해 줘. 아직 다른 파일은 만들지 마.",
             "",
             "[받은 피드백 붙여넣기]",
             "",
             "1. 피드백마다 반영 / 보류 / 거절을 표시하고 이유를 한 줄로",
             "2. MUST는 5개 이하로 유지",
             "3. 보안 피드백은 '보안 요구사항' 항목을 새로 만들어 정리",
             "4. 바뀐 부분만 보여 주고, 내가 확인하면 저장해 줘"])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    lcard(s, rx, TOP, rw, 1.86, ACCENT, "판단 기준 · 모두 받지 않는다",
          ["데모 범위를 넘으면 보류", "근거 없는 제안은 거절"], size=11)
    lcard(s, rx, TOP + 2.02, rw, 1.90, GREEN, "결정 기록 · PRD 맨 아래",
          ["## 결정 기록 (6차시)", "피드백 · 결정 · 이유"], label_color=GREEN_D, size=11)
    bar_note(s, [[("PRD는 이 다음부터 ", {}), ("GitHub 이슈의 원본", {"bold": True, "color": ACCENT_D}),
                  ("이 됩니다", {})]])
    page(s, "발표")
    note(n(), [
        "[시간] 발표가 끝난 조부터 각자 진행 (다른 조 발표 중에는 하지 않음 · 발표 종료 후 3분)",
        "",
        "[대본]",
        "\"피드백을 전부 받을 필요는 없습니다. 반영 · 보류 · 거절 중 하나로 결정하고 이유를 남기는 게 중요합니다.",
        "보안 피드백은 PRD에 '보안 요구사항' 항목으로 따로 정리하세요. 예를 들어 '닉네임은 12자 이하, HTML은 이스케이프'.",
        "이렇게 정리한 PRD의 MUST 기능이 잠시 후 GitHub 이슈가 됩니다.\"",
        "",
        "[운영]",
        "PRD.md는 아직 각자 노트북에 있습니다. 저장소를 만든 뒤 Owner가 올립니다(슬라이드 '저장소 생성').",
    ])

    # ------------------------------------------------------------ 7. PART · GitHub 협업
    s = S()
    B.part_slide(s, "P A R T   2", "GitHub 협업", "하나의 저장소, 여러 명의 개발자", 2, 3)

    note(n(), [
        "[시간] 전환",
        "[대본] \"이제 PRD를 코드로 옮길 차례입니다. 그런데 조원 네다섯 명이 같은 게임을 고치면 무슨 일이 생길까요?\"",
    ])

    # ------------------------------------------------------------ 8. 혼자 Git vs 팀 GitHub
    s = S()
    header(s)
    title_block(s, "Git과 GitHub", "4차시의 로컬 버전 관리에서 팀 단위 원격 협업으로")
    table(s, ML, TOP, CW, [2.6, 4.6, 4.67],
          ["", "Git · 4차시", "GitHub · 오늘"],
          [["저장 위치", "내 노트북의 .git 폴더", "원격 저장소 (github.com) + 각자의 복제본"],
           ["공유", "백업 파일을 전달", "push · pull로 동기화"],
           ["변경 반영", "내가 커밋하면 끝", "PR → 리뷰 → 머지"],
           ["충돌", "생기지 않음", "같은 줄을 동시에 고치면 발생"],
           ["기록", "내 커밋만", "누가 · 언제 · 왜 바꿨는지 팀 전체"],
           ["권한", "나 혼자", "Owner · Collaborator · 보호 규칙"]],
          row_h=0.48, accent_col=None)
    bar_note(s, [[("Git", {"bold": True, "color": ACCENT_D}), ("은 버전 관리 도구,  ", {}),
                  ("GitHub", {"bold": True, "color": ACCENT_D}), ("은 그 저장소를 공유하고 협업하는 서비스", {})]], y=5.92)
    footer_rule(s); page(s, "개념")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"4차시에 커밋으로 세이브했죠. 그건 내 노트북 안의 기록이었습니다. 오늘은 그 저장소를 github.com에 올려서 팀이 같이 씁니다.",
        "가장 큰 차이는 '변경 반영' 줄입니다. 혼자일 때는 커밋하면 끝이지만, 팀에서는 PR을 열고, 다른 사람이 리뷰하고, 승인되면 머지합니다.",
        "그리고 충돌. 두 사람이 같은 줄을 동시에 고치면 Git이 알아서 합치지 못합니다. 오늘 일부러 한 번 만들어 보고 해결합니다.\"",
    ])

    # ------------------------------------------------------------ 9. 핵심 용어
    s = S()
    header(s)
    title_block(s, "협업 핵심 용어", "오늘 실습에서 그대로 쓰는 말")
    left = [["Repository", "프로젝트 파일과 변경 기록 전체"],
            ["Clone", "원격 저장소를 내 컴퓨터에 복제"],
            ["Branch", "main과 분리된 작업 공간"],
            ["Commit", "변경 내용을 하나의 버전으로 기록"],
            ["Push", "내 커밋을 원격 저장소로 전송"],
            ["Pull", "원격의 최신 커밋을 내 컴퓨터로"]]
    right = [["Issue", "할 일 · 버그 · 제안 단위의 작업 카드"],
             ["Pull Request", "내 브랜치를 main에 합쳐 달라는 요청"],
             ["Review", "PR의 변경 내용을 검토 · 승인 · 수정 요청"],
             ["Merge", "승인된 브랜치를 main에 통합"],
             ["Conflict", "같은 위치의 서로 다른 변경 — 수동 해결"],
             ["Protected branch", "규칙을 통과해야만 바뀌는 브랜치"]]
    w, gap = grid(2, 0.26)
    table(s, ML, TOP, w, [1.95, w - 1.95], ["용어", "의미"], left, row_h=0.50)
    table(s, ML + w + gap, TOP, w, [1.95, w - 1.95], ["용어", "의미"], right, row_h=0.50)
    footer_rule(s); page(s, "개념")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"용어는 열두 개입니다. 왼쪽은 4차시에 써 본 것과 비슷하고, 오른쪽이 오늘 새로 쓰는 것들입니다.",
        "Pull Request를 줄여서 PR이라고 부릅니다. '내 작업을 main에 합쳐 주세요'라는 요청서예요.",
        "Protected branch는 보호된 브랜치입니다. 오늘 main을 보호해서 리뷰 없이는 아무도 못 바꾸게 할 겁니다. Owner도요.\"",
    ])

    # ------------------------------------------------------------ 10. GitHub Flow
    s = S()
    header(s)
    title_block(s, "협업 흐름 · GitHub Flow", "모든 변경은 이슈에서 시작해 PR로 main에 들어갑니다")
    step_row(s, TOP, [("1", "Issue", "할 일 등록 · 담당자 지정"),
                      ("2", "Branch", "feat/기능이름 생성"),
                      ("3", "Commit", "작게 · 자주 · 의미 있게"),
                      ("4", "Push", "원격에 브랜치 올리기")], h=1.40)
    step_row(s, TOP + 1.62, [("5", "Pull Request", "변경 설명 · Closes #번호"),
                             ("6", "Review", "검토 · 수정 요청 · 승인"),
                             ("7", "Merge", "main에 통합 · 브랜치 삭제"),
                             ("8", "Pull", "모두 main 최신화")], h=1.40)
    bar_note(s, [[("main은 언제나 ", {}), ("실행 가능한 상태", {"bold": True, "color": ACCENT_D}),
                  ("  ·  작업은 브랜치에서  ·  합치는 건 PR로", {})]], y=5.56, h=0.62)
    footer_rule(s); page(s, "개념")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"GitHub Flow라는 흐름입니다. 실제 회사에서도 이 방식을 많이 씁니다. 여덟 단계지만 원칙은 하나예요. main은 언제나 실행 가능한 상태.",
        "그래서 아무도 main에서 직접 작업하지 않습니다. 이슈를 하나 잡고, 그 이슈용 브랜치를 만들고, 거기서 커밋합니다.",
        "다 되면 PR을 열고, 다른 조원이 리뷰해서 승인하면 머지합니다. 머지 후에는 모두가 main을 pull해서 최신으로 맞춥니다.",
        "PR 설명에 'Closes #3'이라고 쓰면, 머지될 때 3번 이슈가 자동으로 닫힙니다.\"",
    ])

    # ------------------------------------------------------------ 11. 역할과 권한
    s = S()
    header(s)
    title_block(s, "역할과 권한", "최소 권한 · main 보호 · 리뷰 필수")
    w, gap = grid(3, 0.26)
    wcard(s, ML, TOP, w, 2.80, ACCENT, "OWNER · 1명", "저장소 관리",
          ["저장소 생성 · 조원 초대", "main 보호 규칙 설정", "Pages 배포 설정"], strong=True, size=11)
    wcard(s, ML + w + gap, TOP, w, 2.80, BLUE, "COLLABORATOR", "기능 개발",
          ["이슈 담당 · 브랜치 작업", "PR 작성 · 다른 PR 리뷰", "main 직접 push 불가"], label_color=BLUE_D, size=11)
    wcard(s, ML + 2 * (w + gap), TOP, w, 2.80, GREEN, "REVIEWER · 매 PR", "품질 · 보안 검증",
          ["PRD 범위 · 실행 확인", "비밀값 · 위험 코드 확인", "승인 또는 수정 요청"], label_color=GREEN_D, size=11)
    codebox(s, ML, 5.24, CW, 1.14, "조 안의 규칙",
            ["내 PR은 내가 머지하지 않는다  ·  리뷰어는 PR마다 바꾼다  ·  Owner도 PR로만 main을 바꾼다"])
    footer_rule(s); page(s, "개념")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"조에서 Owner 한 명을 정하세요. 저장소를 만들고 조원을 초대하는 사람입니다. 나머지는 Collaborator예요.",
        "Reviewer는 고정된 사람이 아니라 PR마다 바뀌는 역할입니다. 내 PR은 다른 사람이 리뷰하고, 나는 다른 사람 PR을 리뷰합니다.",
        "해킹방어과니까 이건 익숙할 거예요. 최소 권한 원칙입니다. 리뷰는 품질 관리이면서 동시에 보안 통제입니다. 한 사람이 혼자 main을 바꿀 수 없게 만드는 거죠.\"",
        "",
        "[운영]",
        "Owner는 가능하면 GitHub를 써 본 학생으로 정하게 합니다.",
    ])

    # ------------------------------------------------------------ 12. 팀 저장소 구조
    s = S()
    header(s)
    title_block(s, "팀 저장소 구조", "기능별로 파일을 나누면 충돌이 줄어듭니다")
    cw_ = 5.80
    codebox(s, ML, TOP, cw_, 3.92, "폴더 구조 예시",
            ["mini-project/",
             "├── index.html       ← 화면 뼈대 · 담당 1명",
             "├── css/style.css    ← 디자인 토큰 (DESIGN.md)",
             "├── js/",
             "│   ├── main.js      ← 시작 · 화면 전환",
             "│   ├── game.js      ← 핵심 로직",
             "│   └── score.js     ← 점수 · 기록",
             "├── PRD.md",
             "├── DESIGN.md",
             "├── CLAUDE.md        ← 팀 공통 규칙",
             "├── README.md",
             "└── .gitignore"])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    codebox(s, rx, TOP, rw, 2.40, "CLAUDE.md · 팀 규칙",
            ["# 팀 규칙",
             "- 담당 이슈의 파일만 수정한다",
             "- index.html 구조 변경은 담당자와 상의",
             "- 외부 라이브러리 · 키 · 서버 금지",
             "- 커밋 전에 바꾼 파일 목록을 보여 준다",
             "- 디자인은 @DESIGN.md를 따른다"])
    wcard(s, rx, TOP + 2.54, rw, 1.38, GREEN, "왜 나누나", None,
          ["한 파일을 여럿이 고칠수록 충돌이 늘어난다"], label_color=GREEN_D, size=11)
    bar_note(s, [[("4차시의 \"파일 하나\" 규칙은 개인 작업용  —  ", {}),
                  ("팀 작업은 기능별 파일 분리", {"bold": True, "color": ACCENT_D})]])
    page(s, "개념")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"4차시에는 index.html 하나로 만들라고 했죠. 혼자 할 때는 그게 제일 편합니다. 그런데 다섯 명이 한 파일을 고치면 충돌이 계속 납니다.",
        "그래서 팀 프로젝트는 기능별로 파일을 나눕니다. 게임 로직, 점수, 화면 전환. 이슈 하나가 파일 하나에 대응하면 충돌이 거의 없습니다.",
        "CLAUDE.md에 팀 규칙을 넣어 두면, 조원 모두의 Claude가 같은 규칙으로 일합니다. 특히 '담당 이슈의 파일만 수정한다' 한 줄이 충돌을 많이 줄여 줍니다.",
        "빌드 도구 없이 script 태그로 js 파일을 불러오면 됩니다. GitHub Pages에서 그대로 실행됩니다.\"",
        "",
        "[운영]",
        "기존 단일 HTML 작품을 쓰는 조는 'index.html을 기능별 js 파일로 분리해 줘. 동작은 바꾸지 마'로 먼저 분리합니다(Owner가 1회).",
    ])

    # ------------------------------------------------------------ 13. PART · 저장소 세팅
    s = S()
    B.part_slide(s, "P A R T   3", "저장소 세팅 · 협업 실습", "Owner 세팅 10분 → 전원 실습 25분", 3, 3)

    note(n(), [
        "[시간] 전환",
        "[운영] Owner는 다음 슬라이드를 바로 진행합니다. 나머지 조원은 GitHub 로그인과 '이메일 비공개' 설정(슬라이드 '클론과 계정 설정')을 먼저 합니다.",
    ])

    # ------------------------------------------------------------ 14. STEP 1 저장소 생성
    s = S()
    header(s)
    title_block(s, "STEP 1 · 저장소 생성", "Owner · 5분 · GitHub 웹에서 진행합니다")
    step_row(s, TOP, [("①", "New repository", "github.com 오른쪽 위 + → New repository"),
                      ("②", "설정", "이름 · Public · README · .gitignore(Node)"),
                      ("③", "파일 업로드", "PRD.md · DESIGN.md · 작품 파일"),
                      ("④", "조원 초대", "Settings → Collaborators → Add people")], h=1.70)
    w, gap = grid(2, 0.30)
    lcard(s, ML, 4.24, w, 2.08, ACCENT, "저장소 이름",
          ["영문 소문자 · 하이픈  예) dimigo-6-escape", "Public — Pages 배포와 보호 규칙에 필요"], size=11)
    lcard(s, ML + w + gap, 4.24, w, 2.08, BLUE, "조원이 할 일",
          ["메일 또는 github.com/notifications에서 초대 수락", "수락 전에는 push 권한이 없다 (403)"],
          label_color=BLUE_D, size=11)
    footer_rule(s); page(s, "세팅")
    note(n(), [
        "[시간] 5분",
        "",
        "[대본 · Owner에게]",
        "\"Owner는 github.com에서 오른쪽 위 더하기 버튼, New repository를 누르세요. 이름은 영문 소문자와 하이픈으로.",
        "Public으로 만듭니다. 무료 계정에서 Pages 배포와 브랜치 보호 규칙을 쓰려면 Public이 편합니다. 그래서 코드에 개인정보나 키가 없어야 합니다.",
        "README와 .gitignore를 체크하고 만드세요. 그다음 Add file → Upload files로 PRD.md, DESIGN.md, 작품 파일을 올립니다.",
        "마지막으로 Settings → Collaborators에서 조원의 GitHub 아이디로 초대합니다.\"",
        "",
        "[대본 · 조원에게]",
        "\"초대 메일이 오면 반드시 수락하세요. 수락하지 않으면 나중에 push할 때 403 에러가 납니다.\"",
        "",
        "[운영]",
        "GitHub 화면 메뉴 이름은 개편에 따라 조금 다를 수 있습니다. 수업 전에 교사 계정으로 한 번 따라 해 보세요.",
        "Private 저장소로 하고 싶은 조: 무료 플랜에서는 일부 보호 규칙과 Pages 사용에 제약이 있을 수 있음을 안내합니다.",
    ])

    # ------------------------------------------------------------ 15. STEP 2 main 보호
    s = S()
    header(s)
    title_block(s, "STEP 2 · main 보호", "Owner · 3분 · 리뷰 없이는 아무도 main을 바꿀 수 없게")
    cw_ = 6.40
    table(s, ML, TOP, cw_, [2.75, 3.65], ["설정", "값"],
          [["메뉴", "Settings → Rules (또는 Branches)"],
           ["대상 브랜치", "main (Default branch)"],
           ["Require a pull request", "켬 · 승인 1개"],
           ["Block force pushes", "켬"],
           ["Restrict deletions", "켬"],
           ["예외 (Bypass)", "없음 — Owner 포함"]], row_h=0.48)
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    wcard(s, rx, TOP, rw, 1.80, RED, "확인 실험", "main에 직접 push",
          ["거부되면 설정 성공"], label_color=RED, size=11)
    wcard(s, rx, TOP + 1.96, rw, 1.42, GREEN, "막는 것", None,
          ["실수 · 리뷰 우회 · 기록 덮어쓰기"], label_color=GREEN_D, size=11)
    bar_note(s, [[("보호 규칙 = ", {}), ("코드에 대한 접근 통제", {"bold": True, "color": ACCENT_D}),
                  ("  ·  리뷰 = 변경에 대한 2인 승인", {})]])
    page(s, "세팅")
    note(n(), [
        "[시간] 3분",
        "",
        "[대본]",
        "\"이제 main을 보호합니다. Settings에서 Rules 또는 Branches 메뉴로 들어가 main 브랜치에 규칙을 겁니다.",
        "Pull request 필수, 승인 1개. Force push 금지, 삭제 금지. 그리고 예외 없음. Owner도 예외가 아닙니다.",
        "설정이 끝나면 확인 실험을 하세요. main에서 아무 파일이나 고쳐서 직접 push해 보세요. 거부당하면 성공입니다.",
        "보안 수업에서 배우는 접근 통제와 같습니다. 보호 규칙은 누가 무엇을 바꿀 수 있는지 정하는 거고, 리뷰는 한 사람이 혼자 바꾸지 못하게 하는 2인 승인이에요.\"",
        "",
        "[운영]",
        "GitHub는 브랜치 보호(Branch protection rules)와 규칙 세트(Rulesets) 두 방식을 제공합니다. 화면에 보이는 쪽을 쓰면 됩니다.",
        "Force push: git push --force로 원격 기록을 덮어쓰는 것. 협업 저장소에서는 금지합니다.",
    ])

    # ------------------------------------------------------------ 15-2. 작업 승인 모드
    s = S()
    header(s)
    title_block(s, "작업 승인 모드", "Claude가 묻지 않고 처리하는 범위 · 2026년 10월 기준")
    table(s, ML, TOP, CW, [2.75, 4.10, 3.30, 1.72],
          ["모드", "자동 처리", "확인 요청", "수업"],
          [["Manual · 매뉴얼", "파일 읽기", "수정 · 명령 · 네트워크", "기본 학습"],
           ["Plan · 플랜", "읽기 · 탐색용 명령", "승인 전 파일 수정 차단", "필요할 때"],
           ["Accept edits · 억셉트 에딧츠", "파일 수정 · 삭제 · 이동 · 복사", "그 외 명령 · 폴더 밖 경로", "오늘 기본"],
           ["Auto · 오토", "대부분 · AI 안전 검사", "위험 판단 시에만", "사용 안 함"],
           ["Bypass permissions", "거의 전부", "거의 없음", "금지"]],
          row_h=0.48, accent_col=3)
    w, gap = grid(2, 0.30)
    lcard(s, ML, 5.24, w, 1.42, ACCENT, "변경 위치",
          ["앱: 입력창 아래 모드 선택 · Ctrl+Shift+M  ·  VS Code: 모드 표시 클릭"], size=10.5)
    lcard(s, ML + w + gap, 5.24, w, 1.42, RED, "주의",
          ["VS Code는 Auto로 시작 — 작업 전 모드부터 확인"], label_color=RED, size=10.5)
    footer_rule(s); page(s, "세팅")
    note(n(), [
        "[시간] 2분 (세팅 시간 안에서)",
        "",
        "[대본]",
        "\"오늘부터 Claude가 git 명령을 실행합니다. 그래서 Claude가 어디까지 묻지 않고 처리하는지 먼저 정합니다. 모드는 다섯 가지예요.",
        "매뉴얼은 읽기만 자동이고 나머지는 전부 물어봅니다. 플랜은 계획을 승인하기 전까지 파일 수정이 기술적으로 막힙니다.",
        "오늘 기본은 억셉트 에딧츠입니다. 파일 수정과 삭제 · 이동은 자동이지만, git push나 설치 같은 그 밖의 명령은 반드시 물어봅니다.",
        "오토는 AI 안전 검사 모델이 판단해서 대부분 실행합니다. 바이패스는 거의 묻지 않습니다. 둘 다 오늘은 쓰지 않습니다.",
        "VS Code를 쓰는 사람은 주의하세요. 최근 버전은 오토로 시작합니다. 작업 전에 모드부터 확인하세요.\"",
        "",
        "[운영]",
        "'모드가 Accept edits인 사람 손?'으로 전원 확인 후 다음 슬라이드로 넘어갑니다.",
        "모드 이름과 위치는 앱 버전에 따라 다를 수 있습니다. 수업 전에 교사 앱에서 확인하세요.",
        "앱에서는 Shift+Tab으로 모드가 바뀌지 않습니다. 모드 선택 메뉴 또는 Ctrl+Shift+M(맥은 Cmd+Shift+M)을 씁니다.",
    ])

    # ------------------------------------------------------------ 15-3. Accept edits vs Auto
    s = S()
    header(s)
    title_block(s, "Accept edits와 Auto의 차이", "수정 속도는 같고, 명령을 누가 승인하는지가 다릅니다")
    cw_ = 7.30
    table(s, ML, TOP, cw_, [3.10, 2.10, 2.10], ["작업", "Accept edits", "Auto"],
          [["index.html 수정", "자동", "자동"],
           ["폴더 안 파일 삭제 · 이동", "자동", "자동"],
           ["git commit · push", "확인", "대부분 자동"],
           ["설치 · 스크립트 실행", "확인", "대부분 자동"],
           ["강제 push · 외부 스크립트 실행", "확인", "차단 또는 확인"]],
          row_h=0.52)
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    wcard(s, rx, TOP, rw, 2.50, ACCENT, "오늘의 기준", "Accept edits + 계획 먼저",
          ["요청문에 \"계획을 먼저 보여 줘\"", "push · 설치는 직접 승인", "잘 되면 커밋"], strong=True, size=10.5)
    lcard(s, rx, TOP + 2.62, rw, 1.32, BLUE, "Plan 모드를 쓰는 때",
          ["큰 구조 변경 · 다른 조 코드 분석"], label_color=BLUE_D, size=10.5)
    bar_note(s, [[("요청문의 \"계획 먼저\" = ", {}), ("정책", {"bold": True, "color": ACCENT_D}),
                  ("   ·   Plan 모드 = ", {}), ("강제 통제", {"bold": True, "color": ACCENT_D}),
                  ("   ·   삭제 · 이동도 자동이므로 커밋이 안전망", {})]])
    page(s, "세팅")
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"억셉트 에딧츠와 오토는 코드 수정 속도가 같습니다. 둘 다 묻지 않고 고쳐요. 차이는 명령입니다.",
        "억셉트 에딧츠에서는 git push, 설치, 스크립트 실행을 여러분이 승인합니다. 오토에서는 AI가 위험하지 않다고 판단하면 그냥 실행합니다.",
        "오늘은 push 하나가 팀 전체에 영향을 줍니다. 그래서 무엇이 실행되는지 직접 보고 승인하는 억셉트 에딧츠로 합니다.",
        "플랜 모드는 필요할 때만 씁니다. 요청문에 '계획 먼저 보여 줘'라고 쓰는 건 AI에게 부탁하는 정책이고, 플랜 모드는 수정 자체를 막는 강제 통제예요. 다른 조 코드를 분석할 때처럼 절대 바뀌면 안 되는 상황에서 씁니다.",
        "마지막으로, 억셉트 에딧츠는 파일 삭제와 이동도 묻지 않습니다. 그래서 잘 될 때마다 커밋하는 습관이 안전망입니다.\"",
        "",
        "[운영]",
        "흐름에 익숙한 학생의 개인 작업(9 · 10회기 폴리싱)에서만 Auto를 허용할지 교사가 결정합니다. Bypass permissions는 수업 전체에서 금지합니다.",
        "어떤 모드에서도 settings.json의 금지(deny) 규칙이 우선합니다.",
    ])

    # ------------------------------------------------------------ 16. STEP 3 클론 · 계정
    s = S()
    header(s)
    title_block(s, "STEP 3 · 클론과 계정 설정", "전원 · 5분 · 빈 폴더에서 Claude Code로 진행합니다")
    cw_ = 7.05
    codebox(s, ML, TOP, cw_, 3.92, "PROMPT",
            ["이 폴더에 아래 저장소를 클론해 줘. 현재 폴더에 바로.",
             "https://github.com/[Owner아이디]/[저장소이름].git",
             "",
             "그다음:",
             "1. git 작성자 이름을 [GitHub 아이디]로 이 저장소에만 설정",
             "2. 이메일은 [GitHub noreply 이메일]로 설정",
             "3. 폴더 구조와 브랜치 목록을 보여 줘",
             "4. 아직 아무 파일도 수정하지 마"])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    lcard(s, rx, TOP, rw, 2.10, ACCENT, "이메일 비공개 · noreply",
          ["GitHub → Settings → Emails", "Keep my email private 켜기"], size=10.5)
    lcard(s, rx, TOP + 2.24, rw, 1.68, BLUE, "첫 push 때",
          ["브라우저 로그인 창 · 1회"], label_color=BLUE_D, size=10.5)
    bar_note(s, [[("커밋의 작성자 이메일은 ", {}), ("공개 저장소에 영구히 남습니다", {"bold": True, "color": ACCENT_D})]])
    page(s, "세팅")
    note(n(), [
        "[시간] 5분 · 모드를 Accept edits로 맞춘 뒤 시작",
        "",
        "[대본]",
        "\"모두 바탕화면에 새 빈 폴더를 만들고 Claude 앱 Code 탭에서 그 폴더를 여세요. 그리고 이 프롬프트를 보냅니다.",
        "저장소 주소는 Owner가 GitHub 저장소 화면의 초록색 Code 버튼에서 복사해 단톡방에 올려 주세요.",
        "중요한 건 이메일입니다. 커밋에는 작성자 이메일이 들어가고, 공개 저장소에서는 누구나 볼 수 있습니다. 한 번 올라가면 지우기 어렵습니다.",
        "GitHub Settings → Emails에서 'Keep my email private'을 켜면 noreply 주소가 보입니다. 숫자+아이디@users.noreply.github.com 형태예요. 그걸 쓰세요.\"",
        "",
        "[운영]",
        "처음 push할 때 Git이 브라우저 로그인 창을 띄웁니다(Git for Windows의 로그인 연동). 학생이 직접 로그인합니다. 교사가 대신 입력하지 않습니다.",
        "로그인 창이 뜨지 않거나 실패하면: GitHub 웹에서 직접 파일을 수정하고 PR을 여는 방식으로 실습을 대체할 수 있습니다.",
        "Claude가 폴더를 다시 선택하라고 하면, 클론된 폴더를 Select folder로 다시 엽니다.",
    ])

    # ------------------------------------------------------------ 17. 이슈로 일 나누기
    s = S()
    header(s)
    title_block(s, "협업 ① · 이슈로 일 나누기", "PRD의 MUST 기능을 이슈 카드로 바꿉니다")
    cw_ = 6.20
    codebox(s, ML, TOP, cw_, 3.92, "PROMPT",
            ["PRD.md의 MUST 기능을 GitHub 이슈로 쓸 수 있게",
             "정리해 줘. 파일은 만들지 마.",
             "",
             "이슈마다:",
             "- 제목: [feat] 기능 이름",
             "- 설명: 무엇을 · 왜 (PRD 근거 문장)",
             "- 완료 조건: 체크박스 2~3개",
             "- 예상 수정 파일",
             "- 보안 확인 항목 (있으면)"])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    table(s, rx, TOP, rw, [1.70, rw - 1.70], ["이슈 화면", "입력"],
          [["Title", "[feat] 점수 표시"],
           ["Description", "Claude가 정리한 내용"],
           ["Assignees", "담당 조원 1명"],
           ["Labels", "feat · bug · docs"]], row_h=0.50)
    text(s, rx, TOP + 2.62, rw, 0.30, [[("1인 1이슈 이상 · 담당자 없는 이슈 0개", {"bold": True, "size": 11, "color": ACCENT_D})]])
    bar_note(s, [[("이슈 번호가 이후 브랜치 · 커밋 · PR을 잇는 ", {}), ("추적 키", {"bold": True, "color": ACCENT_D})]])
    page(s, "협업 실습")
    note(n(), [
        "[시간] 5분",
        "",
        "[대본]",
        "\"이제 일을 나눕니다. PRD의 MUST 기능 하나가 이슈 하나예요. Claude에게 이슈 형태로 정리해 달라고 하고, 그 내용을 GitHub 저장소의 Issues 탭 → New issue에 붙여 넣으세요.",
        "Assignees에 담당자를 꼭 지정합니다. 한 사람이 최소 하나씩.",
        "이슈를 만들면 번호가 붙습니다. #1, #2. 이 번호가 브랜치 이름, 커밋 메시지, PR에 계속 따라다닙니다. 나중에 '이 코드는 왜 들어왔지?'를 추적하는 키예요.\"",
        "",
        "[운영]",
        "이슈 등록은 조원이 나눠서 합니다(한 명이 다 하지 않게). '예상 수정 파일'이 겹치는 이슈는 충돌 위험이 크다고 미리 짚어 줍니다.",
    ])

    # ------------------------------------------------------------ 18. 브랜치 · 커밋 · 푸시
    s = S()
    header(s)
    title_block(s, "협업 ② · 브랜치 → 커밋 → 푸시", "내 이슈는 내 브랜치에서만 작업합니다")
    cw_ = 7.05
    codebox(s, ML, TOP, cw_, 3.92, "PROMPT · 순서대로",
            ["[시작]  main을 최신으로 pull하고,",
             "        feat/3-score 브랜치를 만들어 이동해 줘.",
             "",
             "[작업]  이슈 #3을 구현해 줘. 담당 파일만 수정.",
             "        계획을 먼저 보여 줘.",
             "",
             "[기록]  바뀐 파일 목록을 보여 주고 커밋해 줘.",
             "        메시지: feat: 점수 표시 (#3)",
             "",
             "[전송]  이 브랜치를 원격에 push해 줘."])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    table(s, rx, TOP, rw, [1.25, rw - 1.25], ["접두어", "용도"],
          [["feat/", "새 기능"], ["fix/", "버그 수정"], ["docs/", "문서 · README"], ["style/", "디자인만 변경"]],
          row_h=0.46)
    lcard(s, rx, TOP + 2.44, rw, 1.48, GREEN, "커밋 단위",
          ["실행되는 상태로 · 작게 · 자주"], label_color=GREEN_D, size=11)
    bar_note(s, [[("브랜치 이름 = ", {}), ("접두어/이슈번호-내용", {"bold": True, "color": ACCENT_D}),
                  ("   ·   커밋 메시지 = 타입: 요약 (#이슈번호)", {})]])
    page(s, "협업 실습")
    note(n(), [
        "[시간] 7분 (구현 시간 포함 · 오늘은 작게)",
        "",
        "[대본]",
        "\"작업은 항상 main 최신화부터 시작합니다. 그다음 내 이슈용 브랜치를 만듭니다. 이름은 feat/3-score처럼 접두어, 이슈 번호, 내용.",
        "오늘은 기능을 다 만들 필요 없습니다. 실습 목적은 흐름을 익히는 거예요. 화면에 점수 칸 하나 추가하는 정도로 작게 하세요.",
        "커밋하기 전에 꼭 '바뀐 파일 목록'을 보세요. 내 담당이 아닌 파일이 바뀌었으면 커밋하지 말고 되돌립니다.",
        "커밋 메시지는 타입, 요약, 이슈 번호. 이 형식을 Conventional Commits라고 부릅니다.\"",
        "",
        "[운영]",
        "첫 push에서 브라우저 로그인 창이 뜹니다. 학생이 직접 로그인하게 하세요.",
        "push가 거부되면(rejected): 대부분 초대 미수락(403) 또는 main에서 작업한 경우입니다. 트러블슈팅 슬라이드 참고.",
    ])

    # ------------------------------------------------------------ 19. Pull Request
    s = S()
    header(s)
    title_block(s, "협업 ③ · Pull Request", "변경의 목적 · 확인 방법 · 위험을 설명합니다")
    cw_ = 6.20
    codebox(s, ML, TOP, cw_, 3.92, "PROMPT",
            ["현재 브랜치와 main의 차이를 보고",
             "PR 설명을 아래 형식으로 써 줘.",
             "",
             "## 무엇을",
             "## 왜 (이슈 · PRD 근거)",
             "## 확인 방법 (재현 순서)",
             "## 스크린샷",
             "## 보안 · 위험 요소",
             "Closes #3"])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    step_x = rx
    rows = [("①", "Compare & pull request", "push 직후 저장소 화면 노란 배너"),
            ("②", "base: main ← compare: 내 브랜치", "방향 확인"),
            ("③", "Reviewers", "조원 1명 지정"),
            ("④", "Create pull request", "Draft는 작업 중일 때")]
    for i, (num, h1, h2) in enumerate(rows):
        yy = TOP + i * 0.98
        rect(s, step_x, yy, rw, 0.86, fill=CARD, line=LINE, radius=0.06)
        text(s, step_x + 0.20, yy + 0.12, 0.4, 0.3, [[(num, {"bold": True, "size": 12, "color": ACCENT})]])
        text(s, step_x + 0.62, yy + 0.10, rw - 0.80, 0.32, [[(h1, {"bold": True, "size": 12.5, "color": INK})]])
        text(s, step_x + 0.62, yy + 0.46, rw - 0.80, 0.30, [[(h2, {"size": 10.5, "color": MUTED})]])
    bar_note(s, [[("좋은 PR = ", {}), ("리뷰어가 코드를 실행하지 않고도 의도를 이해하는 설명", {"bold": True, "color": ACCENT_D})]])
    page(s, "협업 실습")
    note(n(), [
        "[시간] 4분",
        "",
        "[대본]",
        "\"push하면 GitHub 저장소 화면에 노란 배너가 뜹니다. Compare & pull request를 누르세요.",
        "방향이 중요합니다. base는 main, compare는 내 브랜치. 반대로 되어 있으면 안 됩니다.",
        "설명은 Claude에게 쓰게 하세요. 브랜치와 main의 차이를 보고 형식에 맞춰 써 줍니다. 단, 확인 방법과 보안 항목은 내가 직접 검토해서 고칩니다.",
        "마지막 줄 Closes #3. 머지되면 3번 이슈가 자동으로 닫힙니다.",
        "아직 작업 중이면 Create draft pull request로 열어도 됩니다. 리뷰를 미리 받을 수 있어요.\"",
    ])

    # ------------------------------------------------------------ 20. 코드 리뷰
    s = S()
    header(s)
    title_block(s, "협업 ④ · 코드 리뷰와 머지", "리뷰어는 실행하고, 읽고, 판정합니다")
    w, gap = grid(2, 0.26)
    table(s, ML, TOP, w, [0.55, w - 0.55], ["", "리뷰 체크리스트"],
          [["1", "PR 브랜치를 받아 직접 실행했다"],
           ["2", "이슈 · PRD 범위 안의 변경이다"],
           ["3", "담당 외 파일이 바뀌지 않았다"],
           ["4", "키 · 토큰 · 개인정보가 없다"],
           ["5", "입력값을 그대로 HTML에 넣지 않는다"],
           ["6", "기존 기능이 그대로 동작한다"]], row_h=0.50, first_bold=True)
    rx = ML + w + gap
    table(s, rx, TOP, w, [1.55, w - 1.55], ["판정", "언제"],
          [["Approve", "체크리스트 통과"],
           ["Comment", "질문 · 사소한 제안"],
           ["Request changes", "기능 손상 · 보안 문제"]], row_h=0.50)
    codebox(s, rx, TOP + 2.20, w, 1.52, "리뷰어 PROMPT",
            ["PR #5 브랜치를 받아서 실행 방법을 알려 주고,",
             "변경 코드의 보안 위험을 점검해 줘. 수정하지 마."])
    bar_note(s, [[("승인 후 머지 → 브랜치 삭제 → ", {}), ("전원 main pull", {"bold": True, "color": ACCENT_D}),
                  ("   ·   내 PR은 내가 머지하지 않는다", {})]], y=6.14)
    footer_rule(s); page(s, "협업 실습")
    note(n(), [
        "[시간] 5분",
        "",
        "[대본]",
        "\"리뷰는 코드를 눈으로만 보는 게 아닙니다. PR 브랜치를 받아서 직접 실행해 보세요. Claude에게 'PR #5 브랜치를 받아 실행 방법을 알려 줘'라고 하면 됩니다.",
        "체크리스트 4번과 5번이 보안 항목입니다. 키나 토큰이 들어갔는지, 사용자 입력이 innerHTML로 그대로 들어가는지. 여러분이 찾을 수 있는 대표적인 취약점이에요.",
        "판정은 세 가지. 통과면 Approve, 질문이면 Comment, 고쳐야 하면 Request changes.",
        "승인되면 Merge pull request를 누르고, Delete branch로 브랜치를 지웁니다. 그리고 조원 전원이 main을 pull합니다.\"",
        "",
        "[운영]",
        "머지 방식은 Squash and merge를 권장합니다(이슈 하나 = main 커밋 하나로 깔끔해짐).",
        "리뷰 코멘트는 줄 단위로 남기게 합니다(Files changed 탭에서 줄 번호 옆 + 버튼).",
    ])

    # ------------------------------------------------------------ 21. 충돌 해결
    s = S()
    header(s)
    title_block(s, "협업 ⑤ · 충돌 해결", "같은 줄을 다르게 고치면 사람이 결정해야 합니다")
    cw_ = 5.80
    codebox(s, ML, TOP, cw_, 2.30, "충돌 표시",
            ["<<<<<<< HEAD (내 브랜치)",
             "<h1>우주 탈출</h1>",
             "=======",
             "<h1>SPACE ESCAPE</h1>",
             ">>>>>>> main"])
    codebox(s, ML, TOP + 2.44, cw_, 1.48, "충돌 만들기 실습 · 2명",
            ["두 사람이 각자 브랜치에서 README 첫 줄을 다르게 수정 → 순서대로 PR"])
    rx = ML + cw_ + 0.26
    rw = CW - cw_ - 0.26
    codebox(s, rx, TOP, rw, 2.50, "PROMPT",
            ["main을 내 브랜치에 병합해 줘.",
             "충돌이 나면 해결하지 말고 먼저",
             "1. 충돌 파일과 위치",
             "2. 양쪽 변경의 의도",
             "3. 해결안 2가지를 보여 줘.",
             "내가 고른 안으로 해결하고 커밋해 줘."])
    lcard(s, rx, TOP + 2.62, rw, 1.30, RED, "금지", ["충돌 표시가 남은 채 커밋 · push"], label_color=RED, size=11)
    bar_note(s, [[("충돌은 오류가 아니라 ", {}), ("의사결정이 필요하다는 신호", {"bold": True, "color": ACCENT_D}),
                  ("   ·   해결 후 반드시 실행 확인", {})]])
    page(s, "협업 실습")
    note(n(), [
        "[시간] 5분",
        "",
        "[대본]",
        "\"충돌을 일부러 만들어 봅니다. 두 사람이 각자 브랜치에서 README 첫 줄을 다르게 고치고, 한 명 PR을 먼저 머지하세요.",
        "그러면 두 번째 PR에 'This branch has conflicts' 표시가 뜹니다.",
        "충돌 표시를 보면 HEAD는 내 쪽, 아래는 main 쪽입니다. Git은 둘 중 무엇이 맞는지 모릅니다. 사람이 결정해야 해요.",
        "Claude에게 바로 해결하라고 하지 말고, 먼저 양쪽 의도와 해결안을 보여 달라고 하세요. 고르는 건 여러분입니다.",
        "해결하고 나면 꼭 실행해서 확인하고, 커밋 · push합니다. 그러면 PR의 충돌 표시가 사라집니다.\"",
        "",
        "[운영]",
        "충돌 표시(<<<<<<< ======= >>>>>>>)가 남은 파일이 push되면 화면이 깨집니다. 리뷰 체크리스트로 잡게 합니다.",
    ])

    # ------------------------------------------------------------ 22. 협업 보안
    s = S()
    header(s)
    title_block(s, "협업 보안 체크", "공개 저장소는 공격자도 읽는다는 전제로 운영합니다")
    table(s, ML, TOP, CW, [3.10, 4.40, 4.37], ["위험", "원인", "대응"],
          [["비밀값 유출", "키 · 토큰을 코드에 넣고 커밋", "코드에 넣지 않음 · 노출 시 즉시 키 폐기"],
           ["기록에 남는 삭제", "삭제 커밋을 해도 과거 커밋에 남음", "삭제가 아니라 폐기 · 재발급"],
           ["불필요 파일 업로드", ".env · 로그 · 개인 메모", ".gitignore에 먼저 등록"],
           ["작성자 정보 노출", "개인 이메일로 커밋", "noreply 이메일 사용"],
           ["리뷰 우회", "main 직접 push · 자기 승인", "보호 규칙 · 예외 없음"],
           ["클라이언트 신뢰", "점수 · 기록을 브라우저에서 계산", "조작 가능함을 PRD 보안 요구사항에 명시"]],
          row_h=0.50, accent_col=None)
    bar_note(s, [[("GitHub의 비밀값 탐지 기능은 ", {}), ("보조 수단", {"bold": True, "color": ACCENT_D}),
                  ("  ·  1차 방어선은 커밋 전 확인과 리뷰", {})]], y=6.00)
    footer_rule(s); page(s, "협업 실습")
    note(n(), [
        "[시간] 3분",
        "",
        "[대본]",
        "\"여러분 전공이니까 조금 깊게 봅니다. 공개 저장소는 누구나 읽을 수 있고, 키를 찾는 자동 프로그램이 실제로 돌아다닙니다.",
        "가장 흔한 실수가 두 번째 줄입니다. 키를 올렸다가 지우는 커밋을 하면 괜찮을 거라고 생각하는데, 과거 커밋에 그대로 남아 있습니다. 답은 삭제가 아니라 키를 폐기하고 새로 발급하는 겁니다.",
        "마지막 줄도 중요합니다. 지금 게임들은 점수를 브라우저에서 계산하죠. 개발자 도구로 얼마든지 바꿀 수 있습니다. 랭킹을 붙일 거라면 이 한계를 PRD에 적어 두세요.",
        "GitHub도 공개 저장소에 키가 올라가는 걸 일부 탐지해서 막아 주지만, 그건 보조 수단입니다. 1차 방어선은 커밋 전에 바뀐 파일을 보는 습관과 리뷰예요.\"",
        "",
        "[운영]",
        "실습 중 실제 키가 커밋된 것을 발견하면: 저장소에서 지우는 것보다 키 폐기가 먼저입니다. 교사에게 바로 알리게 합니다.",
    ])

    # ------------------------------------------------------------ 23. 배포
    s = S()
    header(s)
    title_block(s, "배포 · GitHub Pages", "Owner · main에 머지되면 자동으로 갱신됩니다")
    step_row(s, TOP, [("STEP 1", "Settings → Pages", "저장소 설정 왼쪽 메뉴"),
                      ("STEP 2", "Source 선택", "Deploy from a branch"),
                      ("STEP 3", "main · / (root)", "Save"),
                      ("STEP 4", "1~3분 대기", "상단에 주소 표시")], h=1.60)
    w, gap = grid(2, 0.30)
    codebox(s, ML, 4.14, w, 1.30, "배포 주소 형식",
            ["https://[Owner아이디].github.io/[저장소이름]/"])
    lcard(s, ML + w + gap, 4.14, w, 1.30, ACCENT, "갱신", ["PR 머지 → 자동 재배포 (Netlify Drop 대체)"], size=11)
    bar_note(s, [[("404가 뜨면 ", {}), ("index.html이 저장소 최상위에 있는지", {"bold": True, "color": ACCENT_D}),
                  (" 먼저 확인", {})]], y=5.66, h=0.62)
    footer_rule(s); page(s, "배포 · 회고")
    note(n(), [
        "[시간] 4분",
        "",
        "[대본]",
        "\"5차시에는 Netlify Drop으로 폴더를 끌어다 놓았죠. 이제는 GitHub가 직접 배포합니다.",
        "Owner가 Settings → Pages에서 main 브랜치, 루트 폴더를 선택하고 저장하세요. 1~3분 뒤 위쪽에 주소가 뜹니다.",
        "이제부터는 PR이 머지될 때마다 자동으로 다시 배포됩니다. 끌어다 놓을 필요가 없어요.\"",
        "",
        "[운영]",
        "주소를 반 게시판에 올리고 QR로 서로 접속해 봅니다.",
        "404: index.html 위치 · 파일 이름 대소문자 · 배포 완료 전 접속 여부를 확인합니다.",
    ])

    # ------------------------------------------------------------ 24. 트러블슈팅
    s = S()
    header(s)
    title_block(s, "막혔을 때", "증상별 원인과 조치")
    table(s, ML, TOP, CW, [3.55, 3.70, 4.62], ["증상", "원인", "조치"],
          [["push 403 · Permission denied", "초대 미수락", "알림에서 초대 수락 후 다시 push"],
           ["push rejected (non-fast-forward)", "원격이 더 최신", "pull 후 다시 push"],
           ["protected branch 거부", "main에서 직접 작업", "정상 · 브랜치로 옮겨 PR"],
           ["로그인 창이 안 뜸", "로그인 연동 문제", "GitHub 웹에서 수정 · PR로 대체"],
           ["This branch has conflicts", "같은 줄 동시 수정", "충돌 해결 프롬프트 사용"],
           ["내 변경이 main에 없음", "머지 전 · pull 안 함", "PR 상태 확인 → main pull"],
           ["Pages 404", "경로 · 배포 대기", "index.html 위치 · 1~3분 대기"]],
          row_h=0.50)
    footer_rule(s); page(s, "배포 · 회고")
    note(n(), [
        "[시간] 참고용 — 실습 중 띄워 둡니다",
        "",
        "[운영]",
        "가장 흔한 순서: 초대 미수락(403) → main에서 작업(보호 규칙 거부) → pull 안 하고 push(rejected).",
        "protected branch 거부는 오류가 아니라 설정이 잘 된 증거라고 말해 줍니다.",
        "Claude에게 맡길 때: '지금 git 상태를 보여 주고, 이 오류의 원인을 설명해 줘. 아직 고치지 마.'",
    ])

    # ------------------------------------------------------------ 25. 협업 미션 체크
    s = S()
    header(s)
    title_block(s, "협업 미션 체크", "조별로 확인하고 저장소 링크를 제출합니다")
    checklist(s, ML, TOP, 6.55, [
        ("저장소 · 보호", "main 보호 규칙 · 직접 push 거부 확인"),
        ("이슈", "MUST 기능 수만큼 · 담당자 지정"),
        ("PR · 리뷰", "조원 1인당 PR 1개 머지 · 리뷰 코멘트"),
        ("충돌", "충돌 1회 이상 직접 해결"),
        ("배포", "GitHub Pages 주소 접속 확인"),
    ], row_h=0.70, gap=0.105)
    rx = ML + 6.55 + 0.26
    rw = CW - 6.55 - 0.26
    wcard(s, rx, TOP, rw, 1.90, ACCENT, "제출", "반 게시판",
          ["저장소 주소 · Pages 주소"], size=11, strong=True)
    wcard(s, rx, TOP + 2.04, rw, 1.88, GREEN, "평가 증거", None,
          ["커밋 · PR · 리뷰 기록", "= 개인별 기여"], label_color=GREEN_D, size=11)
    bar_note(s, [[("GitHub의 Insights → Contributors에서 ", {}), ("조원별 기여가 자동으로 기록", {"bold": True, "color": ACCENT_D}),
                  ("됩니다", {})]])
    page(s, "배포 · 회고")
    note(n(), [
        "[시간] 3분",
        "",
        "[대본]",
        "\"조별로 다섯 가지를 확인하세요. 다 되면 저장소 주소와 배포 주소를 게시판에 올립니다.",
        "앞으로 협업 평가는 이 기록으로 합니다. 누가 어떤 이슈를 맡았고, 몇 번 커밋했고, 어떤 리뷰를 남겼는지가 GitHub에 그대로 남아요.\"",
        "",
        "[운영]",
        "기여 평가: 저장소 Insights → Contributors, Pull requests 탭(Closed), 각 PR의 리뷰 기록을 확인합니다.",
        "Insights는 Public 저장소에서 누구나 볼 수 있습니다.",
    ])

    # ------------------------------------------------------------ 26. 회고 · 다음 차시
    s = S()
    header(s)
    title_block(s, "회고 · 다음 차시", "협업 과정에서 막힌 지점을 기록합니다")
    w, gap = grid(3, 0.26)
    for i, (num, h1, h2) in enumerate([("①", "가장 오래 걸린 단계", "세팅 · 이슈 · PR · 리뷰 · 충돌 중"),
                                       ("②", "리뷰에서 찾은 문제", "기능 · 범위 · 보안 중 한 가지"),
                                       ("③", "팀 규칙 추가", "CLAUDE.md에 넣을 규칙 한 줄")]):
        x = ML + i * (w + gap)
        rect(s, x, TOP, w, 1.30, fill=CARD, line=LINE, radius=0.05)
        text(s, x + 0.24, TOP + 0.20, 0.5, 0.3, [[(num, {"bold": True, "size": 12, "color": ACCENT})]])
        text(s, x + 0.24, TOP + 0.50, w - 0.48, 0.34, [[(h1, {"bold": True, "size": 14, "color": INK})]])
        text(s, x + 0.24, TOP + 0.88, w - 0.48, 0.30, [[(h2, {"size": 10.5, "color": MUTED})]])
    rect(s, ML, 3.92, CW, 1.30, fill=DARK, radius=0.06)
    text(s, ML + 0.40, 4.10, CW - 0.8, 0.28, [[("NEXT · 7차시", {"bold": True, "size": 10, "color": ACCENT_L, "space": 1.0})]])
    text(s, ML + 0.40, 4.44, CW - 0.8, 0.50, [[("이슈 기반 기능 구현 · 교차 플레이테스트", {"bold": True, "size": 22, "color": WHITE})]])
    bar_note(s, [[("다음 시간 시작 전: ", {}), ("담당 이슈 브랜치에서 작업을 이어서 PR로 제출", {"bold": True, "color": ACCENT_D})]],
             y=5.46, h=0.62)
    footer_rule(s); page(s, "배포 · 회고")
    note(n(), [
        "[시간] 3분",
        "",
        "[대본]",
        "\"오늘 정리할 건 세 가지입니다. 어느 단계에서 가장 오래 걸렸는지, 리뷰에서 무엇을 찾았는지, 그리고 팀 규칙에 한 줄을 더한다면 무엇인지.",
        "세 번째는 실제로 CLAUDE.md에 추가해서 PR로 올리세요. 그것도 협업 기록입니다.",
        "다음 시간에는 각자 맡은 이슈를 구현하고, 다른 조 게임을 서로 플레이하면서 버그를 찾습니다.\"",
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
    prs.save(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, DECK))   # 인자로 저장 위치 지정 가능
    return prs


if __name__ == "__main__":
    B.CUR["n"] = 0
    del B.WARN[:]
    prs = build()
    print("[6차시] %s" % DECK)
    print("  슬라이드 %d장 · 발표자 노트 %d개" % (len(prs.slides), len(NOTES)))
    if B.WARN:
        print("  [카드 넘침] %d건" % len(B.WARN))
        for w_ in B.WARN:
            print("    " + w_)
    probs = [x for x in V3.lint(prs) if ": 06" not in x]
    if probs:
        print("  [글상자 검사] %d건" % len(probs))
        for pr in probs:
            print("    " + pr)
    if not B.WARN and not probs:
        print("  넘침 없음")
