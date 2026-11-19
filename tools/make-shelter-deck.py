# -*- coding: utf-8 -*-
"""
쉼터 3D 제작 과정 PPT 생성기

실행:  python tools/make-shelter-deck.py [출력경로]
결과:  C:/Users/haha6/Desktop/hj_html/healing-space/쉼터_제작과정.pptx

- 디자인은 9·10차시 덱과 같은 헬퍼(make-lesson03-v2 / 04-v2 / 06)를 그대로 재사용
- 내용은 수업과 무관한 일반 제작기: 아이디어 → 비교 → 수정 → Supabase → Vercel → GitHub
"""
import importlib.util
import os
import sys

from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
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

OUT = r"C:/Users/haha6/Desktop/hj_html/healing-space/쉼터_제작과정.pptx"
RIGHT = "SHELTER 3D · 제작 과정"

ML, CW, SW, SH, TOP, FOOT_Y = B.ML, B.CW, B.SW, B.SH, B.TOP, B.FOOT_Y
INK, BODY, MUTED, FAINT = B.INK, B.BODY, B.MUTED, B.FAINT
ACCENT, ACCENT_D, ACCENT_L = B.ACCENT, B.ACCENT_D, B.ACCENT_L
LINE, CARD, TINT, DARK, WHITE = B.LINE, B.CARD, B.TINT, B.DARK, B.WHITE
BLUE, GREEN, PURPLE = B.BLUE, B.GREEN, B.PURPLE
rect, text, footer_rule, title_block = B.rect, B.text, B.footer_rule, B.title_block
wcard, codebox, step_row = B.wcard, B.codebox, B.step_row
grid, minicard, checklist = L4.grid, L4.minicard, L4.checklist
BLUE_D, GREEN_D, PURPLE_D, RED = L4.BLUE_D, L4.GREEN_D, L4.PURPLE_D, L4.RED
table, bar_note = L6.table, L6.bar_note

NOTES = {}
PAGES = []
TOTAL_PARTS = 6


def note(i, lines):
    NOTES[i] = "\n".join(lines)


def page(s, section):
    PAGES.append((s, section))


def header(s):
    B.header(s, right=RIGHT)


def part_slide(s, num, title, sub):
    header(s)
    text(s, ML, 2.55, CW, 0.34,
         [[("P A R T   %d" % num, {"bold": True, "size": 11, "color": ACCENT, "space": 1.6})]])
    text(s, ML, 2.98, CW, 0.85, [[(title, {"bold": True, "size": 38, "color": INK})]])
    rect(s, ML, 3.98, 0.80, 0.045, fill=ACCENT)
    text(s, ML, 4.22, CW * 0.8, 0.4, [[(sub, {"size": 13.5, "color": MUTED})]])
    footer_rule(s)
    text(s, SW - ML - 3, FOOT_Y + 0.12, 3, 0.26,
         [[("PART %d / %d" % (num, TOTAL_PARTS), {"size": 9, "color": FAINT, "space": 0.8})]],
         align=PP_ALIGN.RIGHT)


def std(prs, title, sub, section):
    s = B.blank(prs)
    header(s)
    title_block(s, title, sub)
    footer_rule(s)
    page(s, section)
    return s


def build():
    prs = B.new_deck()
    n = lambda: B.CUR["n"]

    # ============================================================ 표지
    s = B.blank(prs)
    text(s, ML, 0.44, 5, 0.3, [[("A I   ×   B U I L D", {"bold": True, "size": 11, "color": ACCENT, "space": 1.3})]])
    text(s, SW - ML - 7.4, 0.46, 7.4, 0.3,
         [[("바이브 코딩 제작기 · three.js × Supabase × Vercel", {"bold": True, "size": 10, "color": FAINT, "space": 1.0})]],
         align=PP_ALIGN.RIGHT)
    text(s, ML, 1.70, 6, 1.2, [[("SHELTER 3D", {"bold": True, "size": 54, "color": ACCENT_L})]])
    text(s, ML, 2.84, CW, 1.0, [[("쉼터 — 3D 힐링 공간 만들기", {"bold": True, "size": 40, "color": INK})]])
    rect(s, ML, 3.94, 0.9, 0.05, fill=ACCENT)
    text(s, ML, 4.20, CW * 0.9, 0.5,
         [[("아이디어 한 줄에서 여러 명이 함께 걷는 배포 사이트까지, 프롬프트로 만든 과정", {"size": 17, "color": BODY})]])
    w, gap = grid(3, 0.22)
    for i, (h1, h2) in enumerate([("만들기", "아이디어 · 비교 · 수정"),
                                  ("함께 쓰기", "Supabase 실시간 접속"),
                                  ("공개하기", "Vercel 배포 · GitHub 자동 배포")]):
        minicard(s, ML + i * (w + gap), 5.20, w, 1.20, h1, h2, bar=ACCENT if i == 1 else ACCENT_L)
    footer_rule(s)
    text(s, ML, FOOT_Y + 0.12, 8, 0.26,
         [[("shelter-gamma.vercel.app", {"size": 9.5, "color": FAINT, "space": 0.8})]])
    note(n(), [
        "[구성] PART 1 아이디어 · 2 비교 · 3 수정 · 4 Supabase · 5 Vercel · 6 GitHub",
        "",
        "[대본]",
        "\"브라우저에서 걸어 다니는 3D 힐링 공간 '쉼터'를 만든 과정을 소개합니다.",
        "코드는 AI가 썼고, 저는 무엇을 만들지 정하고, 비교하고, 고칠 점을 말하고, 배포 버튼을 눌렀습니다.\"",
    ])

    # ============================================================ 전체 흐름
    s = std(prs, "전체 흐름", "여섯 단계로 만들고 공개했습니다", "개요")
    step_row(s, TOP, [("1", "아이디어", "무엇을 · 어떤 제약으로"),
                      ("2", "비교", "비슷한 사례와 나란히"),
                      ("3", "수정", "보고 → 고침 → 실행")], h=1.40)
    step_row(s, TOP + 1.60, [("4", "Supabase", "여러 명 · 대화 · 저장"),
                             ("5", "Vercel", "주소로 공개"),
                             ("6", "GitHub", "푸시하면 자동 배포")], h=1.40)
    bar_note(s, [[("AI가 코드를 쓰는 동안 사람은 ", {}),
                  ("정하고, 비교하고, 확인했다", {"bold": True, "color": ACCENT_D})]], y=5.56, h=0.62)
    note(n(), [
        "[대본]",
        "\"1~3은 '만들기', 4는 '함께 쓰기', 5~6은 '공개하기'입니다.",
        "각 단계마다 실제로 입력한 프롬프트를 함께 보여 드립니다.\"",
    ])

    # ============================================================ PART 1 아이디어
    s = B.blank(prs)
    part_slide(s, 1, "아이디어 선정", "무엇을 만들지, 어떤 제약 안에서 만들지 정합니다")
    note(n(), ["[대본] \"처음 정한 것은 '어떤 공간'이 아니라 '어떤 느낌'이었습니다. 잠깐 쉬어 가는 곳.\""])

    s = std(prs, "무엇을 만들까", "느낌 · 제약 · 함께를 먼저 정했습니다", "아이디어")
    w, gap = grid(3, 0.26)
    wcard(s, ML, TOP, w, 2.70, ACCENT, "느낌", "힐링 3D 공간",
          ["걸어 다니며 쉬는 곳", "연못 · 언덕 · 카페 · 별", "해 질 녘 따뜻한 색"], label_color=ACCENT_D, size=11, strong=True)
    wcard(s, ML + w + gap, TOP, w, 2.70, BLUE, "제약", "외부 에셋 0개",
          ["three.js 하나만 CDN", "텍스처 · 소리 모두 코드 생성", "더블클릭만으로 실행"], label_color=BLUE_D, size=11)
    wcard(s, ML + 2 * (w + gap), TOP, w, 2.70, GREEN, "함께", "여러 명 접속",
          ["같은 방에서 서로 보이기", "말풍선 대화", "쪽지 · 소원 저장"], label_color=GREEN_D, size=11)
    bar_note(s, [[("제약을 먼저 말하면 ", {}), ("AI가 고르는 방법이 좁혀진다", {"bold": True, "color": ACCENT_D}),
                  ("  — 파일 몇 개, 무엇을 쓰지 않을지", {})]], y=5.40, h=0.62)
    note(n(), [
        "[대본]",
        "\"'외부 에셋 0개'는 중요한 제약이었습니다. 그림 파일이나 음원을 받지 않으니 저작권 걱정이 없고,",
        "파일을 더블클릭만 해도 열립니다. 대신 나무, 풀, 소리까지 전부 코드로 만들어야 했습니다.\"",
    ])

    s = std(prs, "공간 구성", "장소마다 '할 수 있는 일'을 하나 이상 붙였습니다", "아이디어")
    table(s, ML, TOP, CW, [2.5, 1.6, CW - 4.1],
          ["장소", "좌표", "할 수 있는 일"],
          [["쉼터 (광장 · 연못)", "0, 0", "연못가에 앉기 · 물소리 · 3단 분수 · 연잎 · 잉어"],
           ["소나무 언덕", "-30, -16", "벤치에서 숨 고르기 · 쪽지 걸기 / 읽기"],
           ["작은 온실", "30, -16", "물뿌리개로 화분에 물 주기"],
           ["창가 카페", "-28, 22", "창가에 앉아 공원 보기 · 음료 · 빗소리"],
           ["음악실", "28, 22", "피아노 앞에 앉아 한 소절 치기 · 힐링 음악"],
           ["별빛 운동장", "0, -46", "누워서 별 보기 · 별자리 · 별똥별 · 소원 별"],
           ["멍냥 쉼터", "0, 34", "간식 주기 · 쓰다듬기 · 공 던지기"],
           ["북쪽 냇가", "z ≈ 72", "나무다리 · 정자 · 돌탑"]],
          row_h=0.46, size=10.5)
    note(n(), [
        "[대본]",
        "\"장소만 있으면 구경하고 끝납니다. 장소마다 앉기, 물 주기, 별 보기처럼 '할 일'을 붙였습니다.",
        "NPC 다섯 명과 강아지, 고양이도 함께 있습니다.\"",
    ])

    s = std(prs, "아이디어 → 프롬프트", "느낌 · 장소 · 할 일 · 제약을 한 번에 적습니다", "아이디어")
    codebox(s, ML, TOP, 7.40, 3.90, "PROMPT · 처음 요청 (예시)",
            ["브라우저에서 걸어 다니는 3D 힐링 공간을 만들어 줘.",
             "",
             "- 장소: 연못 광장, 소나무 언덕, 온실, 카페,",
             "        음악실, 별 보는 운동장",
             "- 장소마다 할 일: 앉기, 물 주기, 피아노, 별 보기",
             "- 제약: three.js 하나만, 그림·음원 파일 없이",
             "        전부 코드로, 더블클릭으로 실행",
             "- 나중에 여러 명이 같이 접속할 수 있게"])
    rx = ML + 7.40 + 0.26
    wcard(s, rx, TOP, CW - 7.66, 3.90, ACCENT, "빠지면 생기는 일", "네 가지 체크",
          ["느낌이 없으면 → 밋밋한 데모", "할 일이 없으면 → 구경만", "제약이 없으면 → 파일이 흩어짐",
           "함께가 없으면 → 나중에 큰 수정"], label_color=ACCENT_D, size=11)
    note(n(), [
        "[대본]",
        "\"한 문장으로 '힐링 공간 만들어 줘'라고 해도 무언가는 나옵니다. 하지만 장소와 할 일,",
        "그리고 '무엇을 쓰지 말라'는 제약까지 적으면 첫 결과물이 훨씬 원하는 쪽에 가깝습니다.\"",
    ])

    # ============================================================ PART 2 비교
    s = B.blank(prs)
    part_slide(s, 2, "유사 사례 검색 · 비교", "비슷한 작품을 기준으로 놓고 무엇이 부족한지 찾습니다")
    note(n(), ["[대본] \"첫 결과물은 동작은 했지만 어딘가 아쉬웠습니다. 기준이 필요했습니다.\""])

    s = std(prs, "기준 사례 찾기", "산사 메타버스(sansatemple.com)를 기준으로 정했습니다", "비교")
    step_row(s, TOP, [("1", "검색", "3D 힐링 · 메타버스 · 산책"),
                      ("2", "후보 고르기", "직접 걸어 보고 2~3곳"),
                      ("3", "기준 하나", "\"이만큼 해 줘\""),
                      ("4", "비교 요청", "항목별로 나란히")], h=1.50)
    codebox(s, ML, 4.04, CW, 2.02, "PROMPT · 실제로 입력한 문장",
            ["산사 홈페이지와 지금 만든 메타버스를 비교해봐",
             "디자인이 너무 형편없어",
             "3d로 현실의 느낌이 들 만큼 구성해",
             "사람도 귀여운캐릭터로 실제 사람 느낌이 나게"])
    bar_note(s, [[("\"더 예쁘게\" 대신 ", {}), ("\"이 사이트만큼\"", {"bold": True, "color": ACCENT_D}),
                  (" — 기준이 있으면 AI가 차이를 스스로 찾는다", {})]], y=6.18, h=0.50)
    note(n(), [
        "[대본]",
        "\"'더 예쁘게'는 기준이 없습니다. 기준 사이트를 하나 정해 주면 AI가 그 사이트와 지금 결과물을",
        "항목별로 비교해서 부족한 점을 먼저 정리하고, 그다음에 고칩니다.\"",
    ])

    s = std(prs, "비교표", "항목별로 나란히 놓으니 고칠 순서가 보였습니다", "비교")
    table(s, ML, TOP, CW, [2.2, 3.2, 3.2, CW - 8.6],
          ["항목", "기준 (산사 메타버스)", "처음 쉼터", "고친 뒤"],
          [["디자인 · 색", "따뜻한 자연광 · 질감", "단색 면 · 평평한 색", "해 질 녘 팔레트 · 코드 질감"],
           ["3D 깊이감", "먼 산 · 안개 · 지형", "평평한 바닥 · 보라 뾰족 산", "낮은 능선 · 대기 원근"],
           ["캐릭터", "귀엽지만 사람 같은 비율", "단순 막대 인형", "3.5등신 · 관절 · 자세 4종"],
           ["NPC", "주변에 사람들이 있음", "나 혼자", "NPC 5명 · 강아지 · 고양이"],
           ["조명", "지나가면 등이 켜짐", "고정 조명", "석등 · 실내 조명 · 블룸"],
           ["소리", "자연음 · 분위기", "없음 / 웅 소리", "바람 · 물 · 새 · 풀벌레"]],
          row_h=0.50, size=10.5, accent_col=3)
    note(n(), [
        "[대본]",
        "\"표로 놓고 보니 '형편없다'는 느낌이 구체적인 항목 여섯 개로 바뀌었습니다.",
        "오른쪽 열이 이후 수정에서 실제로 바뀐 결과입니다.\"",
    ])

    s = std(prs, "비교 후 바뀐 것", "장면 · 사람 · 분위기 세 갈래로 다시 만들었습니다", "비교")
    w, gap = grid(3, 0.26)
    wcard(s, ML, TOP, w, 2.90, ACCENT, "장면", "현실감 있는 지형",
          ["지형 잡음 + 냇가 골", "먼 산은 대기 원근 직접 계산", "풀 14,000 · 낙엽 인스턴싱"], label_color=ACCENT_D, size=11)
    wcard(s, ML + w + gap, TOP, w, 2.90, BLUE, "사람", "귀여운데 사람 같은",
          ["3.5등신 · 팔꿈치 · 무릎", "서기 · 앉기 · 눕기 · 쪼그리기", "얼굴은 단순하게"], label_color=BLUE_D, size=11)
    wcard(s, ML + 2 * (w + gap), TOP, w, 2.90, PURPLE, "분위기", "빛과 소리",
          ["다가가면 켜지는 석등", "밤 전환 · 별자리 · 오로라", "음악 대신 자연음"], label_color=PURPLE_D, size=11)
    bar_note(s, [[("한 번에 다 고치지 않고 ", {}), ("갈래별로 실행해 보며", {"bold": True, "color": ACCENT_D}),
                  (" 다음 피드백으로 넘어갔다", {})]], y=5.50, h=0.62)
    note(n(), ["[대본] \"얼굴은 단순할수록 낫다는 것도 이 과정에서 배웠습니다. 눈썹·눈꺼풀을 얹으면 가면처럼 보였습니다.\""])

    # ============================================================ PART 3 수정
    s = B.blank(prs)
    part_slide(s, 3, "추가 수정하기", "보이는 대로 말하고, 고치고, 다시 실행합니다")
    note(n(), ["[대본] \"가장 시간이 많이 든 단계입니다. 프롬프트를 열 번 넘게 주고받았습니다.\""])

    s = std(prs, "수정 루프", "말 → 원인 → 고침 → 실행 확인을 반복했습니다", "수정")
    step_row(s, TOP, [("1", "보고", "본 것을 그대로 말하기"),
                      ("2", "원인", "AI가 코드에서 찾기"),
                      ("3", "수정", "원인을 고치기"),
                      ("4", "실행 확인", "직접 띄워서 보기")], h=1.50)
    w, gap = grid(2, 0.26)
    wcard(s, ML, 4.06, w, 1.90, GREEN, "잘 통한 말", None,
          ["\"다리가 빠지는 느낌이야, 다리가 가려져\"", "\"벽 앞에서 헛걸음으로 움직이고 있어\""],
          label_color=GREEN_D, size=11)
    wcard(s, ML + w + gap, 4.06, w, 1.90, RED, "잘 안 통한 말", None,
          ["\"이상해\" · \"별로야\"", "→ 어디가, 어떻게 이상한지 한 줄 더"],
          label_color=RED, size=11)
    bar_note(s, [[("말로 어려우면 ", {}), ("스크린샷 · 영상을 폴더에 넣고", {"bold": True, "color": ACCENT_D}),
                  (" \"이 영상 봐\"라고 한다", {})]], y=6.14, h=0.50)
    note(n(), [
        "[대본]",
        "\"좋은 피드백은 '본 것'을 그대로 말하는 것입니다. 원인은 AI가 찾습니다.",
        "화면이 떨리는 문제는 말로 설명이 안 돼서 화면 녹화 영상을 폴더에 넣고 보여 줬습니다.\"",
    ])

    s = std(prs, "수정 프롬프트 ①", "한 번에 여러 개를 말해도 됩니다", "수정")
    codebox(s, ML, TOP, 7.60, 4.10, "PROMPT · 실제 입력",
            ["쉼터 연못 언덕에 다리가 빠지는 느낌이야 다리가 가려져",
             "캐릭터의 얼굴이 머리카락과 눈이 너무 조화가 안 맞아",
             "산사처럼 주변에 npc 처럼 사람들이 몇명 있게 만들어줘",
             "작은온실의 풀들이 그냥 동그라미 같아",
             "물주기도 하늘에서 물방울 몇개가 튀어 내려오는 느낌이야",
             "배경 음악소리도 안켜져",
             "주변의 산이 보라색으로 뾰족하기만해",
             "피아노 건반은 도레미파솔라시도 음계가 엉터리야"])
    rx = ML + 7.60 + 0.26
    wcard(s, rx, TOP, CW - 7.86, 4.10, ACCENT, "고친 결과", None,
          ["발 묻힘 → 서는 높이 +0.05", "NPC 5명 · 머리 · 키 다르게", "잎 모양 풀 · 손에 든 물뿌리개",
           "실제 산 능선 · 다장조 건반"], label_color=ACCENT_D, size=11)
    note(n(), [
        "[대본]",
        "\"이 프롬프트 하나에 여덟 가지가 들어 있습니다. 순서나 문법은 상관없습니다.",
        "본 것을 줄마다 하나씩 적으면 AI가 하나씩 원인을 찾아 고칩니다.\"",
    ])

    s = std(prs, "수정 프롬프트 ②", "같은 문제가 남아 있으면 다시 말합니다", "수정")
    codebox(s, ML, TOP, 7.60, 4.10, "PROMPT · 실제 입력",
            ["건물에 닿으면 튕겨지면서 화면이 흔들려",
             "",
             "폴더 안의 동영상을 첨부했어",
             "카페나 벽에 닿으면 이렇게 화면이 떨려",
             "",
             "벽앞에서 계속 헛걸음으로 움직이고 있으면 안돼",
             "다른곳으로 가야지",
             "",
             "여전히 강아지, 테이블, 의자들은 그냥 뚫고 지나가고 있어"])
    rx = ML + 7.60 + 0.26
    wcard(s, rx, TOP, CW - 7.86, 4.10, BLUE, "고친 결과", None,
          ["벽 · 건물은 회전 상자 충돌", "정면 충돌은 그냥 멈춤", "클릭 이동은 돌아서 가기",
           "영상에서 프레임을 떠 원인 확인"], label_color=BLUE_D, size=11)
    note(n(), [
        "[대본]",
        "\"화면 떨림은 세 번 넘게 다시 말했습니다. 결국 화면 녹화 영상을 넣어 주자 AI가 프레임을 하나씩 떠서",
        "'벽에 정면으로 부딪혀도 미끄러지려다 진동한다'는 원인을 찾았습니다.\"",
    ])

    s = std(prs, "수정 프롬프트 ③", "새 기능도 같은 방식으로 요청합니다", "수정")
    codebox(s, ML, TOP, 7.60, 4.10, "PROMPT · 실제 입력",
            ["하늘에 누워서 별을 봐야 하는데 보이지 않아",
             "별똥별이나 북두칠성 같은 별자리, 오로라도 보이면 좋겠어",
             "분수대를 조금 더 웅장하게",
             "",
             "한쪽에 고양이, 강아지 카페가 있어도 좋겠어",
             "먹이주고 쓰담쓰담 하면 동물들이 행복해 하면 좋겠어",
             "",
             "의자에 바르게 앉아서 편안하게 커피마시는 모습을 보여줘"])
    rx = ML + 7.60 + 0.26
    wcard(s, rx, TOP, CW - 7.86, 4.10, PURPLE, "생긴 것", None,
          ["밤 전환 · 별자리 5개 · 오르골", "3단 분수 · 물줄기", "멍냥 쉼터 · 간식 · 공 던지기",
           "창가에서 공원을 보며 앉기"], label_color=PURPLE_D, size=11)
    note(n(), ["[대본] \"고치기와 만들기의 경계는 없습니다. '있으면 좋겠어'도 같은 말투로 요청했습니다.\""])

    s = std(prs, "이미 밟은 지뢰", "같은 실수를 반복하지 않도록 CLAUDE.md에 적어 두었습니다", "수정")
    table(s, ML, TOP, CW, [3.3, 4.4, CW - 7.7],
          ["증상", "원인", "해결"],
          [["장면 전체가 돌아감", "add()는 부모를 돌려줌", "변수에 받아 회전 후 add"],
           ["연못 물이 안 보임", "속 찬 원통이 뚜껑이 됨", "윗면 없는 원통"],
           ["꽃이 프로펠러 같음", "꽃잎이 접선 방향", "바깥 방향 + 45° 세우기"],
           ["나무가 초록 덩어리", "잎이 줄기를 가림", "줄기 · 가지 · 위쪽 잎"],
           ["산이 베이지 종이", "안개에 묻힘", "산만 안개 끄고 직접 계산"],
           ["벽 앞 화면 떨림", "정면 충돌에도 미끄러짐", "정면이면 멈춤 · 클릭은 우회"],
           ["Vercel에 옛 버전", "index.html이 먼저 응답", ".vercelignore로 제외"]],
          row_h=0.44, size=10.5)
    note(n(), [
        "[대본]",
        "\"CLAUDE.md는 AI가 작업 전에 읽는 인수인계 문서입니다. 여기에 '이미 밟은 지뢰'를 적어 두면",
        "다른 계정이나 새 대화에서도 같은 실수를 하지 않습니다.\"",
    ])

    s = std(prs, "고쳤다는 증거", "\"고쳤습니다\"로 끝내지 않고 실제로 띄워서 확인했습니다", "수정")
    table(s, ML, TOP, 7.60, [1.6, 6.0],
          ["무엇", "방법"],
          [["오류", "브라우저 자동 실행 → 페이지 오류 0건"],
           ["문법", "node로 app.v2.js 읽어 보기"],
           ["흔들림", "영상 녹화 → 프레임 차이 평균(YAVG)"],
           ["소리", "오실레이터 생성 횟수 세기"]],
          row_h=0.50, size=10.5)
    rx = ML + 7.60 + 0.26
    rw = CW - 7.86
    rect(s, rx, TOP, rw, 2.40, fill=DARK, radius=0.06)
    text(s, rx + 0.3, TOP + 0.28, rw - 0.6, 0.3, [[("흔들림 수치", {"bold": True, "size": 10, "color": ACCENT_L, "space": 1.0})]])
    text(s, rx + 0.3, TOP + 0.70, rw - 0.6, 0.7, [[("12~14  →  0.1", {"bold": True, "size": 28, "color": WHITE})]])
    text(s, rx + 0.3, TOP + 1.55, rw - 0.6, 0.6, [[("고치기 전  →  고친 뒤", {"size": 11, "color": B.DARK_MUT})]])
    bar_note(s, [[("요청 끝에 ", {}), ("\"고친 뒤 실행해서 확인해 줘\"", {"bold": True, "color": ACCENT_D}),
                  ("를 붙이면 증거까지 받아 온다", {})]], y=5.40, h=0.62)
    note(n(), [
        "[대본]",
        "\"AI가 '고쳤습니다'라고 해도 실제로 안 고쳐진 경우가 있었습니다. 그래서 실제 브라우저로 띄워 오류를 세고,",
        "흔들림은 영상 프레임 차이를 숫자로 비교했습니다. 12~14였던 값이 0.06~0.15로 내려갔습니다.\"",
    ])

    # ============================================================ PART 4 Supabase
    s = B.blank(prs)
    part_slide(s, 4, "Supabase 사용", "여러 명이 같은 공간에 들어오고, 남긴 것을 저장합니다")
    note(n(), ["[대본] \"혼자 쓰는 공간에서 함께 쓰는 공간으로 바꾸는 단계입니다.\""])

    s = std(prs, "왜 필요할까", "내 컴퓨터 밖에 '함께 보는 곳'이 있어야 합니다", "Supabase")
    w, gap = grid(3, 0.26)
    wcard(s, ML, TOP, w, 2.70, ACCENT, "실시간", "서로 보이기",
          ["이동 · 옷 · 머리 모양 전송", "같은 방 코드끼리만", "저장하지 않음"], label_color=ACCENT_D, size=11)
    wcard(s, ML + w + gap, TOP, w, 2.70, BLUE, "실시간", "말풍선 대화",
          ["인사 · 짧은 대화", "6초 뒤 사라짐", "저장하지 않음"], label_color=BLUE_D, size=11)
    wcard(s, ML + 2 * (w + gap), TOP, w, 2.70, GREEN, "저장", "쪽지 · 소원 · 카운터",
          ["언덕에 건 쪽지", "운동장의 소원 별", "물 주기 · 응원 횟수"], label_color=GREEN_D, size=11)
    codebox(s, ML, 5.14, CW, 1.42, "PROMPT · 실제 입력",
            ["쉼터도 코드를 입력해서 들어오고 여러사람이 접속하게 하려면 버셀과 수파베이스에서 해야 하는거야?",
             "여러사람들이 들어와서 대화창에 인사하는거까지 할 수 있어?"])
    note(n(), [
        "[대본]",
        "\"움직임과 대화는 그 순간만 전달하고, 쪽지와 소원처럼 남아야 하는 것만 데이터베이스에 저장합니다.",
        "Supabase를 쓰지 않으면 각자 자기 화면에서 혼자 걷게 됩니다.\"",
    ])

    s = std(prs, "설정 순서", "화면에서 다섯 번만 하면 됩니다", "Supabase")
    step_row(s, TOP, [("1", "가입", "GitHub 계정으로"),
                      ("2", "새 프로젝트", "이름 · 비밀번호 · 지역"),
                      ("3", "SQL 실행", "schema.v2.sql 붙여넣기"),
                      ("4", "키 복사", "Settings › API"),
                      ("5", "config.js", "주소 · 공개 키 넣기")], h=1.60)
    codebox(s, ML, 4.12, CW, 1.76, "config.js",
            ["url:     'https://xxxxxxxx.supabase.co',",
             "anonKey: 'sb_publishable_••••••••••••',   // 공개 키만",
             "room:    'MAIN'"])
    bar_note(s, [[("SQL은 AI가 만든 파일을 ", {}), ("SQL Editor에 붙여 넣고 Run", {"bold": True, "color": ACCENT_D}),
                  (" — Success가 나오면 끝", {})]], y=6.06, h=0.50)
    note(n(), [
        "[대본]",
        "\"SQL은 데이터베이스에 표를 만드는 명령입니다. 직접 쓸 필요 없이 AI가 만든 schema.v2.sql을 복사해서",
        "SQL Editor에 붙여 넣고 Run을 누르면 됩니다. 그다음 주소와 공개 키를 config.js에 넣습니다.\"",
    ])

    s = std(prs, "키 두 종류", "공개 키만 쓰고, 비밀 키는 누구에게도 보여 주지 않습니다", "Supabase")
    w, gap = grid(2, 0.26)
    wcard(s, ML, TOP, w, 2.90, GREEN, "PUBLISHABLE · 공개 키", "sb_publishable_…",
          ["config.js에 넣어도 됨", "사이트를 여는 누구에게나 보임", "보안은 RLS 규칙이 담당"],
          label_color=GREEN_D, size=11, strong=True)
    wcard(s, ML + w + gap, TOP, w, 2.90, RED, "SECRET · 비밀 키", "sb_secret_… · service_role",
          ["파일 · 채팅 · 캡처에 절대 금지", "AI에게도 보여 주지 않기", "노출되면 바로 재발급"],
          label_color=RED, size=11)
    codebox(s, ML, 5.44, CW, 1.10, "PROMPT · 실제 질문",
            ["시크릿말고 퍼블리시만 복사하라는거지?? 시크릿은 절대 비공개로 하고"])
    note(n(), [
        "[대본]",
        "\"키가 두 종류입니다. 공개 키는 원래 브라우저에 보이는 키라 config.js에 넣어도 됩니다.",
        "비밀 키는 모든 데이터를 지울 수 있는 열쇠라서, 캡처할 때도 가려야 합니다.\"",
    ])

    s = std(prs, "모르면 그대로 묻기", "화면에서 막힌 곳을 그대로 물어봤습니다", "Supabase")
    codebox(s, ML, TOP, 7.30, 4.10, "PROMPT · 실제 질문들",
            ["새프로젝트만들기 누르면 돼?",
             "한글로 되어 있으니 이 기준으로 설명해봐",
             "새쿼리가 어디 있지?",
             "sql은 뭐야?",
             "너가 준 시퀄 붙여 넣었어 그다음엔? 실행 눌러?",
             "anon public 이거 못찾겠어",
             "키를 못찾겠는데 어떻게 점 표시로 가린채로 주니"])
    rx = ML + 7.30 + 0.26
    wcard(s, rx, TOP, CW - 7.56, 4.10, ACCENT, "덕분에 생긴 것", None,
          ["한글 화면 · 영문 화면 안내서", "캡처 위치를 짚은 단계별 설명", "SQL이 무엇인지 쉬운 설명"],
          label_color=ACCENT_D, size=11)
    note(n(), [
        "[대본]",
        "\"'어디 있지?'처럼 짧게 물어도 됩니다. 이 질문들이 모여 한글 화면 · 영문 화면 안내서 두 개가 만들어졌습니다.\"",
    ])

    # ============================================================ PART 5 Vercel
    s = B.blank(prs)
    part_slide(s, 5, "Vercel 배포", "내 폴더를 누구나 들어올 수 있는 주소로 만듭니다")
    note(n(), ["[대본] \"이제 링크 하나로 다른 사람을 초대할 수 있게 만듭니다.\""])

    s = std(prs, "배포 순서", "명령 두 줄로 주소가 생깁니다", "Vercel")
    step_row(s, TOP, [("1", "로그인", "vercel.com · GitHub 계정"),
                      ("2", "미리 배포", "npx vercel"),
                      ("3", "정식 배포", "npx vercel --prod"),
                      ("4", "공유", "주소 · ?room=코드")], h=1.50)
    codebox(s, ML, 4.06, 6.90, 2.20, "터미널",
            ["cd healing-space",
             "npx vercel          # 처음: 질문에 Enter",
             "npx vercel --prod   # 정식 주소에 반영"])
    rx = ML + 6.90 + 0.26
    wcard(s, rx, 4.06, CW - 7.16, 2.20, ACCENT, "PROMPT · 실제 입력", None,
          ["버셀에 배포하는 방법 알려줘", "배포된 주소가 어딨어?"], label_color=ACCENT_D, size=11)
    note(n(), [
        "[대본]",
        "\"npx vercel은 미리보기 주소, --prod를 붙이면 정식 주소에 올라갑니다.",
        "GitHub를 연결하기 전에는 고칠 때마다 이 명령을 다시 실행해야 했습니다.\"",
    ])

    s = std(prs, "꼭 확인할 설정", "이걸 빼먹으면 '안 열려요'가 됩니다", "Vercel")
    table(s, ML, TOP, CW, [3.4, 4.6, CW - 8.0],
          ["설정", "왜", "어떻게"],
          [["Deployment Protection", "켜져 있으면 외부인이 로그인 화면에 막힘", "Settings › 끄기"],
           [".vercelignore", "옛 버전 index.html이 먼저 열림", "v1 파일 · 영상 제외"],
           ["vercel.json", "/ 와 /room/코드 를 새 화면으로", "rewrites"],
           ["방 코드 주소", "코드를 아는 사람만 같은 방", "주소?room=CODE"]],
          row_h=0.56, size=10.5)
    codebox(s, ML, 5.10, CW, 1.36, "PROMPT · 실제 질문",
            ["왜 이게 올라갔어??   (→ 옛 버전이 떠서 .vercelignore 추가)",
             "방코드는 어떻게 설정해? 코드를 아는 사람이 들어오게 하려면?"])
    note(n(), [
        "[대본]",
        "\"처음 배포했을 때 새 버전이 아니라 옛 버전이 떴습니다. Vercel이 index.html을 먼저 보여 주기 때문이라",
        ".vercelignore로 옛 파일을 빼서 해결했습니다. 방 코드는 따로 등록할 필요 없이 주소 뒤에 붙이면 됩니다.\"",
    ])

    s = std(prs, "링크 미리보기", "카카오톡에 공유하면 썸네일과 제목이 보이게", "Vercel")
    codebox(s, ML, TOP, 7.60, 2.60, "index.v2.html · <head>",
            ['<meta property="og:title" content="쉼터 — 잠깐 걸으며 쉬어 가는 곳">',
             '<meta property="og:description" content="...">',
             '<meta property="og:url" content="https://shelter-gamma.vercel.app/">',
             '<meta property="og:image"',
             '      content="https://shelter-gamma.vercel.app/og.png">'])
    rx = ML + 7.60 + 0.26
    wcard(s, rx, TOP, CW - 7.86, 2.60, RED, "안 보였던 이유", None,
          ["og:image가 og.png (상대 주소)", "→ https://…/og.png 전체 주소로"], label_color=RED, size=11)
    step_row(s, 5.14, [("1", "og.png", "1200×630 장면 캡처"),
                       ("2", "전체 주소", "https://로 시작"),
                       ("3", "카카오 초기화", "developers.kakao.com › og")], h=1.30)
    note(n(), [
        "[대본]",
        "\"카카오톡은 썸네일 주소가 https://로 시작해야 읽습니다. 그리고 한 번 읽은 미리보기를 기억하기 때문에",
        "고친 뒤에는 카카오 개발자 사이트의 '공유 디버거'에서 초기화를 눌러야 새 썸네일이 보입니다.\"",
    ])

    # ============================================================ PART 6 GitHub
    s = B.blank(prs)
    part_slide(s, 6, "GitHub 연결", "푸시만 하면 사이트가 저절로 바뀌게 합니다")
    note(n(), ["[대본] \"마지막으로, 고칠 때마다 배포 명령을 치지 않아도 되게 만듭니다.\""])

    s = std(prs, "연결 순서", "처음 한 번만 합니다", "GitHub")
    step_row(s, TOP, [("1", "저장소 만들기", "github.com/new · 빈 저장소"),
                      ("2", "첫 업로드", "AI에게 \"깃허브에 올려 줘\""),
                      ("3", "Vercel 연결", "Settings › Git › Connect"),
                      ("4", "확인", "Git repository connected")], h=1.60)
    codebox(s, ML, 4.20, 6.90, 2.20, ".gitignore · 올리지 않을 것",
            ["*.mp4          # 녹화 영상",
             "_비교샷/  캡쳐/",
             ".vercel/  node_modules/  .env*"])
    rx = ML + 6.90 + 0.26
    wcard(s, rx, 4.20, CW - 7.16, 2.20, ACCENT, "순서 주의", None,
          ["빈 저장소는 연결해도 배포 실패", "→ 먼저 올리고, 그다음 연결"], label_color=ACCENT_D, size=11)
    note(n(), [
        "[대본]",
        "\"저장소가 비어 있으면 Vercel이 배포할 게 없습니다. 그래서 먼저 파일을 올리고 그다음 연결합니다.",
        "올리기 전에 AI가 비밀 키가 섞여 있지 않은지 검사했습니다.\"",
    ])

    s = std(prs, "커밋과 푸시", "커밋만으로는 사이트가 바뀌지 않습니다", "GitHub")
    w, gap = grid(2, 0.26)
    wcard(s, ML, TOP, w, 2.60, BLUE, "COMMIT", "내 컴퓨터에 기록",
          ["바뀐 내용을 저장해 두기", "아직 GitHub에는 없음", "사이트도 그대로"], label_color=BLUE_D, size=11)
    wcard(s, ML + w + gap, TOP, w, 2.60, GREEN, "PUSH", "GitHub로 보내기",
          ["GitHub에 도착", "Vercel이 알아채고 자동 배포", "약 1분 뒤 사이트 반영"], label_color=GREEN_D, size=11, strong=True)
    codebox(s, ML, 5.10, CW, 1.42, "PROMPT · 실제 질문",
            ["비밀키 들어 있지 않은채로 깃허브에 올리면",
             "앞으로 커밋만 하면 내가 버셀에 새롭게 업데이트 하지 않아도 다 된다는거지?"])
    note(n(), [
        "[대본]",
        "\"커밋은 내 컴퓨터 안에 기록, 푸시는 GitHub로 보내기입니다. Vercel은 GitHub에 도착한 것을 보고 배포하므로",
        "푸시까지 해야 사이트가 바뀝니다. AI에게는 '푸시해 줘' 한 마디면 둘 다 해 줍니다.\"",
    ])

    s = std(prs, "이제부터의 작업", "고치고 → \"푸시해 줘\" → 1분 뒤 확인", "GitHub")
    codebox(s, ML, TOP, 6.20, 3.50, "PROMPT · 매번",
            ["(고칠 내용) ... 고쳐 줘.",
             "고친 뒤 실행해서 확인하고,",
             "푸시해 줘."])
    rx = ML + 6.20 + 0.26
    checklist(s, rx, TOP, CW - 6.46, [
        ("비밀 키", "service_role · sb_secret_ 없음"),
        ("푸시", "GitHub에 새 커밋 도착"),
        ("Ready", "Vercel › Deployments 초록색"),
        ("사이트", "새로고침해서 바뀐 것 확인")], row_h=0.78, gap=0.12)
    bar_note(s, [[("배포 실패는 Vercel이 ", {}), ("이메일로 알려 준다", {"bold": True, "color": ACCENT_D}),
                  (" — 일일이 확인하지 않아도 된다", {})]], y=6.08, h=0.55)
    note(n(), [
        "[대본]",
        "\"이제 작업 흐름은 이 세 줄이 전부입니다. 연결을 확인하려고 빈 커밋을 하나 푸시했더니",
        "30초 뒤 Vercel에 새 배포가 Ready로 올라왔습니다.\"",
    ])

    # ============================================================ 정리
    s = std(prs, "전체 정리", "프롬프트 한 줄에서 함께 쓰는 공개 사이트까지", "정리")
    step_row(s, TOP, [("1", "아이디어", "느낌 · 장소 · 제약"),
                      ("2", "비교", "기준 사이트 하나"),
                      ("3", "수정", "본 것 그대로 · 영상")], h=1.40)
    step_row(s, TOP + 1.60, [("4", "Supabase", "공개 키만"),
                             ("5", "Vercel", "보호 끄기 · og"),
                             ("6", "GitHub", "푸시 = 배포")], h=1.40)
    bar_note(s, [[("코드는 AI가 썼지만 ", {}),
                  ("무엇을 · 무엇과 비교해 · 언제 공개할지", {"bold": True, "color": ACCENT_D}),
                  ("는 사람이 정했다", {})]], y=5.56, h=0.62)
    note(n(), ["[대본] \"여섯 단계마다 사람이 한 일은 '정하기'와 '확인하기'였습니다.\""])

    s = std(prs, "프롬프트 모음", "단계별로 그대로 가져다 쓸 수 있는 문장", "정리")
    table(s, ML, TOP, CW, [1.9, CW - 1.9],
          ["단계", "프롬프트"],
          [["아이디어", "걸어 다니는 3D 힐링 공간을 만들어 줘. 장소 · 할 일 · 제약(파일 없이 코드로)은 …"],
           ["비교", "○○ 사이트와 지금 만든 것을 항목별로 비교하고, 그만큼 현실감 있게 고쳐 줘"],
           ["수정", "(본 것 그대로 줄마다) … 폴더에 영상 넣었어, 이렇게 떨려. 고친 뒤 실행해서 확인해 줘"],
           ["Supabase", "여러 명이 같은 방에 들어오게 하려면? 붙여 넣을 SQL 만들어 줘. 공개 키만 쓸게"],
           ["Vercel", "버셀에 배포하는 방법 알려줘 · 방 코드로 들어오게 · 링크 썸네일 보이게"],
           ["GitHub", "깃허브에 올려 줘 (처음) · 푸시해 줘 (매번)"]],
          row_h=0.56, size=10.5)
    note(n(), ["[대본] \"마지막 장은 바로 쓸 수 있는 문장 모음입니다. ○○ 자리에 내 아이디어를 넣으면 됩니다.\""])

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
    prs.save(sys.argv[1] if len(sys.argv) > 1 else OUT)
    return prs


if __name__ == "__main__":
    B.CUR["n"] = 0
    del B.WARN[:]
    prs = build()
    print("[쉼터 제작과정] 슬라이드 %d장 · 발표자 노트 %d개" % (len(prs.slides), len(NOTES)))
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
