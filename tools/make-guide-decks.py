# -*- coding: utf-8 -*-
"""
수파베이스 · 버셀 따라하기 안내서 → PPT 두 개

실행:  python tools/make-guide-decks.py
결과:  healing-space/수파베이스_따라하기.pptx
       healing-space/버셀_배포_따라하기.pptx

- 디자인은 디미고 덱 헬퍼(make-lesson03-v2 / 04-v2 / 06)를 그대로 재사용
- 내용은 healing-space 의 수파베이스_따라하기(_한글화면).md · 벌셀_배포_따라하기.md
- 캡처 그림이 healing-space/캡쳐/ 에 있으면 그 자리에 넣고, 없으면 '캡처 자리' 상자를 둡니다
"""
import importlib.util
import os
import sys

from pptx.util import Inches
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

PROJ = r"C:/Users/haha6/Desktop/hj_html/healing-space"
CAP = os.path.join(PROJ, "캡쳐")

ML, CW, SW, SH, TOP, FOOT_Y = B.ML, B.CW, B.SW, B.SH, B.TOP, B.FOOT_Y
INK, BODY, MUTED, FAINT = B.INK, B.BODY, B.MUTED, B.FAINT
ACCENT, ACCENT_D, ACCENT_L = B.ACCENT, B.ACCENT_D, B.ACCENT_L
LINE, CARD, TINT, DARK, WHITE, SOFT = B.LINE, B.CARD, B.TINT, B.DARK, B.WHITE, B.SOFT
BLUE, GREEN, PURPLE = B.BLUE, B.GREEN, B.PURPLE
rect, text, footer_rule, title_block = B.rect, B.text, B.footer_rule, B.title_block
wcard, codebox, step_row = B.wcard, B.codebox, B.step_row
grid, minicard, checklist = L4.grid, L4.minicard, L4.checklist
BLUE_D, GREEN_D, PURPLE_D, RED = L4.BLUE_D, L4.GREEN_D, L4.PURPLE_D, L4.RED
table, bar_note = L6.table, L6.bar_note


class Deck:
    def __init__(self, right):
        self.right = right
        self.notes = {}
        self.pages = []
        B.CUR["n"] = 0
        self.prs = B.new_deck()

    def n(self):
        return B.CUR["n"]

    def note(self, lines):
        self.notes[self.n()] = "\n".join(lines)

    def header(self, s):
        B.header(s, right=self.right)

    def std(self, title, sub, section):
        s = B.blank(self.prs)
        self.header(s)
        title_block(s, title, sub)
        footer_rule(s)
        self.pages.append((s, section))
        return s

    def part(self, kicker, title, sub):
        s = B.blank(self.prs)
        self.header(s)
        text(s, ML, 2.55, CW, 0.34, [[(kicker, {"bold": True, "size": 11, "color": ACCENT, "space": 1.6})]])
        text(s, ML, 2.98, CW, 0.85, [[(title, {"bold": True, "size": 38, "color": INK})]])
        rect(s, ML, 3.98, 0.80, 0.045, fill=ACCENT)
        text(s, ML, 4.22, CW * 0.8, 0.4, [[(sub, {"size": 13.5, "color": MUTED})]])
        footer_rule(s)
        return s

    def save(self, path):
        total = len(self.prs.slides)
        idx = {id(sl): k for k, sl in enumerate(self.prs.slides, 1)}
        for sl, sec in self.pages:
            k = idx[id(sl)]
            text(sl, SW - ML - 3.4, FOOT_Y + 0.12, 3.4, 0.26,
                 [[("%s  %02d / %02d" % (sec, k, total), {"size": 9, "color": FAINT, "space": 0.8})]],
                 align=PP_ALIGN.RIGHT)
        for k, sl in enumerate(self.prs.slides, start=1):
            if k in self.notes:
                sl.notes_slide.notes_text_frame.text = self.notes[k]
        self.prs.save(path)


def capture(s, x, y, w, h, fname, label):
    """캡처 그림이 있으면 넣고, 없으면 점선 느낌의 '캡처 자리' 상자"""
    path = os.path.join(CAP, fname)
    if os.path.exists(path):
        rect(s, x, y, w, h, fill=CARD, line=LINE, radius=0.04)
        s.shapes.add_picture(path, Inches(x + 0.08), Inches(y + 0.08), Inches(w - 0.16), Inches(h - 0.16))
        return
    rect(s, x, y, w, h, fill=SOFT, line=ACCENT_L, lw=1.2, radius=0.05)
    text(s, x, y + h / 2 - 0.68, w, 0.55, [[("📷", {"size": 22, "color": ACCENT_L})]], align=PP_ALIGN.CENTER)
    text(s, x + 0.2, y + h / 2 - 0.05, w - 0.4, 0.34, [[(label, {"bold": True, "size": 12.5, "color": BODY})]], align=PP_ALIGN.CENTER)
    text(s, x + 0.2, y + h / 2 + 0.32, w - 0.4, 0.3, [[("캡처 자리 · 캡쳐/" + fname, {"size": 9.5, "color": FAINT})]], align=PP_ALIGN.CENTER)


def steps_card(s, x, y, w, h, bar, label, heading, lines, note=None):
    wcard(s, x, y, w, h, bar, label, heading, lines, note=note, label_color=ACCENT_D, size=11.5)


def cover(d, big, title, sub, minis, foot):
    s = B.blank(d.prs)
    text(s, ML, 0.44, 5, 0.3, [[("A I   ×   B U I L D", {"bold": True, "size": 11, "color": ACCENT, "space": 1.3})]])
    text(s, SW - ML - 7.4, 0.46, 7.4, 0.3, [[(d.right, {"bold": True, "size": 10, "color": FAINT, "space": 1.0})]], align=PP_ALIGN.RIGHT)
    text(s, ML, 1.70, 8, 1.2, [[(big, {"bold": True, "size": 54, "color": ACCENT_L})]])
    text(s, ML, 2.84, CW, 1.0, [[(title, {"bold": True, "size": 40, "color": INK})]])
    rect(s, ML, 3.94, 0.9, 0.05, fill=ACCENT)
    text(s, ML, 4.20, CW * 0.9, 0.5, [[(sub, {"size": 17, "color": BODY})]])
    w, gap = grid(len(minis), 0.22)
    for i, (h1, h2) in enumerate(minis):
        minicard(s, ML + i * (w + gap), 5.20, w, 1.20, h1, h2, bar=ACCENT if i == 1 else ACCENT_L)
    footer_rule(s)
    text(s, ML, FOOT_Y + 0.12, 8, 0.26, [[(foot, {"size": 9.5, "color": FAINT, "space": 0.8})]])
    return s


# ====================================================================== 수파베이스
def supabase_deck():
    d = Deck("SUPABASE · 한 단계씩 따라하기")
    cover(d, "SUPABASE", "수파베이스 연결 따라하기",
          "여러 사람이 같은 공간에 들어와 서로 보이게 — 약 15분, 신용카드 필요 없음",
          [("준비", "프로젝트 · 표 만들기"), ("연결", "주소 · 공개 키 → config.js"), ("확인", "함께 모드 · 두 창 시험")],
          "supabase.com")
    d.note(["[대본] \"쉼터에 여러 명이 함께 들어오려면 Supabase라는 무료 서비스에 연결해야 합니다. 15분이면 끝납니다.\""])

    # 어느 화면?
    s = d.std("먼저 · 내 화면은 한글일까 영어일까", "두 화면 모두 순서는 같고, 버튼 이름만 다릅니다", "시작")
    table(s, ML, TOP, CW, [3.4, 4.0, CW - 7.4],
          ["찾을 것", "한글 화면", "영어 화면"],
          [["표 보는 곳", "테이블 편집기", "Table Editor"],
           ["SQL 넣는 곳", "SQL 편집기", "SQL Editor"],
           ["키가 있는 곳", "프로젝트 설정 › API", "Project Settings › API"],
           ["새 SQL 창", "새 쿼리", "New query"],
           ["실행 버튼", "실행", "Run"]], row_h=0.48)
    bar_note(s, [[("화면 언어 바꾸기: 오른쪽 위 ", {}), ("프로필 › 계정 환경설정 › 언어", {"bold": True, "color": ACCENT_D})]], y=5.60, h=0.56)
    d.note(["[대본] \"이 안내는 한글 화면 기준입니다. 영어 화면이면 오른쪽 열의 이름을 찾으세요.\""])

    # 전체 순서
    s = d.std("전체 순서", "일곱 단계 — 앞의 다섯 단계가 Supabase, 뒤의 두 단계가 확인입니다", "시작")
    step_row(s, TOP, [("1", "새 프로젝트", "지역 서울 · 암호 메모"),
                      ("2", "SQL 실행", "schema.v2.sql"),
                      ("3", "표 확인", "표 2개"),
                      ("4", "키 복사", "URL · 공개 키")], h=1.45)
    step_row(s, TOP + 1.65, [("5", "config.js", "붙여넣기 · 저장"),
                             ("6", "함께 모드", "시작 화면 아래 글씨"),
                             ("7", "두 창 시험", "서로 보이는지")], h=1.45)
    bar_note(s, [[("⚠ 비밀 키(service_role · secret)는 ", {}), ("어디에도 복사하지 않습니다", {"bold": True, "color": RED})]], y=5.70, h=0.52)
    d.note(["[대본] \"순서는 일곱 단계입니다. 키를 복사할 때 비밀 키만 조심하면 어려운 부분은 없습니다.\""])

    LW = 6.35
    RX = ML + LW + 0.3
    RW = CW - LW - 0.3

    # 1단계
    s = d.std("1단계 · 새 프로젝트 만들기", "supabase.com 로그인 → 새 프로젝트", "1단계")
    table(s, ML, TOP, LW, [1.9, LW - 1.9], ["칸", "넣을 값"],
          [["프로젝트 이름", "shelter (아무 이름이나)"],
           ["데이터베이스 암호", "아무 암호 → 복사해서 메모장에"],
           ["지역", "동북아시아(서울) 또는 (도쿄)"],
           ["요금제", "무료"],
           ["데이터 API", "✅ 켜 두기"]], row_h=0.46, size=10.5)
    bar_note(s, [[("1~2분 뒤 홈 화면에 ", {}), ("상태: 건강한", {"bold": True, "color": GREEN_D}), (" 이 보이면 준비 끝", {})]], y=6.05, h=0.52)
    s_ = s
    capture(s_, RX, TOP, RW, 3.55, "01_프로젝트만들기.png", "프로젝트 만들기 화면")
    d.note(["[대본] \"지역은 꼭 서울이나 도쿄로 고르세요. 자동으로 두면 먼 곳이 잡혀 느려질 수 있습니다.",
            "암호는 다시 볼 수 없으니 메모장에 붙여 두세요.\""])

    # 2단계
    s = d.std("2단계 · 표(테이블) 만들기", "SQL 편집기에 파일 내용을 붙여 넣고 실행합니다", "2단계")
    steps_card(s, ML, TOP, LW, 3.0, ACCENT, "순서", None,
               ["왼쪽 아이콘 줄 위에서 세 번째 >_ = SQL 편집기",
                "위쪽 '새 쿼리'",
                "supabase/schema.v2.sql 을 메모장으로 열기 → Ctrl+A → Ctrl+C",
                "편집창에 Ctrl+V → 오른쪽 아래 '실행' (Ctrl+Enter)"])
    bar_note(s, [[("결과창에 ", {}), ("\"성공. 반환된 행이 없습니다\"", {"bold": True, "color": GREEN_D}), (" — 여러 번 실행해도 안전", {})]], y=6.05, h=0.52)
    capture(s, RX, TOP, RW, 3.55, "02_SQL실행.png", "SQL 실행 화면")
    d.note(["[대본] \"SQL은 데이터베이스에 표를 만드는 명령입니다. 직접 쓸 필요 없이 준비된 파일을 통째로 붙여 넣으면 됩니다.",
            "빨간 오류가 나오면 전체가 복사됐는지 확인하세요.\""])

    # 3단계
    s = d.std("3단계 · 표가 생겼는지 확인", "테이블 편집기에서 두 개가 보이면 끝", "3단계")
    steps_card(s, ML, TOP, LW, 2.4, BLUE, "확인", None,
               ["왼쪽 아이콘 줄 위에서 두 번째 = 테이블 편집기",
                "shelter_messages — 쪽지 · 소원",
                "shelter_counters — 물 주기 · 응원 · 사랑 숫자"])
    capture(s, RX, TOP, RW, 3.55, "03_테이블확인.png", "테이블 편집기 화면")
    d.note(["[대본] \"표 두 개가 보이면 데이터베이스 준비가 끝난 것입니다.\""])

    # 4단계
    s = d.std("4단계 · 주소와 키 복사하기", "가장 중요한 단계 — 공개용 키만 복사합니다", "4단계")
    table(s, ML, TOP, LW, [2.1, LW - 2.1], ["복사할 것", "어디에 · 생김새"],
          [["프로젝트 URL", "홈 화면 제목 아래 · https://xxxx.supabase.co"],
           ["공개용 키", "프로젝트 설정 › API · anon public (게시 가능 키)"]], row_h=0.56, size=10.5)
    wcard(s, ML, TOP + 1.85, LW, 1.75, RED, "절대 복사 금지", None,
          ["비밀 키 · service_role · secret", "모든 보안 규칙을 통과해 누구나 데이터를 지울 수 있음"],
          label_color=RED, size=11)
    capture(s, RX, TOP, RW, 3.55, "04_키복사.png", "URL · 키 복사 화면")
    d.note(["[대본] \"공개용 키는 원래 브라우저에 보이도록 만든 키라 넣어도 괜찮습니다. 실제 보호는 2단계에서 넣은 규칙(RLS)이 합니다.",
            "비밀 키는 캡처할 때도 가려 주세요.\""])

    # 5단계
    s = d.std("5단계 · config.js 에 붙여넣기", "메모장이나 VS Code 로 열어 따옴표 안에 넣고 저장", "5단계")
    codebox(s, ML, TOP, LW, 1.45, "바꾸기 전 · supabase: { … } 안", ["url: '',", "anonKey: '',"])
    codebox(s, ML, TOP + 1.60, LW, 1.45, "바꾼 뒤 (Ctrl+S 저장)",
            ["url: 'https://xxxx.supabase.co',", "anonKey: 'sb_publishable_••••',"])
    bar_note(s, [[("따옴표 ' 와 끝의 쉼표 , 는 ", {}), ("지우지 않기", {"bold": True, "color": ACCENT_D})]], y=6.05, h=0.52)
    capture(s, RX, TOP, RW, 3.55, "05_config수정.png", "config.js 수정 화면")
    d.note(["[대본] \"따옴표 안쪽에만 붙여 넣으세요. 따옴표나 쉼표가 지워지면 사이트가 열리지 않습니다.\""])

    # 6단계
    s = d.std("6단계 · 연결됐는지 확인", "index.v2.html 을 더블클릭 → 시작 화면 맨 아래 글씨", "6단계")
    table(s, ML, TOP, LW, [3.5, LW - 3.5], ["보이는 글씨", "뜻"],
          [["함께 모드 · 방 \"MAIN\"", "✅ 연결 성공"],
           ["혼자 모드 · 이 브라우저에만 저장", "❌ 아직 연결 안 됨"]], row_h=0.56, size=10.5)
    bar_note(s, [[("❌ 라면 F12 › 콘솔의 빨간 글씨 확인 — ", {}), ("Supabase 연결 실패", {"bold": True, "color": RED}), (" 면 URL · 키 다시", {})]], y=6.05, h=0.52)
    capture(s, RX, TOP, RW, 3.55, "06_함께모드.png", "함께 모드 시작 화면")
    d.note(["[대본] \"시작 화면 맨 아래에 '함께 모드'가 보이면 연결된 것입니다.\""])

    # 7단계
    s = d.std("7단계 · 두 창으로 시험", "일반 창 + 시크릿 창, 같은 방 코드로 들어가기", "7단계")
    checklist(s, ML, TOP, LW, [("인원", "왼쪽 위 '지금 2명이 함께 쉬고 있어요'"),
                               ("움직임", "한쪽이 걸으면 다른 창에서도 걸어감"),
                               ("대화", "💬 인사가 양쪽 머리 위 말풍선으로"),
                               ("쪽지", "한쪽에 건 쪽지가 다른 쪽에도")], row_h=0.66, gap=0.10)
    capture(s, RX, TOP, RW, 3.55, "07_두창시험.png", "두 창 시험 화면")
    d.note(["[대본] \"이 네 가지가 되면 Supabase 작업은 끝입니다. 다음은 Vercel로 인터넷에 올리는 단계입니다.\""])

    # 자주 겪는 문제
    s = d.std("자주 겪는 문제", "대부분 주소 · 키 · 방 코드 중 하나입니다", "문제 해결")
    table(s, ML, TOP, CW, [3.2, CW - 3.2], ["증상", "해결"],
          [["계속 '혼자 모드'", "URL 끝에 / 가 붙었거나 키가 잘림 → 다시 복사"],
           ["콘솔에 Invalid API key", "비밀 키를 넣었을 가능성 → 게시 가능(anon) 키로 교체"],
           ["다른 사람이 안 보임", "방 코드가 서로 다름 → 주소의 ?room= 뒤 확인"],
           ["쪽지가 저장 안 됨", "2단계 SQL 을 일부만 실행 → 다시 실행"],
           ["1주일 뒤 안 됨", "무료 프로젝트는 오래 쉬면 일시정지 → 대시보드에서 복구"]], row_h=0.50)
    bar_note(s, [[("문제 글은 테이블 편집기 › shelter_messages › 그 줄의 ", {}), ("hidden 을 true", {"bold": True, "color": ACCENT_D}), (" 로", {})]], y=5.70, h=0.52)
    d.note(["[대본] \"막히면 이 표부터 보세요. 혼자 모드라면 대부분 주소나 키 복사 문제입니다.\""])

    out = os.path.join(PROJ, "수파베이스_따라하기.pptx")
    d.save(out)
    return d, out


# ====================================================================== 버셀
def vercel_deck():
    d = Deck("VERCEL · 배포 따라하기")
    cover(d, "VERCEL", "버셀 배포 따라하기",
          "수파베이스 연결까지 끝낸 쉼터를 Claude 와 함께 인터넷 주소로 공개합니다",
          [("처음 한 번", "깃허브 저장소 · 버셀 연결"), ("Claude", "올리기 · 배포 확인"), ("그다음부터", "\"푸시해 줘\" 한 마디")],
          "github.com · vercel.com")
    d.note(["[대본] \"올리는 일은 Claude 가 합니다. 사람은 계정과 권한이 필요한 몇 군데만 처음 한 번 누르면 됩니다.\""])

    s = d.std("누가 무엇을 하나", "계정 · 권한이 걸린 일만 내가, 나머지는 Claude 가", "시작")
    table(s, ML, TOP, CW, [0.9, 6.2, CW - 7.1], ["순서", "할 일", "누가"],
          [["1", "깃허브 가입 → 새 저장소(리포지토리) 만들기", "나"],
           ["2", "\"비밀 키 · PPT 빼고 깃허브에 올려 줘\"", "Claude (처음 한 번 깃허브 로그인 창만 내가)"],
           ["3", "버셀 가입 → 깃허브 권한 허락 → 저장소 가져오기(Import) → Deploy", "나"],
           ["4", "버셀 Deployment Protection 끄기", "나"],
           ["5", "그다음부터 \"푸시해 줘\" → 자동 배포 → 반영 확인", "Claude"]], row_h=0.52)
    bar_note(s, [[("깃허브 데스크톱 · 명령어는 ", {}), ("필요 없습니다", {"bold": True, "color": ACCENT_D}), (" — 올리기는 Claude 가 합니다", {})]], y=5.70, h=0.52)
    d.note(["[대본] \"1, 3, 4번은 내 계정에 권한을 주는 일이라 직접 누릅니다. 2번과 5번은 Claude 에게 말하면 됩니다.\""])

    LW = 6.35
    RX = ML + LW + 0.3
    RW = CW - LW - 0.3

    s = d.std("1단계 · 깃허브에 빈 저장소 만들기", "나 · github.com 가입 후 github.com/new", "1단계")
    steps_card(s, ML, TOP, LW, 3.0, ACCENT, "순서", None,
               ["github.com 가입 (있으면 통과)",
                "github.com/new 열기",
                "Repository name: shelter",
                "Private (비공개) 권장 — 사이트는 누구나 들어올 수 있음",
                "README 등 체크는 모두 해제 → Create repository"])
    bar_note(s, [[("만든 뒤 나오는 주소 ", {}), ("https://github.com/아이디/shelter.git", {"bold": True, "color": ACCENT_D}), (" 를 복사", {})]], y=6.05, h=0.52)
    capture(s, RX, TOP, RW, 3.55, "v01_새저장소.png", "새 저장소 만들기 화면")
    d.note(["[대본] \"저장소는 비어 있어야 합니다. README 를 체크하면 처음 올릴 때 충돌이 날 수 있으니 모두 해제하세요.\""])

    s = d.std("2단계 · Claude 에게 올려 달라고 하기", "Claude · 내가 할 일은 말 한 마디와 처음 한 번 로그인", "2단계")
    codebox(s, ML, TOP, LW, 1.55, "PROMPT · Claude 에게",
            ["깃허브 shelter 저장소 만들어 놨어 (주소 붙여넣기)", "비밀 키 · PPT 빼고 올려 줘"])
    wcard(s, ML, TOP + 1.65, LW, 2.0, ACCENT, "Claude 가 하는 일", None,
          ["비밀 키(service_role · sb_secret_)가 없는지 검사", "영상 · PPT · 캡처는 올리지 않게 제외", "커밋 → 깃허브로 올리기(push)"],
          label_color=ACCENT_D, size=11)
    bar_note(s, [[("처음 한 번 ", {}), ("깃허브 로그인 창", {"bold": True, "color": ACCENT_D}), ("이 뜨면 로그인만 — 다음부터는 안 뜸", {})]], y=6.05, h=0.52)
    capture(s, RX, TOP, RW, 3.55, "v02_깃허브올라감.png", "깃허브에 파일이 올라간 화면")
    d.note(["[대본] \"깃허브 데스크톱은 설치하지 않아도 됩니다. Claude 가 이 컴퓨터에서 바로 올립니다.\""])

    s = d.std("3단계 · 버셀에 깃허브 연결", "나 · 내 깃허브를 버셀이 읽어도 되도록 허락", "3단계")
    steps_card(s, ML, TOP, LW, 3.2, ACCENT, "순서", None,
               ["vercel.com › Continue with GitHub 로 가입 · 로그인",
                "Add New… › Project",
                "깃허브 권한을 묻는 창 → Install / 허락",
                "목록의 shelter 옆 Import",
                "설정은 그대로 (Framework: Other) › Deploy → 1분 뒤 주소"])
    capture(s, RX, TOP, RW, 3.55, "v03_버셀연결.png", "버셀 Import · 배포 화면")
    d.note(["[대본] \"권한 허락은 내 계정을 여는 일이라 직접 눌러야 합니다. 설정 화면은 아무것도 바꾸지 않습니다.\""])

    s = d.std("4단계 · Deployment Protection 끄기", "나 · 안 끄면 나만 들어가지고 다른 사람은 못 들어옵니다", "4단계")
    steps_card(s, ML, TOP, LW, 2.75, RED, "순서", None,
               ["버셀 › shelter 프로젝트",
                "Settings › Deployment Protection",
                "Vercel Authentication → 끄기(Disabled)",
                "Save"])
    bar_note(s, [[("확인: ", {}), ("휴대폰(로그아웃 상태)", {"bold": True, "color": ACCENT_D}), (" 으로 주소를 열어 바로 들어가지면 성공", {})]], y=6.05, h=0.52)
    capture(s, RX, TOP, RW, 3.55, "v04_보호끄기.png", "Deployment Protection 화면")
    d.note(["[대본] \"이걸 안 끄면 다른 사람이 주소를 열 때 버셀 로그인을 하라고 나옵니다.\""])

    s = d.std("5단계 · 그다음부터는 \"푸시해 줘\"", "고칠 때마다 말 한 마디로 사이트까지 반영", "5단계")
    step_row(s, TOP, [("1", "고치기", "Claude 에게 수정 요청"),
                      ("2", "\"푸시해 줘\"", "내가 말하기"),
                      ("3", "검사 · 올리기", "Claude"),
                      ("4", "자동 배포", "버셀 · 약 30초"),
                      ("5", "반영 확인", "Claude 가 보고")], h=1.55)
    codebox(s, ML, TOP + 1.80, CW, 1.15, "PROMPT · Claude 에게", ["푸시해 줘  (반영됐는지 버셀까지 확인해서 알려 줘)"])
    bar_note(s, [[("열어 둔 창은 ", {}), ("새로고침", {"bold": True, "color": ACCENT_D}), (" · 카카오톡 안이면 창을 닫고 링크를 다시", {})]], y=5.55, h=0.52)
    d.note(["[대본] \"처음 한 번만 연결해 두면 이후로는 '푸시해 줘' 한 마디면 됩니다.\""])

    s = d.std("배포 뒤 확인", "주소 뒤에 방 코드를 붙여 공유합니다", "확인")
    table(s, ML, TOP, CW, [5.2, CW - 5.2], ["주소", "결과"],
          [["https://…vercel.app", "공개 방(MAIN)"],
           ["https://…vercel.app/?room=DIMIGO1", "DIMIGO1 방"],
           ["https://…vercel.app/room/DIMIGO1", "같은 방 (주소가 깔끔)"]], row_h=0.42)
    checklist(s, ML, TOP + 1.85, CW, [("함께 모드", "시작 화면 맨 아래가 '함께 모드' (혼자 모드면 config.js 가 안 올라감)"),
                                      ("휴대폰", "휴대폰으로도 열어서 들어가지는지"),
                                      ("두 기기", "같은 방 코드로 들어가 서로 보이는지")], row_h=0.66, gap=0.08)
    d.note(["[대본] \"세 가지를 확인하면 배포가 끝납니다.\""])

    s = d.std("수업에서 쓰는 법", "가입도 설치도 없이 주소 하나로", "활용")
    step_row(s, TOP, [("1", "주소 만들기", "…/?room=우리반코드"),
                      ("2", "공유", "링크 또는 QR"),
                      ("3", "입장", "열고 '들어가기'만")], h=1.55)
    bar_note(s, [[("방 코드가 곧 입장 코드 — ", {}), ("같은 코드끼리만 서로 보입니다", {"bold": True, "color": ACCENT_D})]], y=4.30, h=0.56)
    d.note(["[대본] \"학생은 링크를 열고 들어가기만 누르면 됩니다. 반마다 다른 방 코드를 쓰면 서로 섞이지 않습니다.\""])

    s = d.std("자주 겪는 문제 · 알아 둘 것", "대부분 파일이 빠졌거나 키 문제입니다", "문제 해결")
    table(s, ML, TOP, CW, [3.2, CW - 3.2], ["증상", "해결"],
          [["배포 주소가 404", "vercel.json 이 같이 올라갔는지 확인"],
           ["'혼자 모드' 로 뜸", "config.js 에 주소 · 공개 키를 넣고 Claude 에게 \"푸시해 줘\""],
           ["화면이 하얗게", "F12 › Console 의 빨간 글씨를 캡처해서 질문"],
           ["다른 사람이 못 들어옴", "4단계 Deployment Protection 을 끄지 않음 → 끄기"]], row_h=0.48)
    w, gap = grid(3, 0.22)
    for i, (h1, h2) in enumerate([("무료 요금제", "비상업용만 — 수업용은 해당"),
                                  ("공개 키", "config.js 에 올라가도 괜찮음"),
                                  ("비밀 키", "절대 파일에 넣지 않기")]):
        minicard(s, ML + i * (w + gap), 4.75, w, 1.10, h1, h2, bar=RED if i == 2 else ACCENT_L)
    d.note(["[대본] \"Vercel 무료 요금제는 비상업용만 쓸 수 있습니다. 수업용은 괜찮습니다.\""])

    out = os.path.join(PROJ, "버셀_배포_따라하기_v2.pptx")   # 새 파일 (예전 파일은 열려 있을 수 있어 덮어쓰지 않음)
    d.save(out)
    return d, out


if __name__ == "__main__":
    which = sys.argv[1:] or ["supabase", "vercel"]
    for fn in [f for f in (supabase_deck, vercel_deck) if f.__name__.split("_")[0] in which]:
        del B.WARN[:]
        d, out = fn()
        print("[%s] 슬라이드 %d장 · 노트 %d개" % (os.path.basename(out), len(d.prs.slides), len(d.notes)))
        probs = V3.lint(d.prs)
        for w_ in B.WARN:
            print("   카드 넘침: " + w_)
        for pr in probs:
            print("   글상자: " + pr)
        if not B.WARN and not probs:
            print("   넘침 없음")
