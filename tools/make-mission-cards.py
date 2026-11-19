# -*- coding: utf-8 -*-
"""
7·8차시 사이드 미션 카드 PPT 생성기 (HTML 카드 2장 → 한 덱)

실행:  python tools/make-mission-cards.py [출력경로]
결과:  디미고_바이브코딩_7-8차시_사이드미션카드.pptx  (새 파일 — 기존 덱은 건드리지 않음)

- 원본: 7-8차시_쉼터_미션카드.html · 7-8차시_AI대기시간_줄이기.html
- 디자인은 쉼터 제작과정 덱과 같은 헬퍼(make-lesson03-v2 / 04-v2 / 06) 재사용
"""
import importlib.util
import os
import sys

from pptx.enum.text import PP_ALIGN

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

OUT = os.path.join(os.path.dirname(HERE), "디미고_바이브코딩_7-8차시_사이드미션카드.pptx")
RIGHT = "7 · 8차시 · SIDE MISSION"

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
TOTAL_PARTS = 2


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
         [[("한국디지털미디어고등학교 · 7 · 8차시", {"bold": True, "size": 10, "color": FAINT, "space": 1.0})]],
         align=PP_ALIGN.RIGHT)
    text(s, ML, 1.70, 6, 1.2, [[("SIDE MISSION", {"bold": True, "size": 54, "color": ACCENT_L})]])
    text(s, ML, 2.84, CW, 1.0, [[("함께 쓰는 쉼터 · AI 대기 시간 줄이기", {"bold": True, "size": 40, "color": INK})]])
    rect(s, ML, 3.94, 0.9, 0.05, fill=ACCENT)
    text(s, ML, 4.20, CW * 0.9, 0.5,
         [[("메인 프로젝트는 터미널 1, 쉼터는 터미널 2 — 대기 시간에 병렬 진행", {"size": 17, "color": BODY})]])
    w, gap = grid(3, 0.22)
    for i, (h1, h2) in enumerate([("메인 프로젝트", "완료 정의 · 한 바퀴 동작"),
                                  ("함께 쓰는 쉼터", "실시간 접속 · 배포 · 투어"),
                                  ("대기 시간 단축", "요청 단위 · 모델 · 병렬 작업")]):
        minicard(s, ML + i * (w + gap), 5.20, w, 1.20, h1, h2, bar=ACCENT if i == 1 else ACCENT_L)
    footer_rule(s)
    note(n(), [
        "[구성] PART 1 함께 쓰는 쉼터 · PART 2 AI 대기 시간 줄이기",
        "",
        "[대본]",
        "\"오늘 메인은 각 조 프로젝트입니다. 쉼터는 AI 응답을 기다리는 동안 다른 터미널에서 하는 사이드 미션입니다.\"",
    ])

    # ============================================================ PART 1 쉼터
    s = B.blank(prs)
    part_slide(s, 1, "함께 쓰는 쉼터", "생성은 한 줄 — 평가는 다른 사용자와 서로 보이는가")
    note(n(), ["[진행] 선생님 쉼터 배포 주소 · 방 코드 공유 → 학생 몇 명 접속 → 서로 보이는 장면 시연"])

    s = std(prs, "미션 개요", "3D 공간 생성은 한 줄 프롬프트로 충분합니다", "쉼터")
    w, gap = grid(3, 0.26)
    wcard(s, ML, TOP, w, 2.70, GREEN, "한 줄로 되는 것", "생성",
          ["3D 공간 · 이동", "분위기 · 조명 · 소리", "평가 대상 아님"], label_color=GREEN_D, size=11)
    wcard(s, ML + w + gap, TOP, w, 2.70, ACCENT, "한 줄로 안 되는 것", "함께 쓰기",
          ["Supabase Realtime · 방 코드", "계정 · DB · 키 설정은 직접", "평가 대상"], label_color=ACCENT_D, size=11, strong=True)
    wcard(s, ML + 2 * (w + gap), TOP, w, 2.70, BLUE, "한 줄로 안 되는 것", "공개하기",
          ["GitHub → Vercel 배포", "링크 · 휴대폰 접속", "Push 시 자동 재배포"], label_color=BLUE_D, size=11)
    bar_note(s, [[("평가 대상은 생성이 아니라 ", {}),
                  ("링크로 들어온 다른 사용자와 서로 보이는가", {"bold": True, "color": ACCENT_D})]], y=5.40, h=0.62)
    note(n(), ["[대본] \"공간은 한 줄이면 나옵니다. 오늘 과제는 친구가 링크로 들어와서 서로 보이게 하는 것입니다.\""])

    s = std(prs, "단계와 완료 조건", "단계마다 완료 조건을 직접 확인합니다", "쉼터")
    table(s, ML, TOP, CW, [1.8, 4.6, CW - 6.4],
          ["단계", "구현 범위", "완료 조건"],
          [["기본 · 생성", "3D 공간 · 1인칭/3인칭 이동 · 분위기", "로컬 실행 · 콘솔 오류 0"],
           ["기본 · 실시간", "Supabase Realtime · 방 코드 · 위치 동기화", "창 2개 · 같은 방 코드 · 서로의 아바타 이동"],
           ["기본 · 배포", "GitHub 저장소 → Vercel 연동", "배포 주소 · 휴대폰 접속 · Push 시 자동 재배포"],
           ["도전 · 상호작용", "채팅 말풍선 · 쪽지 · 공용 오브젝트 중 1개", "새로고침 후 저장값 유지"],
           ["도전 · 디테일", "레퍼런스 스크린샷 기반 수정", "수정 전후 캡처 비교 · 의도와 일치"],
           ["보스 · 투어", "반 게시판에 초대 링크 공유", "다른 학생 3명 동시 접속 · 접속 화면 캡처"]],
          row_h=0.52, size=10.5)
    note(n(), ["[진행] 순회하며 완료 조건 직접 확인 — \"창 두 개 띄워 봐\" · \"내 휴대폰으로 들어가 볼게\""])

    s = std(prs, "병렬 작업", "메인 프로젝트를 기다리는 동안 쉼터를 진행합니다", "쉼터")
    w, gap = grid(2, 0.26)
    wcard(s, ML, TOP, w, 2.80, ACCENT, "작업 분리", "터미널 2개",
          ["메인 프로젝트와 쉼터는 폴더 분리", "터미널 1 요청 대기 중 → 터미널 2 진행", "작은 수정 → 경량 모델 · 낮은 effort"],
          label_color=ACCENT_D, size=11.5)
    wcard(s, ML + w + gap, TOP, w, 2.80, BLUE, "버전 관리", "단계마다 커밋",
          ["단계 완료마다 커밋", "3D 수정으로 깨지면 직전 커밋으로 복원", "커밋 메시지 = 완료한 단계 이름"],
          label_color=BLUE_D, size=11.5)
    bar_note(s, [[("쉼터는 필수가 아닌 ", {}), ("대기 시간용 사이드 미션", {"bold": True, "color": ACCENT_D}),
                  (" — 메인 프로젝트가 우선", {})]], y=5.50, h=0.62)
    note(n(), ["[대본] \"메인 프로젝트 요청을 보내 놓고 기다리는 시간에만 쉼터를 진행합니다.\""])

    s = std(prs, "요청 예시 · 실시간 연결", "계획을 먼저 받고, 계정 · 키 · 배포 설정은 직접 수행합니다", "쉼터")
    codebox(s, ML, TOP, CW, 3.30, "PROMPT · 실시간 연결",
            ["이 쉼터에 Supabase Realtime으로 여러 사람 접속 기능을 추가해 줘. 먼저 계획만 보여 줘.",
             "",
             "1. 방 코드(?room=CODE)가 같은 사용자끼리 위치 · 방향 공유, 접속자 아바타 표시",
             "2. 테이블이 필요하면 SQL 파일로 분리 · RLS로 읽기/추가만 허용",
             "3. 접속 정보는 config.js에 Project URL과 anon 키만 — 비어 있으면 혼자 모드로 동작",
             "4. 구현 후 창 2개로 확인하는 순서 정리"])
    bar_note(s, [[("SQL은 ", {}), ("Supabase SQL Editor에 붙여 넣고 Run", {"bold": True, "color": ACCENT_D}),
                  (" — 계정 · 키 · 배포 설정은 AI가 대신할 수 없음", {})]], y=5.86, h=0.56)
    note(n(), ["[진행] 계정 설정에서 막히면 선생님 Supabase · Vercel 따라하기 자료를 화면에 띄움"])

    s = std(prs, "키 관리 · 보안 규칙", "공개 키와 비밀 키를 구분합니다", "쉼터")
    w, gap = grid(2, 0.26)
    wcard(s, ML, TOP, w, 2.80, RED, "금지", "비밀 키 노출",
          ["service_role(secret) 키를 코드 · 채팅 · 캡처에 노출", ".env · 비밀번호 파일 커밋", "RLS 비활성화 상태로 배포"],
          label_color=RED, size=11.5, strong=True)
    wcard(s, ML + w + gap, TOP, w, 2.80, BLUE, "알아 둘 것", "anon 키와 RLS",
          ["anon 키는 브라우저 노출 전제 — 보호는 RLS가 담당", "방 코드는 접근 경로일 뿐 인증이 아님", "Vercel Hobby = 비상업 용도 한정"],
          label_color=BLUE_D, size=11.5)
    note(n(), ["[대본] \"service_role 키가 올라가면 누구나 데이터를 지울 수 있습니다. 코드에는 anon 키만.\""])

    s = std(prs, "제출 · 쉼터 투어", "반 게시판에 올리고 8차시 마지막에 서로 방문합니다", "쉼터")
    checklist(s, ML, TOP, CW, [("배포 주소", "Vercel 주소 · 휴대폰에서 열리는지 확인"),
                               ("방 코드 · 초대 링크", "?room=CODE 또는 /room/CODE 형식"),
                               ("도달 단계", "기본 · 도전 · 보스 중 도달 단계와 막힌 지점 1줄")],
              row_h=0.78, gap=0.14)
    bar_note(s, [[("8차시 마지막 · ", {}), ("반 쉼터 투어", {"bold": True, "color": ACCENT_D}),
                  (" — 서로의 링크로 접속 · 3명 동시 접속 시 보스 달성", {})]], y=5.40, h=0.62)
    note(n(), ["[진행] 7차시 끝 중간 제출 → 잘된 링크 2~3개 화면 공유 / 8차시 마지막 투어 · 투표"])

    # ============================================================ PART 2 대기 시간
    s = B.blank(prs)
    part_slide(s, 2, "AI 대기 시간 줄이기", "요청 크기와 설정으로 단축합니다")
    note(n(), ["[대본] \"같은 와이파이에서 선생님이 해도 똑같이 걸렸습니다. 원인은 네트워크가 아닙니다.\""])

    s = std(prs, "대기 시간의 원인", "대부분 네트워크가 아니라 모델의 추론 · 작성 시간", "대기 시간")
    step_row(s, TOP, [("1", "요청 크기", "한 번에 시키는 양"),
                      ("2", "모델 · effort", "작업에 비해 무거운 설정"),
                      ("3", "범위", "프로젝트 전체 탐색"),
                      ("4", "맥락", "길어진 대화")], h=1.50)
    bar_note(s, [[("요청 크기 · 모델 · 범위 · 맥락이 ", {}), ("곧 대기 시간", {"bold": True, "color": ACCENT_D}),
                  (" — 네 가지 모두 조절 가능", {})]], y=4.20, h=0.62)
    note(n(), ["[대본] \"텍스트 데이터는 작습니다. 기다리는 시간의 대부분은 모델이 생각하고 쓰는 시간입니다.\""])

    s = std(prs, "요청 단위 축소", "1회 요청 = 수정 1개", "대기 시간")
    table(s, ML, TOP, CW, [CW / 2, CW / 2],
          ["대신", "이렇게"],
          [["\"전체적으로 더 예쁘게 해 줘\"", "\"타이틀 화면 배경색만 #1e232b로\""],
           ["수정 5개를 한 번에 요청", "1회 요청 = 수정 1개 → 확인 → 다음"],
           ["\"고친 뒤 실행해서 확인해 줘\" (작은 수정)", "작은 수정은 직접 새로고침으로 확인"]],
          row_h=0.62, size=11, accent_col=1)
    note(n(), ["[대본] \"작게 시키면 빨리 나오고, 틀려도 되돌리기 쉽습니다.\""])

    s = std(prs, "모델 · effort 선택", "/model 에서 전환합니다", "대기 시간")
    w, gap = grid(2, 0.26)
    wcard(s, ML, TOP, w, 2.80, BLUE, "상위 모델 · 높은 effort", "설계 · 원인 분석",
          ["구조 설계", "복잡한 3D 로직", "원인 불명 버그"], label_color=BLUE_D, size=11.5)
    wcard(s, ML + w + gap, TOP, w, 2.80, GREEN, "경량 모델 · 낮은 effort", "작은 수정",
          ["Sonnet · Haiku", "색 · 위치 · 문구 · 수치 조정", "결과를 바로 눈으로 확인 가능한 작업"],
          label_color=GREEN_D, size=11.5)
    bar_note(s, [[("모든 요청을 가장 무거운 설정으로 보내지 않기", {"bold": True, "color": ACCENT_D}),
                  (" — effort 조절 위치는 버전에 따라 다름", {})]], y=5.50, h=0.62)
    note(n(), ["[진행] 학생 버전에서 /model 열어 effort 조절 위치 확인 후 안내"])

    s = std(prs, "범위 지정 · 맥락 정리", "읽을 곳을 좁히고 대화를 짧게 유지합니다", "대기 시간")
    w, gap = grid(2, 0.26)
    wcard(s, ML, TOP, w, 2.80, ACCENT, "범위 지정", "@파일 · 함수 단위",
          ["@파일명 으로 수정 대상 파일 지정", "\"player.js의 jump 함수만\"", "프로젝트 전체 탐색 시간 감소"],
          label_color=ACCENT_D, size=11.5)
    wcard(s, ML + w + gap, TOP, w, 2.80, PURPLE, "맥락 정리", "/clear · /compact",
          ["대화가 길수록 응답 지연", "기능 완료 → /clear 로 새 대화", "이어서 작업 → /compact 로 요약"],
          label_color=PURPLE_D, size=11.5)
    note(n(), ["[대본] \"기능 하나가 끝나면 대화를 새로 시작하세요. 지난 대화 전체를 매번 다시 읽습니다.\""])

    s = std(prs, "중단 · 병렬 작업", "기다리지 말고 멈추거나, 다른 일을 진행합니다", "대기 시간")
    w, gap = grid(2, 0.26)
    wcard(s, ML, TOP, w, 2.80, ACCENT, "즉시 중단", "Esc",
          ["의도와 다른 방향 → Esc 로 중단", "끝까지 기다린 뒤 되돌리기보다 빠름", "중단 후 범위를 좁혀 재요청"],
          label_color=ACCENT_D, size=11.5)
    wcard(s, ML + w + gap, TOP, w, 2.80, RED, "병렬 작업 규칙", "터미널 2개 이상",
          ["폴더 또는 파일 분리", "터미널 1 대기 중 → 터미널 2 진행", "같은 파일 동시 수정 금지 — 충돌"],
          label_color=RED, size=11.5, strong=True)
    note(n(), ["[대본] \"한 번 기다리는 시간은 같아도, 같은 시간에 두 배를 진행할 수 있습니다.\""])

    s = std(prs, "권장 작업 흐름", "작게 · 가볍게 · 좁게 · 동시에", "대기 시간")
    step_row(s, TOP, [("1", "범위 지정", "@파일 · 함수"),
                      ("2", "수정 1개", "작게 요청"),
                      ("3", "터미널 2", "대기 중 진행"),
                      ("4", "확인 · 커밋", "직접 새로고침"),
                      ("5", "/clear", "기능 완료 시")], h=1.50)
    bar_note(s, [[("작게", {"bold": True, "color": ACCENT_D}), (" 요청 · ", {}),
                  ("가벼운 모델", {"bold": True, "color": ACCENT_D}), ("로 · ", {}),
                  ("범위", {"bold": True, "color": ACCENT_D}), (" 지정 · 기다리는 동안 ", {}),
                  ("다른 터미널", {"bold": True, "color": ACCENT_D}), ("에서 진행", {})]], y=4.20, h=0.62)
    note(n(), ["[대본] \"이 다섯 단계를 반복하면 됩니다.\""])

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
    out = sys.argv[1] if len(sys.argv) > 1 else OUT
    if os.path.exists(out) and len(sys.argv) <= 1:
        # 사용자가 편집했을 수 있는 파일은 덮어쓰지 않음
        root, ext = os.path.splitext(out)
        k = 2
        while os.path.exists("%s_v%d%s" % (root, k, ext)):
            k += 1
        out = "%s_v%d%s" % (root, k, ext)
    prs.save(out)
    print("저장: " + out)
    return prs


if __name__ == "__main__":
    B.CUR["n"] = 0
    del B.WARN[:]
    prs = build()
    print("[사이드 미션] 슬라이드 %d장 · 발표자 노트 %d개" % (len(prs.slides), len(NOTES)))
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
