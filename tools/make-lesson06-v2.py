# -*- coding: utf-8 -*-
"""
6차시 PPT 생성기 v2 — 「우리 조 작품 정하기」 + 부록(GitHub 협업 · 모션그래픽)

실행:  python tools/make-lesson06-v2.py
결과:  01_수업PPT/디미고_바이브코딩_6차시_v2_우리조작품정하기.pptx   (새 파일)

[왜 v2 인가]
  홈페이지 커리큘럼의 6차시는 "우리 조 작품 정하기"(개인 작품 → 조 작품 전환)입니다.
  기존 통합본(45장)은 발표 · GitHub 협업 · 모션그래픽이라 7차시 "조별 작품"의 전제가 비어 있었습니다.
  그래서 PART 1 을 커리큘럼대로 새로 만들고, 기존 자료는 뒤에 부록으로 붙입니다.

[구조]  총 58장
  PART 1  우리 조 작품 정하기        — 새로 작성 (12장)
  PART 2  미니프로젝트 발표          — 기존 통합본 2~5장
          6차시 미션 체크 · 회고      — 새로 작성 (2장)
  부록 A  GitHub 협업                — 기존 통합본 6~27장
  부록 B  클로드 모션그래픽          — 기존 통합본 28~45장

- 5차시 · 기존 6차시 덱과 같은 디자인(머리말 · 제목 · 카드 · 남색 표 · 꼬리말 구역 표시)을 씁니다.
- 기존 PPT 는 열지도 수정하지도 않습니다. 읽어서 복제만 하고, 새 파일로 저장합니다.
"""
import copy
import importlib.util
import os
import re
import sys

from pptx import Presentation
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
L6 = _load("l6", "make-lesson06.py")          # table · bar_note 재사용
SP = _load("split", "split-decks-07-10.py")   # 표현 정리 규칙(COMMON_SUBS) 재사용

SRC = os.path.join(ROOT, "01_수업PPT", "지난버전", "소스_6차시통합_발표_GitHub협업_모션그래픽.pptx")
DECK = os.path.join("01_수업PPT", "디미고_바이브코딩_06차시_미니프로젝트착수.pptx")

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
table, bar_note = L6.table, L6.bar_note

NOTES = {}
PAGES = []          # (slide, 구역 이름)


def note(i, lines):
    NOTES[i] = "\n".join(lines)


def page(s, section):
    PAGES.append((s, section))


# ================================================================ 기존 덱 복제
def clone_slide(dst, src_slide):
    """기존 덱의 한 장을 그대로 새 덱에 복제한다. (도형 · 글상자 · 발표자 노트)"""
    new = dst.slides.add_slide(dst.slide_layouts[6])
    for sh in list(new.shapes):                      # 빈 레이아웃의 자리표시자 제거
        sh._element.getparent().remove(sh._element)
    for sh in src_slide.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))
    if src_slide.has_notes_slide:
        t = src_slide.notes_slide.notes_text_frame.text
        if t.strip():
            new.notes_slide.notes_text_frame.text = t
    return new


def apply_subs(slide, subs=None):
    """복제해 온 장에 남은 '조별 · 조원' 같은 표현을 정리한다. (7~10차시와 같은 규칙)"""
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                for a, b in (subs or SP.COMMON_SUBS):
                    if a in run.text:
                        run.text = run.text.replace(a, b)


def retitle(slide, rules):
    """슬라이드의 글상자 내용을 규칙대로 바꾼다. rules: [(정규식, 새 문자열 또는 None=글상자 삭제)]"""
    drop = []
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                for pat, rep in rules:
                    if re.fullmatch(pat, run.text.strip()):
                        if rep is None:
                            drop.append(sh)
                        else:
                            run.text = rep
                        break
    for sh in drop:
        sh._element.getparent().remove(sh._element)


# ================================================================ 슬라이드
def build():
    prs = B.new_deck()
    S = lambda: B.blank(prs)
    n = lambda: B.CUR["n"]

    # ============================================================ 1. 표지
    s = S()
    text(s, ML, 0.44, 5, 0.3, [[("A I   ×   B U I L D", {"bold": True, "size": 11, "color": ACCENT, "space": 1.3})]])
    text(s, SW - ML - 7.4, 0.46, 7.4, 0.3,
         [[("한국디지털미디어고등학교 · 6차시", {"bold": True, "size": 10, "color": FAINT, "space": 1.0})]],
         align=PP_ALIGN.RIGHT)
    text(s, ML, 1.70, 2.2, 1.2, [[("06", {"bold": True, "size": 60, "color": ACCENT_L})]])
    text(s, ML, 2.84, CW, 1.0, [[("미니프로젝트 착수", {"bold": True, "size": 40, "color": INK})]])
    rect(s, ML, 3.94, 0.9, 0.05, fill=ACCENT)
    text(s, ML, 4.20, CW * 0.9, 0.5,
         [[("개인 작품 중 하나를 베이스로 선정하고, 범위와 역할을 명세로 확정합니다", {"size": 17, "color": BODY})]])
    w, gap = grid(4, 0.22)
    for i, (h1, h2) in enumerate([("리뷰", "개인 작품을 직접 조작한다"),
                                  ("선정", "기준에 따라 베이스를 고른다"),
                                  ("정의", "완료 조건과 범위를 확정한다"),
                                  ("분담", "담당 기능과 수정 범위를 나눈다")]):
        minicard(s, ML + i * (w + gap), 5.20, w, 1.20, h1, h2, bar=ACCENT if i == 1 else ACCENT_L)
    footer_rule(s)
    text(s, ML, FOOT_Y + 0.12, 9.5, 0.26,
         [[("단독 개발에서 협업 개발로  —  7~10차시는 오늘 확정한 명세 위에서 진행합니다",
            {"size": 9.5, "color": FAINT, "space": 0.8})]])
    note(n(), [
        "[시간] 2분",
        "",
        "[대본]",
        "\"1차시부터 5차시까지는 단독 개발이었습니다. 내가 정하고, 내가 만들고, 내가 고쳤습니다.",
        "오늘부터는 협업 개발입니다. 7차시부터 10차시까지 네 번에 걸쳐 미니프로젝트 하나를 출시 수준까지 끌고 갑니다.",
        "그러려면 오늘 세 가지가 끝나야 합니다. 무엇을 만들지, 어디까지 만들지, 누가 무엇을 맡을지.\"",
        "",
        "[운영]",
        "· 팀은 미리 편성해 두고 시작합니다. 4~5인 1팀.",
        "· 뒤쪽 부록 A(GitHub 협업) · 부록 B(모션그래픽)는 진행 상황에 따라 선택해 씁니다.",
    ])

    # ============================================================ 2. PART 1
    s = S()
    B.part_slide(s, "P A R T   1", "미니프로젝트 선정과 범위 정의",
                 "개인 작품을 리뷰해 베이스를 고르고, 명세로 확정합니다", 1, 2)
    note(n(), [
        "[시간] 1분",
        "",
        "[대본]",
        "\"이번 파트의 산출물은 코드가 아니라 명세입니다. PRD 한 장.",
        "오늘 이 문서가 안 나오면 다음 시간에 각자 다른 것을 만들게 됩니다.\"",
    ])

    # ============================================================ 3. 단독 개발과 협업 개발
    s = S()
    header(s)
    title_block(s, "단독 개발과 협업 개발", "달라지는 것은 규모가 아니라 명세화의 수준입니다")
    w, gap = grid(2, 0.30)
    wcard(s, ML, TOP, w, 2.92, ACCENT_L, "1 ~ 5 차시", "단독 개발",
          ["머릿속에서 정하고 바로 만든다",
           "내가 기억하면 그것이 설계다",
           "고치고 싶으면 그 자리에서 고친다"],
          note="결정을 적지 않아도 작동했습니다")
    wcard(s, ML + w + gap, TOP, w, 2.92, ACCENT, "6 ~ 10 차시", "협업 개발",
          ["구현자 전원이 같은 명세를 봐야 한다",
           "적지 않은 결정은 다르게 기억된다",
           "무엇이 완료인지 먼저 합의해야 한다"],
          note="그래서 오늘 PRD · 완료 조건 · 분담표를 만듭니다", strong=True)
    banner(s, [[("단독 개발에서 생략해도 됐던 결정이, 협업에서는 ", {}),
                ("그대로 결함이 됩니다", {"bold": True, "color": ACCENT_D})]], y=5.56, h=0.80)
    footer_rule(s); page(s, "착수")
    note(n(), [
        "[시간] 3분",
        "",
        "[대본]",
        "\"협업 개발이 어려운 이유는 기능이 많아서가 아닙니다. 여러 사람이 서로 다른 그림을 들고 만들기 때문입니다.",
        "'버튼을 누르면 결과가 나온다'를 각자 다르게 상상하면, 합칠 때 전부 어긋납니다.",
        "그래서 오늘 적는 명세는 형식이 아니라, 다음 네 시간을 줄여 주는 장치입니다.\"",
    ])

    # ============================================================ 5. 작품 시연
    s = S()
    header(s)
    title_block(s, "작품 리뷰", "설명하지 않습니다 — 리뷰어가 직접 조작합니다")
    w, gap = grid(3, 0.24)
    wcard(s, ML, TOP, w, 2.92, ACCENT_L, "발표자", "3분 · 시연 범위",
          ["첫 화면 그대로 띄우기",
           "핵심 동작 한 번만 시연",
           "가장 아쉬운 지점 한 가지"],
          note="기능 목록을 읽지 않습니다")
    wcard(s, ML + w + gap, TOP, w, 2.92, ACCENT_L, "리뷰어", "2분 · 관찰 항목",
          ["설명 없이 먼저 직접 조작",
           "막힌 지점을 말로 표현",
           "확장하고 싶은 부분 한 가지"],
          note="\"재밌다\" 대신 어디서 멈췄는지")
    wcard(s, ML + 2 * (w + gap), TOP, w, 2.92, ACCENT, "기록", "선정 근거",
          ["작품별 핵심 흐름 한 줄",
           "작품별 확장 여지 한 줄",
           "관심이 모인 작품"],
          note="다음 장 선정의 근거가 됩니다", strong=True)
    bar_note(s, [[("인원 수 × 5분 = 약 25분", {"bold": True, "color": ACCENT_D}),
                  ("  —  순서는 자유롭게, 다만 시간은 지킵니다", {})]], y=5.60, h=0.56)
    footer_rule(s); page(s, "리뷰")
    note(n(), [
        "[시간] 25분 (안내 3분 + 리뷰 22분)",
        "",
        "[대본]",
        "\"발표자는 설명하지 마세요. 첫 화면을 띄우고 넘기면, 리뷰어가 먼저 만져 봅니다.",
        "설명을 들으면 그 작품이 실제보다 좋아 보입니다. 다음 네 시간을 걸 베이스를 고르는 자리니까,",
        "설명 없이 작동하는 만큼만 보고 판단하세요.\"",
        "",
        "[운영]",
        "· 타이머를 띄워 두게 합니다. 한 작품 5분을 넘기면 리뷰가 끝나지 않습니다.",
        "· 노트북 하나에 모여 보지 말고, 배포 주소로 각자 기기에서 열게 하면 빠릅니다.",
    ])

    # ============================================================ 6. 대표작 고르기
    s = S()
    header(s)
    title_block(s, "베이스 작품 선정", "가장 완성된 작품과 가장 확장 가능한 작품은 다릅니다")
    yy = table(s, ML, TOP, CW, [2.45, 4.35, CW - 6.80],
               ["기준", "판정 질문", "확인 방법"],
               [["핵심 흐름", "시작에서 결과까지 이미 한 번 도는가", "리뷰어가 끝까지 조작해 본다"],
                ["확장 여지", "기능 3~4개를 더 붙일 자리가 있는가", "붙일 기능을 말로 꺼내 본다"],
                ["분담 가능성", "담당 영역이 인원수만큼 나오는가", "담당 후보를 적어 본다"],
                ["잔여 기간", "4회기 안에 출시까지 갈 수 있는가", "가장 오래 걸릴 기능 1개"]],
               row_h=0.52, accent_col=None)
    banner(s, [[("다수결로 정하지 않습니다. ", {}),
                ("기준마다 근거 한 줄", {"bold": True, "color": ACCENT_D}),
                ("을 붙여 전원이 같은 이유를 말할 수 있어야 합니다", {})]], y=yy + 0.36, h=0.80)
    footer_rule(s); page(s, "선정")
    note(n(), [
        "[시간] 15분",
        "",
        "[대본]",
        "\"조심할 게 하나 있습니다. 가장 예쁘고 완성된 작품이 가장 좋은 선택이 아닐 수 있습니다.",
        "이미 꽉 찬 작품은 손댈 자리가 없습니다. 반대로 거칠지만 뼈대가 단단한 작품은",
        "기능을 여러 개 붙일 수 있습니다. 지금 고르는 건 완성품이 아니라 토대입니다.\"",
        "",
        "[운영]",
        "· 의견이 갈리면: 확장 여지와 분담 가능성 두 기준으로만 다시 비교하게 합니다.",
        "· 고른 뒤 \"왜 이걸 골랐는지\" 세 줄을 PRD 맨 위에 적게 합니다. 7차시에 흔들릴 때 돌아올 지점입니다.",
    ])

    # ============================================================ 7. 대표 주제 한 문장
    s = S()
    header(s)
    title_block(s, "미니프로젝트 정의", "PRD의 첫 줄 — 여기서 어긋나면 나머지가 전부 어긋납니다")
    w, gap = grid(2, 0.30)
    wcard(s, ML, TOP, w, 2.70, ACCENT, "형식", "누구 · 무엇 · 어떻게",
          ["[누가] 쓰는 작품인가",
           "[무엇]을 할 수 있게 되는가",
           "[어떻게] 가 다른 작품과 다른가"], strong=True)
    codebox(s, ML + w + gap, TOP, w, 2.70, "예시 · 두 문장을 비교",
            ["[약함]  재미있는 타이핑 게임을 만든다",
             "",
             "[좋음]  시험 기간 학생이, 영어 단어를",
             "        손으로 쓰지 않고 외우게 하는,",
             "        10초 단위 타이핑 게임",
             "",
             "→ 누가 · 무엇 · 어떻게가 다 들어 있다"])
    codebox(s, ML, 5.22, CW, 1.68, "PROMPT · PRD 초안",
            ["아래 한 문장을 PRD.md 초안으로 만들어 줘. 코드는 아직 만들지 마.  [한 문장 붙여넣기]",
             "1. 목적 · 대상 사용자 · MUST 후보를 항목으로 나눠 주고",
             "2. 내가 정하지 않아 네가 임의로 채운 부분은 따로 표시해 줘"])
    footer_rule(s); page(s, "정의")
    note(n(), [
        "[시간] 12분",
        "",
        "[대본]",
        "\"한 문장이 안 나오면 아직 정한 게 아닙니다. '재미있는 게임'은 문장이 아니라 느낌입니다.",
        "누가 쓰는지, 무엇을 할 수 있게 되는지, 어떻게 다른지 — 세 조각을 다 넣어 보세요.",
        "그리고 마지막 프롬프트에서 중요한 건 두 번째 줄입니다. 클로드가 임의로 채운 부분을 표시하게 하면,",
        "내가 안 정한 것이 무엇인지 바로 보입니다.\"",
        "",
        "[운영]",
        "· 팀별 한 문장을 칠판이나 공유 문서에 모아 붙이면 다음 발표 때 비교가 쉽습니다.",
    ])

    # ============================================================ 8. 완료 조건 정의
    s = S()
    header(s)
    title_block(s, "완료 조건 정의", "\"만들었다\"가 아니라 \"이렇게 하면 이렇게 된다\"로 적습니다")
    yy = table(s, ML, TOP, CW, [3.30, CW - 3.30],
               ["MUST 기능", "완료 조건 · 판정 질문"],
               [["단어 10개 타이핑", "10개를 다 치면 결과 화면으로 자동 전환되는가"],
                ["점수 기록", "새로 고침 뒤에도 최고 점수가 남아 있는가"],
                ["오답 다시 보기", "틀린 단어만 모아 한 번 더 풀 수 있는가"],
                ["난이도 선택", "선택을 바꾸면 제한 시간이 실제로 달라지는가"]],
               row_h=0.50)
    w, gap = grid(2, 0.30)
    minicard(s, ML, yy + 0.32, w, 1.30, "MUST 는 3~5개",
             "다음 네 차시에 나눠 담을 양 · 인원 수와 비슷하게 맞추면 분담이 바로 나온다", bar=ACCENT_L)
    minicard(s, ML + w + gap, yy + 0.32, w, 1.30, "판정은 예 / 아니오",
             "\"자연스럽게\" · \"예쁘게\" 는 조건이 아니다 · 작성자가 아닌 사람이 혼자 판정할 수 있어야 한다", bar=ACCENT)
    footer_rule(s); page(s, "정의")
    note(n(), [
        "[시간] 15분",
        "",
        "[대본]",
        "\"완료 조건이 왜 필요한가. 9차시에 '다 됐나요?'를 판정해야 하는데, 기준이 없으면 그때 부딪칩니다.",
        "'점수 기록 기능 만들기'는 판정이 안 됩니다. '새로 고침 뒤에도 최고 점수가 남아 있는가'는 판정됩니다.",
        "지금 적는 이 표가 9차시 TEST.md 가 됩니다. 그때 다시 쓰지 않도록 지금 제대로 적으세요.\"",
        "",
        "[운영]",
        "· 표의 예시는 타이핑 게임입니다. 미니프로젝트에 맞게 통째로 바꿔 쓰게 합니다.",
        "· MUST 가 6개를 넘으면 다음 장에서 자르게 합니다.",
    ])

    # ============================================================ 9. 범위 확정
    s = S()
    header(s)
    title_block(s, "범위 확정", "넣을 것보다 빼는 것을 먼저 적습니다")
    w, gap = grid(3, 0.24)
    wcard(s, ML, TOP, w, 2.58, ACCENT, "M U S T", "이번 버전",
          ["없으면 작품이 성립하지 않는 것",
           "10차시 발표에서 반드시 보여줄 것"],
          note="3~5개 · 완료 조건 있음", strong=True)
    wcard(s, ML + w + gap, TOP, w, 2.58, BLUE, "L A T E R", "다음 버전",
          ["있으면 좋지만 없어도 완결되는 것",
           "시간이 남으면 그때 꺼낸다"],
          note="적어 두면 미련이 줄어듭니다", label_color=BLUE_D)
    wcard(s, ML + 2 * (w + gap), TOP, w, 2.58, LINE2, "N O N - G O A L", "이번에 안 함",
          ["회원가입 · 서버 · 실시간 순위 같은 것",
           "하고 싶지만 4회기에 안 끝나는 것"],
          note="\"왜 없냐\"는 질문에 답이 됩니다", label_color=MUTED)
    banner(s, [[("발표에서 설계의 수준을 가르는 건 기능 개수가 아니라 ", {}),
                ("무엇을 뺐는지와 그 이유", {"bold": True, "color": ACCENT_D}),
                ("입니다", {})]], y=5.28, h=0.80)
    footer_rule(s); page(s, "정의")
    note(n(), [
        "[시간] 10분",
        "",
        "[대본]",
        "\"NON-GOAL 을 적는 이유는 두 가지입니다. 하나는 우리가 흔들리지 않기 위해서,",
        "다른 하나는 발표 때 '왜 로그인이 없나요'라는 질문에 '안 하기로 정했습니다'라고 답하기 위해서입니다.",
        "빼기로 정한 것은 못 만든 게 아니라 결정한 겁니다.\"",
        "",
        "[운영]",
        "· 서버 · 로그인 · 실시간 동기화는 대부분 NON-GOAL 로 보내는 쪽이 안전합니다.",
    ])

    # ============================================================ 10. 핵심 흐름
    s = S()
    header(s)
    title_block(s, "핵심 사용자 흐름", "구현자 전원이 같은 화면을 그릴 수 있어야 합니다")
    step_row(s, TOP, [("01", "시작", "첫 화면에서 무엇이 보이고, 무엇을 누르면 시작되는가"),
                      ("02", "핵심 동작", "이 작품에서 가장 많이 반복하는 동작 한 가지"),
                      ("03", "결과", "동작이 끝나면 무엇이 보이고 무엇이 저장되는가"),
                      ("04", "재시작", "다시 하려면 어디를 누르고, 무엇이 남아 있는가")], h=1.70)
    w, gap = grid(2, 0.30)
    wcard(s, ML, 4.32, w, 2.52, ACCENT, "확인 방법", "말 없이 그려 보기",
          ["각자 네 칸을 따로 그려 맞춰 본다",
           "그림이 다르면 아직 합의가 안 된 것"],
          note="그림 한 장을 PRD 에 넣습니다", size=11, strong=True)
    wcard(s, ML + w + gap, 4.32, w, 2.52, ACCENT_L, "7차시 연결", "완료 정의가 됩니다",
          ["네 칸이 끊기지 않고 돌면 '한 바퀴 완성'",
           "다음 차시 첫 미션이 바로 이것"], size=11)
    footer_rule(s); page(s, "흐름")
    note(n(), [
        "[시간] 12분",
        "",
        "[대본]",
        "\"이 네 칸이 7차시의 합격 기준입니다. 시작에서 재시작까지 끊기지 않고 돌면 한 바퀴 완성입니다.",
        "지금 각자 종이에 네 칸을 그려 보세요. 그림이 다르면 같은 작품을 다르게 상상하고 있는 겁니다.",
        "지금 10분 쓰는 게, 다음 시간에 두 시간 아끼는 길입니다.\"",
        "",
        "[운영]",
        "· 종이에 그린 그림을 사진으로 찍어 PRD 에 붙이게 하면 충분합니다. 도구를 새로 배우지 않습니다.",
    ])

    # ============================================================ 11. 역할표
    s = S()
    header(s)
    title_block(s, "역할 분담", "기능만 나누지 않고, 수정 범위(파일)까지 함께 나눕니다")
    yy = table(s, ML, TOP, CW, [1.70, 3.10, 3.00, CW - 7.80],
               ["담당", "담당 MUST", "주 수정 파일", "완료 기준"],
               [["A", "단어 타이핑 · 입력 처리", "typing.js", "10개 입력 후 결과로 전환"],
                ["B", "점수 기록 · 저장", "score.js", "새로 고침 뒤에도 유지"],
                ["C", "오답 다시 보기", "review.js", "틀린 단어만 다시 출제"],
                ["D", "화면 · 디자인 시스템", "style.css · index.html", "5차시 DESIGN.md 규칙 준수"]],
               row_h=0.50)
    bar_note(s, [[("같은 파일을 두 사람이 고치면 반드시 충돌합니다  —  ", {}),
                  ("기능별로 파일을 먼저 쪼개 두세요", {"bold": True, "color": ACCENT_D})]],
             y=yy + 0.32, h=0.54)
    w, gap = grid(2, 0.30)
    minicard(s, ML, yy + 1.00, w, 1.10, "맡을 기능이 남을 때",
             "검증 담당 · 발표 담당 · 기록 담당을 역할로 둔다", bar=ACCENT_L)
    minicard(s, ML + w + gap, yy + 1.00, w, 1.10, "맡을 기능이 부족할 때",
             "MUST 를 쪼갠다 — \"점수 기록\" → 저장 / 표시", bar=ACCENT_L)
    footer_rule(s); page(s, "분담")
    note(n(), [
        "[시간] 10분",
        "",
        "[대본]",
        "\"역할을 나눌 때 기능 이름만 나누면 꼭 같은 파일을 둘이 고칩니다. 그러면 합칠 때 충돌합니다.",
        "그래서 '네가 주로 건드릴 파일'까지 적어 두세요. 파일이 겹치면 지금 쪼개면 됩니다.",
        "맡을 기능이 모자라면 검증 · 발표 · 기록을 역할로 두면 됩니다. 아무도 비지 않게.\"",
        "",
        "[운영]",
        "· 표의 파일 이름은 예시입니다. 미니프로젝트 구조에 맞게 바꾸게 합니다.",
        "· 저장소로 협업한다면 부록 A 에서 이슈 · 브랜치로 이 표를 그대로 옮깁니다.",
    ])

    # ============================================================ 12. 작업 규칙
    s = S()
    header(s)
    title_block(s, "작업 규칙", "통합 방식과 점검 시점을 먼저 정합니다")
    w, gap = grid(3, 0.24)
    wcard(s, ML, TOP, w, 2.92, ACCENT_L, "작업 단위", "한 번에 하나",
          ["요청 1회 = 수정 1개",
           "기능이 돌아가면 그 자리에서 커밋",
           "되돌릴 지점을 남기면서 전진"],
          note="4차시 Git 커밋을 그대로 씁니다")
    wcard(s, ML + w + gap, TOP, w, 2.92, ACCENT, "통합", "담당 파일 단위",
          ["각자 담당 파일만 수정",
           "공용 파일은 한 사람이 대표로 반영",
           "통합 직후 반드시 한 바퀴 돌려 본다"],
          note="통합하고 안 돌려 보는 것이 가장 위험", strong=True)
    wcard(s, ML + 2 * (w + gap), TOP, w, 2.92, ACCENT_L, "점검 시점", "차시마다 두 번",
          ["중간: 배포 주소 · 막힌 지점 한 줄",
           "끝: 한 바퀴 동작 여부",
           "미완성은 괜찮고, 미제출만 기록"],
          note="지금 상태를 정해진 시각에 공유")
    bar_note(s, [[("저장소로 함께 관리하려면 ", {}), ("부록 A · GitHub 협업", {"bold": True, "color": ACCENT_D}),
                  ("  /  발표용 15초 영상이 필요하면 ", {}), ("부록 B · 모션그래픽", {"bold": True, "color": ACCENT_D})]],
             y=5.60, h=0.56)
    footer_rule(s); page(s, "분담")
    note(n(), [
        "[시간] 8분",
        "",
        "[대본]",
        "\"통합 방식을 지금 정하지 않으면, 다음 시간 마지막 10분에 전부 몰려서 아무것도 안 합쳐집니다.",
        "원칙은 두 개입니다. 담당 파일만 고친다. 통합한 뒤에는 꼭 처음부터 한 바퀴 돌려 본다.\"",
        "",
        "[운영]",
        "· 수준에 따라 선택합니다. 파일을 주고받는 방식으로도 4회기는 충분히 돌아갑니다.",
        "· 저장소 · 브랜치 · PR 로 간다면 부록 A 를 이어서 진행합니다 (Owner 세팅 10분 필요).",
    ])

    # ============================================================ PART 2 · 기존 2~5장
    src = Presentation(SRC)
    src_slides = list(src.slides)

    for k in range(1, 5):                     # 통합본 2~5장 (PART 2 표지 + 발표 3장)
        sl = clone_slide(prs, src_slides[k])
        apply_subs(sl)
        B.CUR["n"] += 1
        if k == 1:
            retitle(sl, [(r"P A R T   1", "P A R T   2"),
                         (r"PART 1 / 5", None),
                         (r"PRD로 설계를 공유하고 검증받습니다", "오늘 확정한 명세를 다른 팀의 리뷰로 검증받습니다")])
            B.page(sl, 2, 2)
        else:
            page(sl, "발표")

    # ============================================================ 17. 미션 체크
    s = S()
    header(s)
    title_block(s, "6차시 미션 체크", "확인하고 PRD 주소를 제출합니다")
    checklist(s, ML, TOP, CW,
              [("프로젝트 정의", "한 문장 정의 · 누구를 위해 무엇을 어떻게"),
               ("완료 조건", "MUST 3~5개 · 기능별 완료 조건 (예 / 아니오로 판정)"),
               ("범위", "LATER · NON-GOAL 에 뺀 것과 이유를 적었다"),
               ("핵심 흐름", "시작 → 핵심 동작 → 결과 → 재시작 한 장"),
               ("역할 분담", "담당 MUST · 주 수정 파일 · 완료 기준")],
              row_h=0.62, gap=0.10)
    rect(s, ML, 6.00, CW, 0.74, fill=DARK, radius=0.06)
    text(s, ML + 0.34, 6.16, CW - 0.68, 0.44,
         [[("제출 · ", {"bold": True, "size": 10, "color": ACCENT_L, "space": 1.0}),
           ("PRD.md  —  다음 차시는 이 문서의 MUST 부터 바로 시작합니다",
            {"bold": True, "size": 14, "color": WHITE})]], anchor=MSO_ANCHOR.MIDDLE)
    footer_rule(s); page(s, "제출")
    note(n(), [
        "[시간] 4분",
        "",
        "[대본]",
        "\"다섯 칸이 다 채워졌는지 확인하세요. 하나라도 비어 있으면 다음 시간 시작이 늦어집니다.",
        "특히 완료 조건 — '예 / 아니오'로 판정되지 않는 문장이 있으면 지금 고치세요.\"",
        "",
        "[운영]",
        "· PRD.md 주소(저장소 또는 공유 문서)를 반 게시판에 올리게 합니다.",
    ])

    # ============================================================ 18. 회고 · 다음 차시
    s = S()
    header(s)
    title_block(s, "회고 · 다음 차시", "결정 과정에서 생긴 이견을 기록합니다")
    w, gap = grid(3, 0.26)
    for i, (num, h1, h2) in enumerate([("①", "가장 오래 걸린 결정", "선정 · 범위 · 분담 중"),
                                       ("②", "끝까지 갈린 의견", "무엇을 어떤 근거로 정했는가"),
                                       ("③", "미루기로 한 것", "LATER 로 보낸 기능 한 가지")]):
        x = ML + i * (w + gap)
        rect(s, x, TOP, w, 1.30, fill=CARD, line=LINE, radius=0.05)
        text(s, x + 0.24, TOP + 0.20, 0.5, 0.3, [[(num, {"bold": True, "size": 12, "color": ACCENT})]])
        text(s, x + 0.24, TOP + 0.50, w - 0.48, 0.34, [[(h1, {"bold": True, "size": 14, "color": INK})]])
        text(s, x + 0.24, TOP + 0.88, w - 0.48, 0.30, [[(h2, {"size": 10.5, "color": MUTED})]])
    rect(s, ML, 3.92, CW, 1.30, fill=DARK, radius=0.06)
    text(s, ML + 0.40, 4.10, CW - 0.8, 0.28,
         [[("NEXT · 7차시", {"bold": True, "size": 10, "color": ACCENT_L, "space": 1.0})]])
    text(s, ML + 0.40, 4.44, CW - 0.8, 0.50,
         [[("기능 완결과 디버깅  —  시작에서 결과까지 끊기지 않게", {"bold": True, "size": 22, "color": WHITE})]])
    bar_note(s, [[("다음 시간 시작 전: ", {}),
                  ("PRD 의 MUST 중 담당 하나를 골라 두기", {"bold": True, "color": ACCENT_D})]],
             y=5.46, h=0.62)
    footer_rule(s); page(s, "회고")
    note(n(), [
        "[시간] 3분",
        "",
        "[대본]",
        "\"오늘 적을 건 세 가지입니다. 어떤 결정이 가장 오래 걸렸는지, 끝까지 갈린 의견은 무엇이었는지,",
        "그리고 하고 싶지만 미룬 기능 하나. 이 세 줄이 10차시 발표의 '우리가 내린 결정' 부분이 됩니다.",
        "다음 시간에는 PRD 의 MUST 를 실제로 구현합니다. 맡을 하나를 정해서 오세요.\"",
    ])

    # ============================================================ 부록 A · B
    for k in range(5, 45):                    # 통합본 6~45장
        sl = clone_slide(prs, src_slides[k])
        apply_subs(sl)
        B.CUR["n"] += 1
        if k == 5:                            # 통합본 PART 2 표지 → 부록 A
            retitle(sl, [(r"P A R T   2", "부   록   A"),
                         (r"PART 2 / 5", "부록 A"),
                         (r"하나의 저장소, 여러 명의 개발자",
                          "하나의 저장소, 여러 명의 개발자 — 저장소로 협업하는 조만 진행합니다")])
        elif k == 11:                         # 통합본 PART 3 표지 → 부록 A (이어서)
            retitle(sl, [(r"P A R T   3", "부   록   A"), (r"PART 3 / 5", "부록 A")])
        elif k == 27:                         # 통합본 PART 4 표지 → 부록 B
            retitle(sl, [(r"P A R T   4", "부   록   B"),
                         (r"PART 4 / 5", "부록 B"),
                         (r"편집 프로그램 없이, 말로 영상을 만듭니다",
                          "편집 프로그램 없이, 말로 영상을 만듭니다 — 10차시 발표 티저로 씁니다")])
        elif k == 39:                         # 통합본 PART 5 표지 → 부록 B (이어서)
            retitle(sl, [(r"P A R T   5", "부   록   B"), (r"PART 5 / 5", "부록 B")])
        else:
            sec = "부록 A" if k < 27 else "부록 B"
            page(sl, sec)

    # ============================================================ 꼬리말 · 노트
    total = len(prs.slides)
    idx = {id(sl): k for k, sl in enumerate(prs.slides, 1)}

    # 복제된 슬라이드에 남아 있는 옛 꼬리말("개념  07 / 45") 제거
    old = re.compile(r"^.{0,8}\s\d{2}\s/\s\d{2}$")
    for sl, sec in PAGES:
        for sh in list(sl.shapes):
            if sh.has_text_frame and old.fullmatch(sh.text_frame.text.strip()):
                sh._element.getparent().remove(sh._element)

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
    print("[6차시 v2] %s" % DECK)
    print("  슬라이드 %d장 · 새로 만든 발표자 노트 %d개" % (len(prs.slides), len(NOTES)))
    if B.WARN:
        print("  [카드 넘침] %d건" % len(B.WARN))
        for w_ in B.WARN:
            print("    " + w_)
    else:
        print("  카드 넘침 없음")
