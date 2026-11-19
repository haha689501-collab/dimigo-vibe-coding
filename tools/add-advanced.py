# -*- coding: utf-8 -*-
"""
각 차시 덱 마지막에 「심화 과제」 한 장을 넣습니다.

실행:  python tools/add-advanced.py
결과:  01_수업PPT/ 의 1~10차시 덱이 각각 1장씩 늘어납니다.

[왜 넣는가]
  디미고 학생 기준으로 보면 1~3차시는 이미 아는 내용이 많고, 가장 어려운 7·8차시도
  여유가 있습니다. 먼저 끝낸 학생이 할 일이 없으면 수업이 늘어집니다.
  그래서 회기마다 "먼저 끝났다면" 고를 수 있는 과제 세 개를 둡니다.

[안전 규칙]
  1) 원본은 처음 한 번 01_수업PPT/지난버전/심화적용전_*.pptx 로 백업합니다.
     (이미 백업이 있으면 덮어쓰지 않습니다 — 최초 원본을 지킵니다)
  2) 이미 심화 과제 장이 있으면 지우고 다시 넣습니다. 여러 번 돌려도 1장만 유지됩니다.
  3) 쪽 번호 꼬리말이 있는 덱(5차시 이후)은 전체 번호를 다시 계산합니다.
     1~4차시는 꼬리말이 없으므로 건드리지 않습니다.
"""
import importlib.util
import os
import re
import shutil
import sys

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC_DIR = os.path.join(ROOT, "01_수업PPT")
BACKUP_DIR = os.path.join(SRC_DIR, "지난버전")
sys.path.insert(0, HERE)


def _load(name, filename):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, filename))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


B = _load("l3v2", "make-lesson03-v2.py")
L4 = _load("l4v2", "make-lesson04-v2.py")
L6 = _load("l6", "make-lesson06.py")

ML, CW, SW, TOP, FOOT_Y = B.ML, B.CW, B.SW, B.TOP, B.FOOT_Y
INK, BODY, MUTED, FAINT = B.INK, B.BODY, B.MUTED, B.FAINT
ACCENT, ACCENT_D, ACCENT_L = B.ACCENT, B.ACCENT_D, B.ACCENT_L
LINE, CARD, DARK, WHITE = B.LINE, B.CARD, B.DARK, B.WHITE
rect, text, header, footer_rule, title_block = B.rect, B.text, B.header, B.footer_rule, B.title_block
wcard = B.wcard
grid = L4.grid
bar_note = L6.bar_note

TITLE = "심화 과제"
SUBTITLE = "먼저 끝났다면 — 셋 중 하나를 골라 진행합니다"


# ================================================================ 회기별 과제
#  (파일 찾기 패턴, 꼬리말 구역 이름, [ (제목, [설명 두 줄]) × 3 ], 아래 띠 문구)
TASKS = [
    ("*_01차시*.pptx", "심화", [
        ("제약을 세 개 더", ["같은 요청에 글자 수 · 색 · 동작 제약을 추가한다",
                             "무엇이 달라졌는지 한 줄로 기록한다"]),
        ("역할을 부여하면", ["\"너는 접근성 전문가다\" 를 앞에 붙여 본다",
                             "결과에서 새로 생긴 항목을 찾는다"]),
        ("일부러 모호하게", ["조건을 비워 둔 채 요청해 본다",
                             "AI 가 임의로 채운 곳을 전부 표시한다"]),
    ], "세 가지 모두 '요청이 어떻게 결과를 바꾸는가' 를 확인하는 실험입니다"),

    ("*_02차시*.pptx", "심화", [
        ("NON-GOAL 다섯 개", ["이번에 하지 않을 것을 다섯 개까지 늘린다",
                               "각각 왜 뺐는지 한 줄씩 적는다"]),
        ("완료 조건 다시 쓰기", ["모든 기능을 예 / 아니오로 판정되게 고친다",
                                 "\"잘 동작한다\" 같은 문장을 없앤다"]),
        ("남에게 주고 확인", ["PRD 만 주고 화면을 그려 보게 한다",
                               "내 그림과 다른 지점을 찾는다"]),
    ], "PRD 는 내가 이해하는 문서가 아니라, 남이 같은 그림을 그리는 문서입니다"),

    ("*_03차시*.pptx", "심화", [
        ("세 번째 결과", ["제약만 바꿔 C 안을 하나 더 만든다",
                          "A · B · C 를 같은 기준으로 비교한다"]),
        ("코드 구조 비교", ["두 결과의 파일 · 함수 구성 차이를 찾는다",
                            "왜 다르게 짰을지 한 줄로 설명한다"]),
        ("평가표 만들기", ["비교 기준을 직접 5개 정해 점수화한다",
                           "점수가 낮은 항목을 수정 요청으로 바꾼다"]),
    ], "고르는 눈이 곧 설계 능력입니다 — 근거 없이 고른 선택은 다음에 흔들립니다"),

    ("*_04차시*.pptx", "심화", [
        ("계획을 깨 보기", ["받은 계획에서 한 단계를 빼고 진행시킨다",
                            "무엇이 무너지는지 확인한다"]),
        ("커밋 쪼개기", ["하나의 기능을 세 번에 나눠 커밋한다",
                         "중간 지점으로 되돌렸다가 다시 온다"]),
        ("CLAUDE.md 규칙", ["프로젝트 규칙 세 줄을 파일로 남긴다",
                             "다음 요청에서 지켜지는지 확인한다"]),
    ], "AI 를 잘 쓰는 것보다, 되돌릴 수 있게 만드는 것이 먼저입니다"),

    ("*_05차시*.pptx", "심화", [
        ("다크 모드", ["같은 구조로 어두운 토큰 세트를 추가한다",
                       "전환해도 레이아웃이 깨지지 않게 한다"]),
        ("위반 검사", ["토큰을 쓰지 않은 색 · 간격을 찾는 검사를 요청한다",
                       "찾은 위반을 전부 토큰으로 교체한다"]),
        ("컴포넌트 분리", ["반복되는 UI 하나를 재사용 형태로 뺀다",
                           "세 곳 이상에서 같은 것을 쓰게 만든다"]),
    ], "디자인 시스템의 가치는 예뻐지는 것이 아니라, 고칠 곳이 한 곳이 되는 것입니다"),

    ("*_06차시*.pptx", "심화", [
        ("범위를 흔들어 보기", ["NON-GOAL 하나를 MUST 로 올려 본다",
                                "대신 무엇이 밀리는지 계산한다"]),
        ("완료 조건 교차 판정", ["다른 팀에 완료 조건만 건넨다",
                                 "같은 결론이 나오는지 확인한다"]),
        ("파일 경계 검증", ["역할표의 담당 파일이 실제로 안 겹치는지 본다",
                            "겹치면 지금 쪼갠다"]),
    ], "오늘 명세가 흔들리는지 지금 확인해야, 7차시에 흔들리지 않습니다"),

    ("*_07차시*.pptx", "심화", [
        ("상태 통합", ["흩어진 전역 변수를 하나의 state 로 모은다",
                       "전이 규칙을 함수 하나로 묶는다"]),
        ("저장 손상 복구", ["저장된 값을 일부러 망가뜨려 본다",
                            "멈추지 않고 기본값으로 복구되게 한다"]),
        ("버그 한 건 추적", ["재현 → 원인 → 수정까지 기록으로 남긴다",
                             "같은 원인의 다른 지점도 함께 고친다"]),
    ], "증상 하나를 고치는 것과, 원인 하나를 없애는 것은 다릅니다"),

    ("*_08차시*.pptx", "심화", [
        ("출력 지점 전수 점검", ["입력이 그대로 화면에 나오는 곳을 전부 찾는다",
                                 "한 곳도 빠짐없이 이스케이프한다"]),
        ("저장값 위조", ["저장된 점수 · 기록을 직접 바꿔 본다",
                         "조작을 막거나, 막을 수 없음을 기록한다"]),
        ("포기 지점만 뽑기", ["관찰 기록에서 중단된 순간만 모은다",
                              "그 순서대로 우선순위를 다시 짠다"]),
    ], "공격자의 눈으로 한 번 보면, 사용자의 눈으로도 다시 보입니다"),

    ("*_09차시*.pptx", "심화", [
        ("잘못된 입력 10개", ["빈값 · 공백 · 아주 긴 값 · 특수문자를 만든다",
                              "전부 멈추지 않고 통과하게 만든다"]),
        ("나쁜 환경에서", ["느린 네트워크 · 작은 화면에서 열어 본다",
                           "첫 화면이 뜨는 시간을 적는다"]),
        ("교차 판정", ["TEST.md 판정을 다른 팀에게 맡긴다",
                       "결과가 다르면 조건 문장을 고친다"]),
    ], "출시 후보는 '되는 것' 이 아니라 '안 깨지는 것' 으로 판정합니다"),

    ("*_10차시*.pptx", "심화", [
        ("1분 안에 시작", ["처음 보는 사람에게 주소만 준다",
                           "1분 안에 시작하지 못하면 첫 화면을 고친다"]),
        ("5분을 3분으로", ["발표에서 뺄 수 있는 것을 전부 뺀다",
                           "남은 것이 진짜 핵심인지 확인한다"]),
        ("다음 버전 계획", ["LATER 에서 세 개를 꺼내 우선순위를 매긴다",
                            "각각 왜 그 순서인지 한 줄로 적는다"]),
    ], "끝낸 것보다, 무엇을 끝내지 않기로 했는지가 다음 프로젝트를 만듭니다"),
]


# ================================================================ 도구
FOOT = re.compile(r"^(.+?)\s{2}(\d{1,3}) / (\d{1,3})$")


def has_footer(prs):
    """쪽 번호 꼬리말을 쓰는 덱인지 본다."""
    for s in prs.slides:
        for sh in s.shapes:
            if sh.has_text_frame and FOOT.match(sh.text_frame.text.strip()):
                return True
    return False


def renumber(prs):
    total = len(prs.slides)
    for i, s in enumerate(prs.slides, start=1):
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    m = FOOT.match(run.text.strip())
                    if m:
                        run.text = "%s  %02d / %02d" % (m.group(1), i, total)


def is_advanced(slide):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() == TITLE:
            return True
    return False


def drop_slide(prs, idx):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    rId = slides[idx].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    prs.part.drop_rel(rId)
    xml_slides.remove(slides[idx])


def blank_layout(prs):
    """덱마다 레이아웃 구성이 다릅니다. 빈 레이아웃을 이름으로 찾고, 없으면 마지막 장과 같은 것을 씁니다."""
    for lay in prs.slide_layouts:
        if lay.name.lower() in ("blank", "빈 화면"):
            return lay
    if len(prs.slides):
        return prs.slides[-1].slide_layout
    return prs.slide_layouts[0]


def add_slide(prs, section, tasks, note_line, footer):
    s = prs.slides.add_slide(blank_layout(prs))
    for sh in list(s.shapes):                 # 레이아웃이 끌고 온 자리표시자 제거
        sh._element.getparent().remove(sh._element)
    rect(s, 0, 0, SW, 7.5, fill=B.BG)         # 배경

    header(s)
    title_block(s, TITLE, SUBTITLE)
    w, gap = grid(3, 0.24)
    for i, (head, lines) in enumerate(tasks):
        wcard(s, ML + i * (w + gap), TOP, w, 2.52,
              ACCENT if i == 0 else ACCENT_L, "0%d" % (i + 1), head, lines,
              strong=(i == 0))
    bar_note(s, [[(note_line, {})]], y=5.24, h=0.60)
    footer_rule(s)
    if footer:
        text(s, SW - ML - 3.4, FOOT_Y + 0.12, 3.4, 0.26,
             [[("%s  00 / 00" % section, {"size": 9, "color": FAINT, "space": 0.8})]],
             align=PP_ALIGN.RIGHT)
    try:                                      # 노트 틀이 없는 덱도 있습니다
        tf = s.notes_slide.notes_text_frame
        if tf is not None:
            tf.text = "\n".join([
                "[시간] 남는 시간에만 · 따로 설명하지 않습니다",
                "",
                "[운영]",
                "· 먼저 끝낸 사람에게만 안내합니다. 전체에게 시키지 않습니다.",
                "· 결과는 기록 한 줄로 남기게 하고, 발표 때 근거로 쓰게 합니다.",
            ])
    except Exception:
        pass
    return s


# ================================================================ 실행
def apply(pattern, section, tasks, note_line):
    import glob
    hits = sorted(glob.glob(os.path.join(SRC_DIR, pattern)))
    if not hits:
        print("  SKIP  %-22s 원본 없음" % pattern)
        return
    path = max(hits, key=os.path.getmtime)
    name = os.path.basename(path)

    os.makedirs(BACKUP_DIR, exist_ok=True)
    backup = os.path.join(BACKUP_DIR, "심화적용전_" + name)
    if not os.path.exists(backup):
        shutil.copy2(path, backup)

    prs = Presentation(path)
    before = len(prs.slides)
    if is_advanced(prs.slides[-1]):
        drop_slide(prs, len(prs.slides) - 1)
    footer = has_footer(prs)
    add_slide(prs, section, tasks, note_line, footer)
    if footer:
        renumber(prs)
    prs.save(path)
    print("  %-46s %2d → %2d장%s" % (name[:44], before, len(prs.slides),
                                      "" if footer else "  (쪽번호 없는 덱)"))


if __name__ == "__main__":
    B.CUR["n"] = 0
    del B.WARN[:]
    print("심화 과제 넣기")
    for pattern, section, tasks, note_line in TASKS:
        apply(pattern, section, tasks, note_line)
    if B.WARN:
        print("\n  [카드 넘침] %d건" % len(B.WARN))
        for w_ in B.WARN:
            print("    " + w_)
    else:
        print("\n카드 넘침 없음 · 원본은 01_수업PPT/지난버전/심화적용전_*.pptx 에 백업했습니다")
