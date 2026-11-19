# -*- coding: utf-8 -*-
"""
7-8차시 · 9-10차시 합본 덱을 회기별 덱 4개로 나눕니다.

실행:  python tools/split-decks-07-10.py
결과:  01_수업PPT/디미고_바이브코딩_07차시_조작품한바퀴완성.pptx     (11장)
       01_수업PPT/디미고_바이브코딩_08차시_교차검증과UI완성.pptx     (15장)
       01_수업PPT/디미고_바이브코딩_09차시_출시후보완성.pptx         ( 9장)
       01_수업PPT/디미고_바이브코딩_10차시_배포와쇼케이스.pptx       (10장)

[왜 나누는가]
  홈페이지 커리큘럼은 7 · 8 · 9 · 10차시가 각각 한 칸입니다.
  합본을 그대로 자르면 "PART 2" 부터 시작하거나 "PART 1 / 3" 같은 표기가 남습니다.
  그래서 회기별 표지를 새로 쓰고, PART 번호와 꼬리말 쪽 번호를 다시 매깁니다.

[중복 처리]
  9-10차시 덱 3~6장(우선순위 · 블루팀 패치 · UI 완성 · 크레딧)은
  7-8차시 덱 20~23장과 같은 내용입니다.
  커리큘럼상 UI 폴리싱 · 크레딧은 8차시 목표이므로 8차시에만 넣고 9차시에서는 뺐습니다.

- 원본 PPTX 는 읽기만 합니다. 수정하지 않고 새 파일만 만듭니다.
"""
import copy
import os
import re
import sys

from pptx import Presentation

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC_DIR = os.path.join(ROOT, "01_수업PPT")

DECK_78 = os.path.join(SRC_DIR, "지난버전", "소스_7-8차시합본_조작품완성과교차검증.pptx")
DECK_910 = os.path.join(SRC_DIR, "지난버전", "디미고_바이브코딩_9-10차시_출시후보와쇼케이스.pptx")


# ---------------------------------------------------------------- 공통 도구
def clone_slide(dst, src_slide):
    """기존 덱의 한 장을 그대로 새 덱에 복제한다. (도형 · 글상자 · 발표자 노트)"""
    new = dst.slides.add_slide(dst.slide_layouts[6])
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in src_slide.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))
    if src_slide.has_notes_slide:
        t = src_slide.notes_slide.notes_text_frame.text
        if t.strip():
            new.notes_slide.notes_text_frame.text = t
    return new


# 모든 덱에 공통으로 적용하는 표현 정리
#  "조별 · 조원" 처럼 굳이 반복할 필요가 없는 말을 덜어내고,
#  팀 사이의 관계를 가리킬 때만 "팀" 으로 남깁니다. (순서대로 적용)
COMMON_SUBS = [
    ("한 조가 다른 조 작품을", "한 팀이 다른 팀 작품을"),
    ("다른 조가", "다른 팀이"),
    ("다른 조의", "다른 팀의"),
    ("다른 조 ", "다른 팀 "),
    ("우리 조 저장소", "우리 저장소"),
    ("우리 조", "우리 팀"),
    ("조 PRD", "PRD"),
    ("조 작품", "미니프로젝트"),
    ("조별로 확인하고", "확인하고"),
    ("조별 발표", "팀별 발표"),
    ("조원 1명", "1명"),
    ("조원", "팀원"),
    ("조별", "팀별"),
    ("조당 ", "팀당 "),
]


def retitle(slide, rules=None, subs=COMMON_SUBS):
    """글상자 내용을 바꾼다.
       rules : {글상자 전체와 일치하는 문자열: 새 문자열 또는 None(글상자 삭제)}
       subs  : [(안에서 찾을 조각, 바꿀 조각)] — 부분 치환"""
    rules = rules or {}
    drop = []
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                key = run.text.strip()
                if key in rules:
                    rep = rules[key]
                    if rep is None:
                        drop.append(sh)
                    else:
                        run.text = rep
                    continue
                for a, b in subs:
                    if a in run.text:
                        run.text = run.text.replace(a, b)
    for sh in drop:
        sh._element.getparent().remove(sh._element)


FOOT = re.compile(r"^(.+?)\s{2}(\d{2}) / (\d{2})$")


def renumber_footer(slide, no, total):
    """꼬리말 '구역  03 / 25' 의 쪽 번호를 새 덱 기준으로 다시 쓴다."""
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                m = FOOT.match(run.text.strip())
                if m:
                    run.text = "%s  %02d / %02d" % (m.group(1), no, total)
                    return True
    return False


def build(spec):
    src = Presentation(spec["src"])
    src_slides = list(src.slides)
    out = Presentation()
    out.slide_width = src.slide_width
    out.slide_height = src.slide_height

    picks = [1] + spec["slides"]            # 1 = 표지
    total = len(picks)

    for i, no in enumerate(picks, start=1):
        sl = clone_slide(out, src_slides[no - 1])
        retitle(sl, spec["cover"] if no == 1 else spec.get("edits", {}).get(no))
        renumber_footer(sl, i, total)

    path = os.path.join(SRC_DIR, spec["out"])
    out.save(path)
    print("  %-46s %2d장" % (spec["out"], len(out.slides)))
    return len(out.slides)


# ---------------------------------------------------------------- 회기별 설정
SPECS = [
    # ------------------------------------------------------------ 7차시
    {
        "out": "디미고_바이브코딩_07차시_기능완결과디버깅.pptx",
        "src": DECK_78,
        "slides": list(range(2, 12)),                       # 2~11
        "cover": {
            "한국디지털미디어고등학교 · 7 · 8차시": "한국디지털미디어고등학교 · 7차시",
            "07 · 08": "07",
            "조 작품 완성 · 교차 검증": "기능 완결과 디버깅",
            "끝까지 동작하는 작품을 만들고, 다른 조의 사용자와 공격자 관점으로 검증합니다":
                "시작에서 결과까지 끊기지 않고 한 바퀴 도는 상태로 만듭니다",
            "검증": "통합",
            "사용자 테스트 · 레드팀 점검": "담당 기능을 하나로 합친다",
            "개선": "버그",
            "패치 · UI 완성 · 크레딧": "기록하고 원인을 찾아 고친다",
            "동작하는 작품에서, 검증된 작품으로": "흩어진 기능에서, 한 바퀴 도는 작품으로",
        },
        "edits": {
            2: {"PART 1 / 3": None,
                "7차시 · 시작부터 결과까지 끊기지 않는 한 바퀴": "시작부터 결과까지 끊기지 않는 한 바퀴"},
        },
    },
    # ------------------------------------------------------------ 8차시
    {
        "out": "디미고_바이브코딩_08차시_교차검증과UI완성.pptx",
        "src": DECK_78,
        "slides": list(range(12, 26)),                      # 12~25
        "cover": {
            "한국디지털미디어고등학교 · 7 · 8차시": "한국디지털미디어고등학교 · 8차시",
            "07 · 08": "08",
            "조 작품 완성 · 교차 검증": "교차 검증과 UI 완성",
            "끝까지 동작하는 작품을 만들고, 다른 조의 사용자와 공격자 관점으로 검증합니다":
                "다른 팀의 사용자와 공격자 관점으로 검증하고, 그 결과로 작품을 다듬습니다",
            "완결": "검증",
            "상태 · 저장 · 예외로 한 바퀴 완성": "사용자 테스트 · 레드팀 점검",
            "검증": "패치",
            "사용자 테스트 · 레드팀 점검": "우선순위를 나눠 원인을 제거",
            "개선": "완성",
            "패치 · UI 완성 · 크레딧": "UI 폴리싱 · 크레딧 · 전후 비교",
            "동작하는 작품에서, 검증된 작품으로": "만든 사람이 못 본 것을, 쓰는 사람이 찾아낸다",
        },
        "edits": {
            12: {"P A R T   2": "P A R T   1", "PART 2 / 3": "PART 1 / 2",
                 "8차시 · 다른 조의 사용자와 공격자가 되어 봅니다": "다른 팀의 사용자와 공격자가 되어 봅니다"},
            # 관찰 기록표의 '시점' 은 분·초를 적게 하면 기록이 번거로워집니다. 순서만 남깁니다.
            14: {"시점": "순서", "0:12": "①", "0:40": "②", "1:55": "③"},
            19: {"P A R T   3": "P A R T   2", "PART 3 / 3": "PART 2 / 2"},
        },
    },
    # ------------------------------------------------------------ 9차시
    {
        "out": "디미고_바이브코딩_09차시_출시후보완성.pptx",
        "src": DECK_910,
        "slides": list(range(7, 15)),                       # 7~14
        "cover": {
            "한국디지털미디어고등학교 · 9 · 10차시": "한국디지털미디어고등학교 · 9차시",
            "09 · 10": "09",
            "출시 후보 · 작품 쇼케이스": "출시 후보 완성",
            "검증 결과를 패치해 출시 후보를 만들고, 작품과 개발 과정을 공개합니다":
                "새 기능을 멈추고, 예외 상황과 다른 기기에서도 동작하는지 확정합니다",
            "패치": "동결",
            "우선순위 · 블루팀 · UI 완성": "새 기능을 멈추고 품질 확정",
            "출시 후보": "점검",
            "기능 동결 · TEST.md · v1.0-rc": "예외 상황 · 기기 · 환경",
            "쇼케이스": "문서",
            "배포 · 발표 · 상호 평가": "TEST.md · README.md · v1.0-rc",
            "검증된 작품에서, 공개하는 작품으로": "검증된 작품에서, 출시 후보로",
        },
        "edits": {
            7: {"P A R T   2": "P A R T   1", "PART 2 / 3": None,
                "9차시 후반 · 기능을 멈추고 품질을 확정합니다": "기능을 멈추고 품질을 확정합니다"},
        },
    },
    # ------------------------------------------------------------ 10차시
    {
        "out": "디미고_바이브코딩_10차시_배포와쇼케이스.pptx",
        "src": DECK_910,
        "slides": list(range(15, 24)),                      # 15~23
        "cover": {
            "한국디지털미디어고등학교 · 9 · 10차시": "한국디지털미디어고등학교 · 10차시",
            "09 · 10": "10",
            "출시 후보 · 작품 쇼케이스": "배포와 작품 쇼케이스",
            "검증 결과를 패치해 출시 후보를 만들고, 작품과 개발 과정을 공개합니다":
                "최종 버전을 배포하고, 완성 결과와 여기까지 온 과정을 발표합니다",
            "패치": "배포",
            "우선순위 · 블루팀 · UI 완성": "v1.0 · URL · QR",
            "출시 후보": "발표",
            "기능 동결 · TEST.md · v1.0-rc": "5분 · 판단의 과정을 보여 준다",
            "쇼케이스": "평가",
            "배포 · 발표 · 상호 평가": "상호 평가 · 회고",
            "검증된 작품에서, 공개하는 작품으로": "출시 후보에서, 공개하는 작품으로",
        },
        "edits": {
            15: {"P A R T   3": "P A R T   1", "PART 3 / 3": None,
                 "10차시 · 완성 결과와 발전 과정을 보여 줍니다": "완성 결과와 발전 과정을 보여 줍니다"},
            # 10회기 정리 — 6 · 7 · 8차시 구성이 바뀌었으므로 같이 고칩니다.
            22: {"발표 · 모션": "미니프로젝트 착수", "PRD 검증 · 영상": "정의 · 범위 · 분담",
                 "조 작품 · 검증": "구현 · 교차 검증", "완결 · 레드팀": "기능 완결 · 레드팀"},
        },
    },
]


if __name__ == "__main__":
    print("회기별 덱 만들기")
    for spec in SPECS:
        build(spec)
    print("\n완료 — 01_수업PPT/ 에 저장했습니다. 원본 합본은 그대로 있습니다.")
