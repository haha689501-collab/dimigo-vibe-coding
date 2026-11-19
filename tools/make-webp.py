# -*- coding: utf-8 -*-
"""
tools/_png/lesson-0X/*.PNG  ->  docs/lessons/lesson-0X/slides/001.webp (+ thumbs/001.webp)

사용법:  python tools/make-webp.py

- PowerPoint 가 내보낸 "슬라이드1.PNG, 슬라이드2.PNG ..." 를 번호순으로 정렬해
  001.webp, 002.webp ... 로 변환합니다.
- 원본 16:9 비율을 그대로 유지합니다. (크롭/왜곡 없음)
- 썸네일(400x225)도 함께 만듭니다.
"""
import os
import re
import sys
import shutil
from PIL import Image

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PNG_ROOT = os.path.join(ROOT, "tools", "_png")
DOCS_ROOT = os.path.join(ROOT, "docs", "lessons")

SLIDE_W, SLIDE_H = 1600, 900     # 웹 표시용 (16:9)
THUMB_W, THUMB_H = 400, 225      # 썸네일 (16:9)
SLIDE_Q, THUMB_Q = 88, 74


def slide_number(filename):
    """'슬라이드12.PNG' -> 12  (숫자만 뽑아 정렬용으로 사용)"""
    nums = re.findall(r"\d+", filename)
    return int(nums[0]) if nums else 0


def convert_lesson(lesson_dir_name):
    src_dir = os.path.join(PNG_ROOT, lesson_dir_name)
    if not os.path.isdir(src_dir):
        return None

    files = [f for f in os.listdir(src_dir) if f.lower().endswith(".png")]
    files.sort(key=slide_number)
    if not files:
        return None

    out_slides = os.path.join(DOCS_ROOT, lesson_dir_name, "slides")
    out_thumbs = os.path.join(DOCS_ROOT, lesson_dir_name, "thumbs")
    for d in (out_slides, out_thumbs):
        if os.path.isdir(d):
            shutil.rmtree(d)
        os.makedirs(d, exist_ok=True)

    total_slide = 0
    total_thumb = 0
    for i, name in enumerate(files, start=1):
        img = Image.open(os.path.join(src_dir, name)).convert("RGB")

        big = img.resize((SLIDE_W, SLIDE_H), Image.LANCZOS)
        p_big = os.path.join(out_slides, "%03d.webp" % i)
        big.save(p_big, "WEBP", quality=SLIDE_Q, method=6)
        total_slide += os.path.getsize(p_big)

        small = img.resize((THUMB_W, THUMB_H), Image.LANCZOS)
        p_small = os.path.join(out_thumbs, "%03d.webp" % i)
        small.save(p_small, "WEBP", quality=THUMB_Q, method=6)
        total_thumb += os.path.getsize(p_small)

    print("%s : 슬라이드 %d장  (본문 %.1f MB / 썸네일 %.1f MB)"
          % (lesson_dir_name, len(files), total_slide / 1048576, total_thumb / 1048576))
    return len(files)


def main():
    if not os.path.isdir(PNG_ROOT):
        print("[오류] tools/_png 폴더가 없습니다. 먼저 tools/export-slides.ps1 을 실행하세요.")
        sys.exit(1)

    results = {}
    for name in sorted(os.listdir(PNG_ROOT)):
        if name.startswith("lesson-"):
            n = convert_lesson(name)
            if n:
                results[name] = n

    print("\n--- 완료 ---")
    for k, v in results.items():
        print("  %s -> slideCount: %d" % (k, v))
    print("\n이 숫자를 docs/assets/js/lessons.js 의 slideCount 에 입력하세요.")


if __name__ == "__main__":
    main()
