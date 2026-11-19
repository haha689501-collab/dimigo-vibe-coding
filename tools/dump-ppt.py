# -*- coding: utf-8 -*-
"""PPT 각 장의 글자만 뽑아 출력 (내용 점검용)"""
import sys, io
from pptx import Presentation
sys.stdout.reconfigure(encoding='utf-8')

for path in sys.argv[1:]:
    prs = Presentation(path)
    print("=" * 70)
    print("FILE:", path.split('/')[-1], "| 총", len(prs.slides), "장")
    print("=" * 70)
    for i, s in enumerate(prs.slides, 1):
        texts = []
        for sh in s.shapes:
            if sh.has_text_frame:
                t = sh.text_frame.text.strip()
                if t:
                    texts.append(t.replace('\n', ' / '))
            elif sh.has_table:
                for r in sh.table.rows:
                    texts.append(' | '.join(c.text.strip() for c in r.cells))
        body = '  ||  '.join(texts)
        print(f"[{i:02d}] {body[:600]}")
    print()
