# -*- coding: utf-8 -*-
"""
공개용 학생 작품 갤러리 만들기  (실명 → 익명)

실행:  python tools/build-gallery.py
결과:  docs/gallery/            ← 깃허브에 올라가는 공개용 사본 (이름 가림)
       tools/_private/학생_익명번호표.txt   ← 선생님만 보는 대조표 (깃허브에 올라가지 않음)

[왜 사본을 따로 만드나요]
  docs/students/ 는 선생님 노트북에서만 보는 원본입니다. 폴더 이름·파일 이름이 학생 실명이라
  그대로 올리면 주소(URL)에 실명이 그대로 남습니다. 그래서 공개용은
    - 폴더 이름 : 홍길동  →  s01
    - 화면 표시 : 홍길동  →  홍*동   (가운데 글자 가림)
    - 파일 이름 : 홍길동_익명게시판_PRD.md  →  _익명게시판_PRD.md
    - 파일 내용 : 글 안에 적힌 이름·학번도 가림
  으로 바꿔 docs/gallery/ 에 새로 만듭니다. 원본은 건드리지 않습니다.

[쓰는 순서]
  1) docs/students/ 에 작품을 넣습니다
  2) python tools/sync-students.py      (원본 목록 정리 — 이 스크립트가 알아서 먼저 실행합니다)
  3) python tools/build-gallery.py      (공개용 사본 만들기)
  4) git add docs/gallery  →  커밋  →  푸시
  docs/gallery/ 는 매번 지우고 새로 만듭니다. 이 폴더는 직접 고치지 마세요.
"""
import datetime
import io
import json
import os
import re
import shutil
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "docs", "students")
OUT = os.path.join(ROOT, "docs", "gallery")
MAP_DIR = os.path.join(ROOT, "tools", "_private")
MAP_FILE = os.path.join(MAP_DIR, "학생_익명번호표.txt")
TEXT_EXT = {"html", "htm", "md", "txt", "svg", "css", "js", "json"}

# 파일 이름에서 지울 것들
STUDENT_NO = r"1[1-4]\d{2}"          # 1323 처럼 보이는 학번
JUNK_EDGE = " _-.·~"


def read_text(path):
    for enc in ("utf-8-sig", "cp949"):
        try:
            with io.open(path, encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, OSError):
            continue
    return None


def mask_name(name):
    """홍길동 → 홍*동 · 김디 → 김* · 남궁도연 → 남**연"""
    n = name.strip()
    if len(n) <= 1:
        return n + "*"
    if len(n) == 2:
        return n[0] + "*"
    return n[0] + "*" * (len(n) - 2) + n[-1]


def strip_names(s, names):
    """글자열에서 실명(과 붙어 있는 학번)을 지웁니다 — 파일 이름·제목에 씁니다."""
    for n in names:
        s = re.sub(STUDENT_NO + r"[\s_\-]*" + re.escape(n), "", s)
        s = s.replace(n, "")
    s = re.sub(r"\s{2,}", " ", s)
    s = re.sub(r"_{2,}", "_", s)
    return s.strip(JUNK_EDGE)


def mask_in_text(s, names):
    """글 내용 안의 실명은 지우지 않고 가립니다 (문장이 끊기지 않게)."""
    for n in names:
        s = re.sub(STUDENT_NO + r"[\s_\-]*" + re.escape(n), mask_name(n), s)
        s = s.replace(n, mask_name(n))
    return s


def clean_filename(fn, names, fallback):
    """파일 이름에서 실명·학번을 지웁니다. 남는 게 없으면 fallback 을 씁니다."""
    stem, dot, ext = fn.rpartition(".")
    if not dot:
        stem, ext = fn, ""
    clean = strip_names(stem, names)
    clean = re.sub(r"^" + STUDENT_NO + r"[\s_\-]*", "", clean).strip(JUNK_EDGE)
    if not clean:
        clean = fallback
    return clean + ("." + ext if ext else "")


def load_data():
    path = os.path.join(SRC, "data.js")
    s = read_text(path)
    if not s:
        sys.exit("docs/students/data.js 가 없습니다. python tools/sync-students.py 를 먼저 실행하세요.")
    return json.loads(s[s.index("{"):s.rindex("}") + 1])


def build_viewer(names):
    """원본 갤러리 화면(docs/students/index.html)을 그대로 쓰고, 안내 문구만 더합니다."""
    src = read_text(os.path.join(SRC, "index.html"))
    old = "<p>산출물이 있는 회차마다 만든 작품과 기록을 모았습니다. 작품을 누르면 바로 열어 볼 수 있습니다.</p>"
    new = ("<p>산출물이 있는 회차마다 만든 작품과 기록을 모았습니다. 작품을 누르면 바로 열어 볼 수 있습니다.</p>\n"
           "    <p class=\"g-privacy\">이름은 개인정보 보호를 위해 가운데 글자를 가려 두었습니다.</p>")
    if old in src:
        src = src.replace(old, new, 1)
    else:
        print("  ! 안내 문구를 넣을 자리를 찾지 못했습니다 (index.html 설명문이 바뀐 것 같습니다)")
    src = src.replace('placeholder="이름 · 조 · 작품 제목 검색"', 'placeholder="작품 제목 검색"')
    src = src.replace("<!-- tools/sync-students.py 가 만드는 작품 목록. 직접 고치지 않습니다. -->",
                      "<!-- tools/build-gallery.py 가 만드는 공개용 작품 목록. 직접 고치지 않습니다. -->")
    # 안내 문구 모양 (CSS 는 원본에 없으므로 여기서 붙입니다)
    src = src.replace("</style>",
                      ".g-privacy{margin:6px 0 0;color:var(--ink-muted);font-size:13.5px}\n</style>", 1)
    # 혹시 화면 코드 안에 실명이 들어 있으면 가립니다 (평소에는 해당 없음)
    return mask_in_text(src, names)


def main():
    # 1) 원본 목록을 먼저 새로 고칩니다
    sync = os.path.join(ROOT, "tools", "sync-students.py")
    print("[1/4] 원본 목록 정리 — tools/sync-students.py")
    r = subprocess.run([sys.executable, sync], cwd=ROOT)
    if r.returncode != 0:
        sys.exit("sync-students.py 가 실패했습니다.")

    data = load_data()
    students = data.get("students", [])
    if not students:
        sys.exit("학생 작품이 없습니다.")
    names = sorted({s["name"] for s in students}, key=len, reverse=True)  # 긴 이름부터 지워야 안전합니다
    labels = data.get("labels", {})

    # 2) 공개용 폴더를 비우고 새로 만듭니다
    print("[2/4] 공개용 폴더 새로 만들기 — docs/gallery")
    if os.path.isdir(OUT):
        shutil.rmtree(OUT)
    os.makedirs(OUT)

    pub_students, mapping, warns = [], [], []
    total_files = 0

    for i, st in enumerate(students, 1):
        sid = "s%02d" % i
        shown = mask_name(st["name"])
        mapping.append("%s  %s  (%s)" % (sid, st["name"], shown))
        pub_days = {}

        for d in sorted(st["days"], key=int):
            label = labels.get(str(d), "작품")
            fallback = re.sub(r"\s+", "_", strip_names(label, names)) or "작품"
            ddir = os.path.join(OUT, sid, "day%d" % int(d))
            os.makedirs(ddir, exist_ok=True)
            used, files = set(), []

            for f in st["days"][d]:
                src_path = os.path.join(SRC, *f["path"].split("/"))
                if not os.path.isfile(src_path):
                    warns.append("%s — 원본 파일이 없습니다" % f["path"])
                    continue
                fn = clean_filename(f["name"], names, fallback)
                stem, dot, ext = fn.rpartition(".")
                k = 2
                while fn.lower() in used:                      # 이름이 겹치면 _2, _3 을 붙입니다
                    fn = "%s_%d%s" % (stem or fn, k, ("." + ext) if dot else "")
                    k += 1
                used.add(fn.lower())
                dst_path = os.path.join(ddir, fn)

                ext_l = f["ext"]
                if ext_l in TEXT_EXT:
                    t = read_text(src_path)
                    if t is None:                              # 글자로 못 읽으면 그대로 복사합니다
                        shutil.copy2(src_path, dst_path)
                    else:
                        with io.open(dst_path, "w", encoding="utf-8") as g:
                            g.write(mask_in_text(t, names))
                else:
                    shutil.copy2(src_path, dst_path)

                title = strip_names(f.get("title", ""), names) or os.path.splitext(fn)[0]
                item = {"name": fn, "ext": ext_l, "path": "%s/day%d/%s" % (sid, int(d), fn),
                        "size": os.path.getsize(dst_path), "title": title}
                desc = mask_in_text(f.get("desc", ""), names)
                if desc:
                    item["desc"] = desc
                if "text" in f:
                    item["text"] = mask_in_text(f["text"], names)
                files.append(item)
                total_files += 1

            # 원본 일차 폴더의 하위 폴더(assets 등)도 함께 옮깁니다
            src_ddir = os.path.dirname(os.path.join(SRC, *st["days"][d][0]["path"].split("/"))) if st["days"][d] else None
            if src_ddir and os.path.isdir(src_ddir):
                for sub in sorted(os.listdir(src_ddir)):
                    sp = os.path.join(src_ddir, sub)
                    if os.path.isdir(sp) and not sub.startswith((".", "_")):
                        shutil.copytree(sp, os.path.join(ddir, sub), dirs_exist_ok=True)
                        warns.append("%s/day%s/%s/ — 하위 폴더를 그대로 옮겼습니다 (안에 이름이 없는지 확인하세요)"
                                     % (sid, d, sub))
            if files:
                pub_days[str(d)] = files

        pub_students.append({"id": sid, "name": shown, "team": st.get("team", ""), "days": pub_days})

    # 3) 공개용 목록과 화면
    print("[3/4] 목록·화면 쓰기 — docs/gallery/data.js · index.html")
    out = {"updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
           "days": data.get("days", 10), "labels": labels, "students": pub_students}
    with io.open(os.path.join(OUT, "data.js"), "w", encoding="utf-8") as f:
        f.write("/* tools/build-gallery.py 가 만든 공개용 파일입니다. 직접 고치지 마세요. */\n")
        f.write("window.STUDENT_WORKS = ")
        json.dump(out, f, ensure_ascii=False, indent=1)
        f.write(";\n")
    with io.open(os.path.join(OUT, "index.html"), "w", encoding="utf-8") as f:
        f.write(build_viewer(names))

    os.makedirs(MAP_DIR, exist_ok=True)
    with io.open(MAP_FILE, "w", encoding="utf-8") as f:
        f.write("공개용 번호 ↔ 학생 이름 대조표  (" + out["updated"] + ")\n")
        f.write("이 파일은 깃허브에 올라가지 않습니다 (.gitignore).\n\n")
        f.write("\n".join(mapping) + "\n")

    # 4) 공개용 폴더에 실명이 남아 있는지 다시 확인합니다
    print("[4/4] 실명 남았는지 확인 — docs/gallery 전체")
    leaks = []
    for root, dirs, fs in os.walk(OUT):
        for fn in fs:
            rel = os.path.relpath(os.path.join(root, fn), OUT).replace("\\", "/")
            for n in names:
                if n in rel:
                    leaks.append("%s — 경로에 '%s'" % (rel, n))
            ext = fn.rsplit(".", 1)[-1].lower() if "." in fn else ""
            if ext in TEXT_EXT:
                t = read_text(os.path.join(root, fn)) or ""
                for n in names:
                    if n in t:
                        leaks.append("%s — 내용에 '%s'" % (rel, n))

    print("\n학생 %d명 · 파일 %d개 → docs/gallery" % (len(pub_students), total_files))
    print("대조표 → tools/_private/학생_익명번호표.txt")
    if warns:
        print("\n[확인 필요] %d건" % len(warns))
        for w in warns:
            print("  - " + w)
    if leaks:
        print("\n[실명 남음] %d건 — 올리기 전에 꼭 해결하세요" % len(leaks))
        for l in leaks:
            print("  ! " + l)
    else:
        print("실명 남은 곳 없음 — 올려도 됩니다")


if __name__ == "__main__":
    main()
