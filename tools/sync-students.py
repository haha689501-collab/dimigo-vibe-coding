# -*- coding: utf-8 -*-
"""
학생 작품 갤러리 정리기

실행:  python tools/sync-students.py
결과:  docs/students/data.js  (갤러리 페이지가 읽는 작품 목록 — 직접 고치지 않습니다)

[폴더 규칙]
  docs/students/
  ├── 1조_김디미/            ← 학생 폴더.  "조_이름"  또는  "번호_이름"  또는  그냥 "이름"
  │   ├── 01_기획서/        ← 작품 폴더.  "번호_이름" — 앞 번호로 순서를 정합니다
  │   │   └── PRD.md
  │   ├── 02_웹서비스/
  │   │   └── index.html
  │   └── 03_인터랙티브웹/
  │       ├── index.html
  │       ├── 개발일지.md
  │       └── assets/       ← 하위 폴더는 목록에 나오지 않지만 작품에서 그대로 쓸 수 있습니다
  └── _보관/                ← 밑줄(_)로 시작하는 폴더는 갤러리에 나오지 않습니다

- 작품 폴더 바로 아래의 html · md · 이미지 · pdf 파일이 갤러리에 나옵니다.
- 작품 제목은 html 의 <title>, md 의 첫 번째 '# 제목' 에서 가져옵니다.
- 작품 설명은 작품 폴더의 '설명.txt' 에 한 줄씩 적습니다.  (파일 개수 제한 없음)
      tetris.html : 블록을 돌려 한 줄을 채우는 테트리스
      game (4).html : 시험 기간 보스 레이드 게임
  설명이 비어 있으면 html 의 <meta name="description"> 을 대신 씁니다.
  설명.txt 자체는 갤러리에 나오지 않습니다.
- API 키 · 전화번호 · 이메일 · 주민번호처럼 보이는 내용이 있으면 경고합니다.
  (GitHub Pages 는 누구나 볼 수 있습니다. 경고가 나오면 올리기 전에 꼭 확인하세요.)
"""
import datetime
import html as htmllib
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = os.path.join(ROOT, "docs", "students")
OUT = os.path.join(BASE, "data.js")
DAYS = 5
PUBLISH_DAYS = 5                 # 이 번호까지만 갤러리에 올립니다.
HIDE_STUDENTS = {"00_예시학생"}   # 폴더 규칙 안내용 예시 — 갤러리에는 안 나옵니다

# 갤러리에서 번호 옆에 붙는 작품 이름입니다. 폴더 이름(01_기획서 …)과 맞춰 두세요.
# 작품이 한 개라도 올라온 번호만 갤러리에 나오므로, 아직 없는 칸은 그냥 두면 됩니다.
LABELS = {
    1: "기획서",
    2: "웹서비스",
    3: "인터랙티브 웹",
    4: "개인 프로젝트",
    5: "팀 프로젝트",
}

SHOW_EXT = {"html", "htm", "md", "png", "jpg", "jpeg", "gif", "webp", "svg", "pdf", "txt",
            "url"}          # 배포 사이트 바로가기 (.url) → 새 탭으로 엽니다
# zip(소스 묶음) · exe(실행 파일)는 갤러리에 내보이지 않고 학생 폴더에 백업으로만 둡니다.
# (필요하면 "zip" · "exe" 를 SHOW_EXT 에 다시 넣으면 내려받기로 보입니다)
DESC_FILE = "설명.txt"
TEAM_FILE = "팀.txt"      # docs/students/팀.txt — "팀이름 : 이름, 이름" 한 줄에 한 팀
SKIP_NAMES = {"thumbs.db", "desktop.ini", ".ds_store", DESC_FILE}
MD_EMBED_LIMIT = 200 * 1024      # 이보다 작은 md 는 data.js 에 내용을 넣어 둡니다 (파일을 더블클릭해 열어도 보이게)
BIG_FILE = 10 * 1024 * 1024

PRIVACY = [
    ("API 키 (Anthropic)", re.compile(r"sk-ant-[A-Za-z0-9_\-]{10,}")),
    ("API 키 (OpenAI 형식)", re.compile(r"\bsk-[A-Za-z0-9]{20,}")),
    ("API 키 (Google)", re.compile(r"AIza[0-9A-Za-z_\-]{30,}")),
    ("API 키라는 말", re.compile(r"(api[_-]?key|apikey)\s*[:=]\s*['\"][^'\"]{8,}", re.I)),
    ("휴대폰 번호", re.compile(r"01[016789][-\s.]?\d{3,4}[-\s.]?\d{4}")),
    ("이메일 주소", re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")),
    ("주민등록번호", re.compile(r"\b\d{6}-[1-4]\d{6}\b")),
]


def natural_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def parse_student(folder):
    """'1조_김디미' → (조, 이름) · '03_디미' → ('', '디미') · '디미' → ('', '디미')"""
    head, sep, rest = folder.partition("_")
    if sep and rest:
        if re.fullmatch(r"\d+\s*조", head):
            return head.replace(" ", ""), rest
        if head.isdigit():
            return "", rest
    return "", folder


def parse_day(folder):
    """'01_기획서' → 1 · 예전 이름 'day1' · '1일차' 도 읽습니다"""
    m = re.match(r"(?:day|d)?\s*0*(\d{1,2})(?:_|\s*일차|$)", folder.strip(), re.I)
    if not m:
        return None
    n = int(m.group(1))
    return n if 1 <= n <= DAYS else None


def read_text(path):
    for enc in ("utf-8-sig", "cp949"):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, OSError):
            continue
    return ""


def title_of(path, ext, text):
    if ext in ("html", "htm"):
        m = re.search(r"<title[^>]*>(.*?)</title>", text, re.I | re.S)
        if m:
            t = htmllib.unescape(re.sub(r"\s+", " ", m.group(1))).strip()
            if t:
                return t
    if ext == "md":
        for line in text.splitlines():
            m = re.match(r"#\s+(.+)", line.strip())
            if m:
                return m.group(1).strip()
    return os.path.splitext(os.path.basename(path))[0]


def read_descriptions(ddir):
    """설명.txt → {파일이름: 설명}.  '파일이름 : 설명' 한 줄에 하나씩 ('#' 으로 시작하면 무시)"""
    path = os.path.join(ddir, DESC_FILE)
    out = {}
    if not os.path.isfile(path):
        return out
    for line in read_text(path).splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        # 파일 이름에 ':' 가 들어갈 수 없으므로, 확장자 뒤 첫 ':' 를 기준으로 나눕니다.
        m = re.match(r"(.+?\.[A-Za-z0-9]+)\s*[:：]\s*(.*)$", line)
        if m and m.group(2).strip():
            out[m.group(1).strip().lower()] = m.group(2).strip()
    return out


def meta_description(text):
    m = re.search(r"<meta[^>]+name=[\"']description[\"'][^>]*>", text, re.I)
    if m:
        c = re.search(r"content=[\"']([^\"']*)", m.group(0), re.I)
        if c and c.group(1).strip():
            return htmllib.unescape(c.group(1)).strip()
    return ""


TEAM_KEYS = {"팀원": "members", "분류": "category", "한줄": "tagline", "소개": "desc", "미리보기": "preview"}


def read_teams(problems):
    """팀.txt → (팀 목록 [{name, members, category, tagline, desc, preview}], {학생이름: 팀이름})

    [팀이름] 으로 팀을 시작하고 그 아래 '팀원 : …' '한줄 : …' 처럼 적습니다.
    예전 형식 '팀이름 : 이름, 이름' 한 줄도 그대로 읽습니다.
    """
    teams, who, cur = [], {}, None
    path = os.path.join(BASE, TEAM_FILE)
    if not os.path.isfile(path):
        return teams, who

    def add_members(team, names):
        for n in re.split(r"[,，]", names):
            n = n.strip()
            if n:
                team["members"].append(n)
                who[n] = team["name"]

    for line in read_text(path).splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        m = re.fullmatch(r"\[(.+)\]", line)
        if m:
            cur = {"name": m.group(1).strip(), "members": []}
            teams.append(cur)
            continue
        key, sep, val = line.partition(":")
        key, val = key.strip(), val.strip()
        if not sep:
            continue
        if cur is not None and key in TEAM_KEYS:
            if key == "팀원":
                add_members(cur, val)
            elif key == "미리보기":
                if os.path.isfile(os.path.join(BASE, *val.split("/"))):
                    cur["preview"] = val
                else:
                    problems.append("%s — 미리보기 그림이 없습니다: %s" % (TEAM_FILE, val))
            else:
                cur[TEAM_KEYS[key]] = val
        else:                                   # 예전 한 줄 형식
            cur = None
            t = {"name": key, "members": []}
            teams.append(t)
            add_members(t, val)
    return teams, who


def main():
    if not os.path.isdir(BASE):
        os.makedirs(BASE)
    students, warns, problems = [], [], []
    team_order, team_of = read_teams(problems)
    total_files = 0

    for folder in sorted(os.listdir(BASE), key=natural_key):
        sdir = os.path.join(BASE, folder)
        if not os.path.isdir(sdir) or folder.startswith(("_", ".")) or folder in HIDE_STUDENTS:
            continue
        team, name = parse_student(folder)
        team = team_of.get(name, team)
        days = {}
        for dfolder in sorted(os.listdir(sdir), key=natural_key):
            ddir = os.path.join(sdir, dfolder)
            if not os.path.isdir(ddir) or dfolder.startswith(("_", ".")):
                continue
            d = parse_day(dfolder)
            if d is None:
                problems.append("%s/%s — 작품 폴더 이름을 알 수 없습니다 (예: 01_기획서)" % (folder, dfolder))
                continue
            if d > PUBLISH_DAYS:
                continue
            files = []
            descs = read_descriptions(ddir)
            for fn in sorted(os.listdir(ddir), key=lambda x: (0 if x.lower() == "index.html" else 1, natural_key(x))):
                fp = os.path.join(ddir, fn)
                if not os.path.isfile(fp) or fn.lower() in SKIP_NAMES or fn.startswith(("~$", ".")):
                    continue
                ext = fn.rsplit(".", 1)[-1].lower() if "." in fn else ""
                if ext not in SHOW_EXT:
                    continue
                size = os.path.getsize(fp)
                rel = "/".join([folder, dfolder, fn])
                text = read_text(fp) if ext in ("html", "htm", "md", "txt", "svg", "url") else ""
                item = {"name": fn, "ext": ext, "path": rel, "size": size, "title": title_of(fp, ext, text)}
                if ext == "url":
                    m = re.search(r"^\s*URL\s*=\s*(\S+)", text, re.M | re.I)
                    if not m:
                        problems.append("%s — 바로가기 주소(URL=)를 찾지 못했습니다" % rel)
                        continue
                    item["url"] = m.group(1)
                desc = descs.get(fn.lower()) or (meta_description(text) if ext in ("html", "htm") else "")
                if desc:
                    item["desc"] = desc
                if ext in ("md", "txt") and size <= MD_EMBED_LIMIT:
                    item["text"] = text
                files.append(item)
                total_files += 1
                if size > BIG_FILE:
                    warns.append("%s — 파일이 큽니다 (%.1f MB)" % (rel, size / 1048576))
                for label, rx in PRIVACY:
                    m = rx.search(text)
                    if m:
                        snippet = m.group(0)
                        shown = snippet[:6] + "…" if len(snippet) > 8 else snippet
                        warns.append("%s — %s 처럼 보이는 내용: %s" % (rel, label, shown))
            if files:
                days[str(d)] = files
            elif any(not n.startswith(".") and n.lower() not in SKIP_NAMES for n in os.listdir(ddir)):
                # 빈 작품 폴더는 아직 제출 전이므로 조용히 넘어갑니다
                problems.append("%s/%s — 보여 줄 파일이 없습니다 (html · md · 이미지)" % (folder, dfolder))
        students.append({"id": folder, "name": name, "team": team, "days": days})

    data = {
        "updated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "days": DAYS,
        "labels": {str(k): v for k, v in LABELS.items()},
        "teams": team_order,
        "students": students,
    }
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("/* tools/sync-students.py 가 만든 파일입니다. 직접 고치지 마세요. */\n")
        f.write("window.STUDENT_WORKS = ")
        json.dump(data, f, ensure_ascii=False, indent=1)
        f.write(";\n")

    print("학생 %d명 · 파일 %d개 → docs/students/data.js" % (len(students), total_files))
    if problems:
        print("\n[확인 필요] %d건" % len(problems))
        for p in problems:
            print("  - " + p)
    if warns:
        print("\n[개인정보·보안 경고] %d건 — 올리기 전에 꼭 확인하세요" % len(warns))
        for w in warns:
            print("  ! " + w)
    if not problems and not warns:
        print("문제 없음")


if __name__ == "__main__":
    main()
