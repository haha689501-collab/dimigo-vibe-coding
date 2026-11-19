# -*- coding: utf-8 -*-
"""
학생 작품 한 파일로 묶기  (실명 그대로 · 깃허브에 올리지 않음)

실행:  python tools/build-offline.py
결과:  학생작품_모아보기.html   ← 이 파일 하나만 있으면 어디서든 열립니다

[왜 필요한가요]
  docs/students/index.html 은 혼자서는 동작하지 않습니다. 화면 틀만 들어 있고
    ../assets/css/styles.css  (색·글꼴)
    data.js                   (학생·작품 목록)
    홍길동/01_기획서/...        (작품 파일)
  을 바깥에서 불러오기 때문입니다. 그래서 index.html 만 보내면 빈 화면이 됩니다.
  이 스크립트는 그 세 가지를 모두 한 파일 안에 넣어 둡니다.
  USB·메일·카톡으로 파일 하나만 보내도 그대로 열리고, 인터넷이 없어도 됩니다.

[주의]
  학생 실명이 들어 있는 파일입니다. 깃허브·블로그처럼 공개된 곳에 올리지 마세요.
  (.gitignore 에 넣어 두어 실수로 올라가지 않습니다.)
"""
import base64
import hashlib
import io
import json
import os
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "docs", "students")
CSS = os.path.join(ROOT, "docs", "assets", "css", "styles.css")
OUT = os.path.join(ROOT, "학생작품_모아보기.html")
TEXT_EXT = {"html", "htm", "md", "txt", "svg"}
NO_EMBED_SIZE = 30 * 1024 * 1024   # 이보다 큰 파일(실행 파일 등)은 보관본에 넣지 않고 안내만 합니다
MIME = {"html": "text/html", "htm": "text/html", "md": "text/markdown", "txt": "text/plain",
        "svg": "image/svg+xml", "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "gif": "image/gif", "webp": "image/webp", "pdf": "application/pdf", "zip": "application/zip"}

# 작품을 iframe 으로 열 때 넣어 주는 보호 장치.
# 한 파일 안에서 실행되면 주소가 없는 상태라 localStorage(최고 점수 저장)가 막힙니다.
# 막히면 게임이 멈추므로, 대신 쓸 임시 저장소를 미리 끼워 넣습니다.
SHIM = (
    "<script>(function(){try{window.localStorage.getItem('_');}catch(e){var m={};"
    "Object.defineProperty(window,'localStorage',{configurable:true,value:{"
    "getItem:function(k){return Object.prototype.hasOwnProperty.call(m,k)?m[k]:null;},"
    "setItem:function(k,v){m[k]=String(v);},removeItem:function(k){delete m[k];},"
    "clear:function(){m={};},key:function(i){return Object.keys(m)[i]||null;},"
    "get length(){return Object.keys(m).length;}}});}})();</script>\n"
)

# 왼쪽 위 로고. 이 파일만 열었을 때는 돌아갈 홈페이지가 없으므로 맨 위로 올라갑니다.
# (build-site-offline.py 가 이 문장을 찾아 '뒤로 가기' 로 바꿉니다 — 문장을 고치면 그쪽도 같이 고치세요)
BRAND_SOLO = ('<a class="brand" href="#" aria-label="맨 위로" '
              'onclick="window.scrollTo({top:0,behavior:\'smooth\'});return false;">')

# 한 파일 안에서 작품을 꺼내 쓰는 도우미. 원래 화면 코드의 enc() 자리를 대신합니다.
HELPER = """
<script>
/* 작품 파일을 파일 안에서 꺼내 씁니다 (주소 대신 blob 주소를 만들어 줍니다) */
var OFF = (function () {
  var F = window.OFFLINE_FILES || {}, cache = {};
  var SHIM = %s;
  /* 팀원들이 낸 똑같은 파일은 한 번만 담고, 나머지는 {a: 원본 경로} 로 가리킵니다 */
  function real(p) { return F[p] && F[p].a ? F[p].a : p; }
  function mime(p) { p = real(p); return F[p] && F[p].m || 'application/octet-stream'; }
  function url(p) {
    p = real(p);
    if (cache[p]) return cache[p];
    var f = F[p];
    if (!f) return '#';
    var blob;
    if (f.t === 't') {
      blob = new Blob([(f.m === 'text/html' ? SHIM : '') + f.d], { type: f.m + ';charset=utf-8' });
    } else {
      var bin = atob(f.d), arr = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
      blob = new Blob([arr], { type: f.m });
    }
    cache[p] = URL.createObjectURL(blob);
    return cache[p];
  }
  function html(p) { var f = F[real(p)]; return f && f.t === 't' ? SHIM + f.d : ''; }
  return { url: url, html: html, mime: mime };
})();
</script>
"""


def read_text(path):
    for enc in ("utf-8-sig", "cp949"):
        try:
            with io.open(path, encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, OSError):
            continue
    return None


def js_literal(obj):
    """<script> 안에 그대로 넣어도 깨지지 않는 JSON 문자열"""
    s = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
    return (s.replace("</", "<\\/")
             .replace(" ", "\\u2028")
             .replace(" ", "\\u2029"))


def main():
    print("[1/3] 작품 목록 정리 — tools/sync-students.py")
    if subprocess.run([sys.executable, os.path.join(ROOT, "tools", "sync-students.py")], cwd=ROOT).returncode:
        sys.exit("sync-students.py 가 실패했습니다.")

    raw = read_text(os.path.join(SRC, "data.js"))
    data = json.loads(raw[raw.index("{"):raw.rindex("}") + 1])

    print("[2/3] 작품 파일 담기")
    files, total, missing, skipped = {}, 0, [], []
    first_of, dup_count, dup_bytes = {}, 0, 0     # 내용(해시) → 처음 담은 경로
    for st in data["students"]:
        for d, fs in st["days"].items():
            for f in fs:
                p = f["path"]
                if f["ext"] == "url":            # 바로가기는 주소만 있으면 됩니다 (목록에 들어 있음)
                    continue
                if f["ext"] == "exe" or f.get("size", 0) > NO_EMBED_SIZE:
                    f["note"] = "용량이 커서 보관본에는 넣지 않았습니다. 원본 폴더(docs/students/%s)에 있습니다." % p
                    skipped.append(p)
                    continue
                fp = os.path.join(SRC, *p.split("/"))
                if not os.path.isfile(fp):
                    missing.append(p)
                    continue
                with open(fp, "rb") as g:
                    digest = hashlib.sha1(g.read()).hexdigest()
                if digest in first_of:                 # 다른 팀원이 낸 똑같은 파일 → 이름표만
                    files[p] = {"a": first_of[digest]}
                    dup_count += 1
                    dup_bytes += f.get("size", 0)
                    total += 1
                    continue
                first_of[digest] = p
                ext = f["ext"]
                m = MIME.get(ext, "application/octet-stream")
                if ext in TEXT_EXT:
                    t = read_text(fp)
                    if t is not None:
                        files[p] = {"t": "t", "m": m, "d": t}
                        total += 1
                        continue
                with open(fp, "rb") as g:
                    files[p] = {"t": "b", "m": m, "d": base64.b64encode(g.read()).decode("ascii")}
                total += 1

    # 팀 프로젝트 소개 카드의 미리보기 그림 (팀.txt 의 '미리보기')
    for t in data.get("teams", []):
        p = t.get("preview") if isinstance(t, dict) else None
        fp = p and os.path.join(SRC, *p.split("/"))
        if fp and os.path.isfile(fp):
            ext = p.rsplit(".", 1)[-1].lower()
            with open(fp, "rb") as g:
                files[p] = {"t": "b", "m": MIME.get(ext, "image/" + ext),
                            "d": base64.b64encode(g.read()).decode("ascii")}

    print("[3/3] 한 파일로 쓰기 — 학생작품_모아보기.html")
    page = read_text(os.path.join(SRC, "index.html"))
    css = read_text(CSS)

    # 1) 바깥 CSS 를 안으로
    page = page.replace('<link rel="stylesheet" href="../assets/css/styles.css">',
                        "<style>\n" + css + "\n</style>", 1)

    # 2) 바깥 메뉴(사이트의 다른 쪽으로 가는 링크)는 한 파일에서는 열 수 없으므로 뺍니다
    nav_start = page.find('<nav class="site-nav"')
    nav_end = page.find("</nav>", nav_start)
    if nav_start != -1 and nav_end != -1:
        page = page[:nav_start] + page[nav_end + 6:]

    # 2-1) 왼쪽 위 로고도 바깥 주소(../index.html)를 가리켜 빈 화면이 됩니다.
    #      이 파일만 따로 열었을 때는 갈 곳이 없으므로 '맨 위로' 로 바꿉니다.
    #      (강의 홈페이지 보관본 안에 담길 때는 build-site-offline.py 가 BRAND_SOLO 를
    #       '뒤로 가기' 로 다시 바꿔, 로고를 누르면 홈페이지로 돌아갑니다)
    brand = '<a class="brand" href="../index.html#/" aria-label="AI × BUILD 홈으로">'
    if brand not in page:
        sys.exit("화면 코드가 바뀌었습니다 — 로고 링크를 찾지 못했습니다.")
    page = page.replace(brand, BRAND_SOLO, 1)

    # 3) 보관본 안내 한 줄
    page = page.replace(
        "<p>완성한 작품과 기록을 모았습니다. 작품을 누르면 바로 열어 볼 수 있습니다.</p>",
        "<p>완성한 작품과 기록을 모았습니다. 작품을 누르면 바로 열어 볼 수 있습니다.</p>\n"
        '    <p class="g-offline">작품·목록·디자인을 모두 이 파일 하나에 담았습니다. 인터넷이 없어도 열립니다.<br>'
        "<b>학생 실명이 들어 있으니 공개된 곳에 올리지 마세요.</b></p>", 1)
    page = page.replace("</style>", ".g-offline{margin:8px 0 0;color:var(--ink-muted);font-size:13.5px;line-height:1.7}"
                                    ".g-offline b{color:var(--accent-strong)}\n</style>", 1)

    # 4) 목록 파일(data.js) 과 작품 파일을 안으로
    payload = ("<script>\nwindow.STUDENT_WORKS = " + js_literal(data) + ";\n"
               "window.OFFLINE_FILES = " + js_literal(files) + ";\n</script>\n"
               + HELPER % js_literal(SHIM))
    page = page.replace('<script src="data.js"></script>', payload, 1)

    # 5) 작품을 여는 주소를 '파일 안에서 꺼낸 주소' 로 바꿉니다
    old_enc = "function enc(p) { return p.split('/').map(encodeURIComponent).join('/'); }"
    if old_enc not in page:
        sys.exit("화면 코드가 바뀌었습니다 — enc() 자리를 찾지 못했습니다.")
    page = page.replace(old_enc, "function enc(p) { return OFF.url(p); }", 1)

    page = page.replace("<title>학생 작품 |", "<title>학생 작품 (보관본) |", 1)

    with io.open(OUT, "w", encoding="utf-8") as f:
        f.write(page)

    mb = os.path.getsize(OUT) / 1048576
    print("\n학생 %d명 · 파일 %d개 → 학생작품_모아보기.html  (%.1f MB)" % (len(data["students"]), total, mb))
    if dup_count:
        print("팀원이 함께 낸 똑같은 파일 %d개는 한 번만 담았습니다 (%.1f MB 절약)" % (dup_count, dup_bytes / 1048576))
    if skipped:
        print("\n[보관본에서 뺀 큰 파일] %d건 — 원본 폴더에서 직접 전달하세요" % len(skipped))
        for p in skipped:
            print("  - " + p)
    if missing:
        print("\n[빠진 파일] %d건" % len(missing))
        for p in missing:
            print("  - " + p)
    print("더블클릭하면 바로 열립니다. 공개된 곳에는 올리지 마세요.")


if __name__ == "__main__":
    main()
