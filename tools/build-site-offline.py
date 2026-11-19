# -*- coding: utf-8 -*-
"""
강의 홈페이지 전체를 한 파일로 묶기  (학생 작품 포함 · 실명 그대로 · 깃허브에 올리지 않음)

실행:  python tools/build-site-offline.py
결과:  강의홈페이지_모아보기.html    ← 이 파일 하나만 있으면 어디서든 열립니다

[무엇이 들어가나요]
  - 커리큘럼 · 회기별 강의자료 · 프로젝트 과정  (docs/index.html 그대로)
  - 슬라이드 이미지 전부 (lessons/*/slides · thumbs)
  - 실습 도구 2개 (PRD 기획 도우미 · MVP 비교·검증 도우미)
  - 학생 작품 갤러리 (tools/build-offline.py 가 만든 한 파일을 그대로 품습니다)

[어떻게 한 파일이 되나요]
  원래 홈페이지는 CSS·JS·이미지를 바깥에서 불러옵니다. 그래서 index.html 만 보내면 빈 화면이 됩니다.
  이 스크립트는 그 파일들을 전부 안에 넣고, 주소를 요청하는 자리를 '파일 안에서 꺼내 쓰기'로 바꿉니다.

[주의]
  학생 실명이 들어 있습니다. 공개된 곳(깃허브·블로그)에 올리지 마세요.
  (.gitignore 에 넣어 두어 실수로 올라가지 않습니다.)
"""
import argparse
import base64
import io
import os
import json
import re
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS = os.path.join(ROOT, "docs")
OUT = os.path.join(ROOT, "강의홈페이지_모아보기.html")
STUDENT_BUNDLE = os.path.join(ROOT, "학생작품_모아보기.html")

MIME = {"webp": "image/webp", "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "gif": "image/gif", "svg": "image/svg+xml", "pdf": "application/pdf",
        "html": "text/html", "css": "text/css", "js": "text/javascript"}

# 한 파일 안에서는 주소가 없어 localStorage(저장 기능)가 막힙니다. 대신 쓸 임시 저장소를 끼워 넣습니다.
SHIM = (
    "<script>(function(){try{window.localStorage.getItem('_');}catch(e){var m={};"
    "Object.defineProperty(window,'localStorage',{configurable:true,value:{"
    "getItem:function(k){return Object.prototype.hasOwnProperty.call(m,k)?m[k]:null;},"
    "setItem:function(k,v){m[k]=String(v);},removeItem:function(k){delete m[k];},"
    "clear:function(){m={};},key:function(i){return Object.keys(m)[i]||null;},"
    "get length(){return Object.keys(m).length;}}});}})();</script>\n"
)

HELPER = """
<script>
/* 이미지·도구·학생 갤러리를 파일 안에서 꺼내 씁니다 */
var OFFA = (function () {
  var F = window.OFFLINE_ASSETS || {}, cache = {}, miss = {};
  var SHIM = %s;
  function key(p) {
    p = String(p || '').split('#')[0].split('?')[0];
    try { p = decodeURI(p); } catch (e) {}
    return p.replace(/^\\.\\//, '').replace(/^\\//, '');
  }
  function url(p) {
    var k = key(p);
    if (cache[k]) return cache[k];
    var f = F[k];
    if (!f) { if (!miss[k]) { miss[k] = 1; console.warn('파일 안에 없는 자료:', k); } return p; }
    var blob;
    if (f.t === 't') {
      blob = new Blob([(f.m === 'text/html' ? SHIM : '') + f.d], { type: f.m + ';charset=utf-8' });
    } else {
      var bin = atob(f.d), arr = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
      blob = new Blob([arr], { type: f.m });
    }
    cache[k] = URL.createObjectURL(blob);
    return cache[k];
  }
  return { url: url, missing: function () { return Object.keys(miss); } };
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
    s = json.dumps(obj, ensure_ascii=False, separators=(",", ":"))
    return s.replace("</", "<\\/").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")


def inline_js(code):
    """<script> 안에 그대로 넣어도 끊기지 않게"""
    return code.replace("</script>", "<\\/script>")


def patch(text, old, new, what):
    if old not in text:
        sys.exit("홈페이지 코드가 바뀌었습니다 — %s 자리를 찾지 못했습니다." % what)
    return text.replace(old, new, 1)


def limit_lessons(lessons_js, keep):
    """고른 회기만 슬라이드를 넣습니다. 나머지는 '준비 중' 으로 바꿔 빈 그림이 안 뜨게 합니다."""
    for slug in re.findall(r"slidesPath: 'lessons/(lesson-\d+)/slides/'", lessons_js):
        if slug in keep:
            continue
        lessons_js = re.sub(
            r"status: 'open',(\s*)slideCount: \d+,(\s*)slidesPath: 'lessons/%s/slides/'" % slug,
            r"status: 'coming-soon',\1slideCount: 0,\2slidesPath: 'lessons/%s/slides/'" % slug,
            lessons_js)
    return lessons_js


def main():
    ap = argparse.ArgumentParser(description="강의 홈페이지를 한 파일로 묶습니다.")
    ap.add_argument("--lessons", default="", metavar="1,2",
                    help="이 회기의 슬라이드만 넣습니다 (예: --lessons 1,2). 비우면 전부")
    ap.add_argument("--no-students", action="store_true", help="학생 작품을 빼고 만듭니다 (용량이 확 줄어듭니다)")
    ap.add_argument("--out", default="", metavar="파일이름.html", help="저장할 파일 이름")
    args = ap.parse_args()

    keep = set()
    if args.lessons:
        keep = {"lesson-%02d" % int(n) for n in re.findall(r"\d+", args.lessons)}
    out_path = os.path.join(ROOT, args.out) if args.out else OUT
    with_students = not args.no_students

    # 1) 학생 작품 한 파일을 먼저 만듭니다 (이 파일을 통째로 품습니다)
    if with_students:
        print("[1/4] 학생 작품 한 파일 만들기 — tools/build-offline.py")
        if subprocess.run([sys.executable, os.path.join(ROOT, "tools", "build-offline.py")], cwd=ROOT).returncode:
            sys.exit("build-offline.py 가 실패했습니다.")
    else:
        print("[1/4] 학생 작품은 넣지 않습니다 (--no-students)")

    print("[2/4] 홈페이지 자료 담기")
    assets, total_bytes = {}, 0

    def add_bin(rel):
        nonlocal total_bytes
        fp = os.path.join(DOCS, *rel.split("/"))
        ext = rel.rsplit(".", 1)[-1].lower()
        with open(fp, "rb") as f:
            raw = f.read()
        assets[rel] = {"t": "b", "m": MIME.get(ext, "application/octet-stream"),
                       "d": base64.b64encode(raw).decode("ascii")}
        total_bytes += len(raw)

    def add_text(rel, text):
        nonlocal total_bytes
        ext = rel.rsplit(".", 1)[-1].lower()
        assets[rel] = {"t": "t", "m": MIME.get(ext, "text/plain"), "d": text}
        total_bytes += len(text.encode("utf-8"))

    # 슬라이드 · 썸네일 이미지
    img_count = 0
    lessons_dir = os.path.join(DOCS, "lessons")
    for lesson in sorted(os.listdir(lessons_dir)):
        if keep and lesson not in keep:
            continue
        for sub in ("slides", "thumbs"):
            d = os.path.join(lessons_dir, lesson, sub)
            if not os.path.isdir(d):
                continue
            for fn in sorted(os.listdir(d)):
                if fn.lower().endswith((".webp", ".png", ".jpg", ".jpeg")):
                    add_bin("lessons/%s/%s/%s" % (lesson, sub, fn))
                    img_count += 1

    # 실습 도구 2개 — 홈페이지로 돌아가는 링크는 '뒤로 가기' 로 바꿉니다
    tool_count = 0
    for rel, back in (("PRD/index.html", '../index.html#/lesson-02'), ("MVP/index.html", '../index.html#/tools')):
        fp = os.path.join(DOCS, *rel.split("/"))
        if not os.path.isfile(fp):
            print("  ! %s 가 없습니다 — 건너뜁니다" % rel)
            continue
        t = read_text(fp)
        t = t.replace('href="%s"' % back, 'href="#" onclick="history.back();return false;"')
        add_text(rel, t)
        tool_count += 1

    # 학생 작품 갤러리 (이미 한 파일로 묶여 있는 것을 그대로)
    if with_students:
        if not os.path.isfile(STUDENT_BUNDLE):
            sys.exit("학생작품_모아보기.html 이 없습니다.")
        gallery = read_text(STUDENT_BUNDLE)
        # 갤러리 로고는 단독 파일용('맨 위로')으로 만들어져 있습니다.
        # 홈페이지 안에 담길 때는 로고를 누르면 홈페이지로 돌아가게 바꿉니다.
        brand_solo = ('<a class="brand" href="#" aria-label="맨 위로" '
                      'onclick="window.scrollTo({top:0,behavior:\'smooth\'});return false;">')
        if brand_solo not in gallery:
            sys.exit("학생 갤러리의 로고 링크를 찾지 못했습니다 — tools/build-offline.py 의 BRAND_SOLO 와 맞춰 주세요.")
        gallery = gallery.replace(
            brand_solo,
            '<a class="brand" href="#" aria-label="강의 홈페이지로 돌아가기" '
            'onclick="history.back();return false;">', 1)
        add_text("students/index.html", gallery)

    print("[3/4] 홈페이지 코드 고치기")
    page = read_text(os.path.join(DOCS, "index.html"))
    css = read_text(os.path.join(DOCS, "assets", "css", "styles.css"))
    lessons_js = read_text(os.path.join(DOCS, "assets", "js", "lessons.js"))
    app_js = read_text(os.path.join(DOCS, "assets", "js", "app.js"))

    # 슬라이드 큰 이미지
    app_js = patch(app_js,
                   "function srcFor(i) { return lesson.slidesPath + pad3(i + 1) + '.webp'; }",
                   "function srcFor(i) { return OFFA.url(lesson.slidesPath + pad3(i + 1) + '.webp'); }",
                   "슬라이드 이미지")
    # 아래 썸네일
    app_js = patch(app_js,
                   "'<img src=\"' + esc(lesson.thumbsPath) + pad3(i) + '.webp\" alt=\"\" '",
                   "'<img src=\"' + esc(OFFA.url(lesson.thumbsPath + pad3(i) + '.webp')) + '\" alt=\"\" '",
                   "썸네일 이미지")
    # 실습 도구 링크
    app_js = patch(app_js,
                   "'<a class=\"btn btn-primary card-link\" href=\"' + esc(t.href) + '\">'",
                   "'<a class=\"btn btn-primary card-link\" href=\"' + esc(OFFA.url(t.href)) + '\">'",
                   "실습 도구 링크")
    # 학생 작품 메뉴 — 한 파일 안에서는 항상 보여 줍니다
    old_link = ("  var studentsLink = document.getElementById('studentsLink');\n"
                "  if (studentsLink) {\n"
                "    if (location.protocol === 'file:') studentsLink.hidden = false;\n"
                "    else fetch('students/index.html', { method: 'HEAD' })\n"
                "      .then(function (r) { if (r.ok) studentsLink.hidden = false; })\n"
                "      .catch(function () {});\n"
                "  }")
    if with_students:
        new_link = ("  var studentsLink = document.getElementById('studentsLink');\n"
                    "  if (studentsLink) {\n"
                    "    studentsLink.hidden = false;\n"
                    "    studentsLink.href = OFFA.url('students/index.html');\n"
                    "  }")
    else:
        new_link = ("  var studentsLink = document.getElementById('studentsLink');\n"
                    "  if (studentsLink) studentsLink.remove();   // 학생 작품을 빼고 만든 파일")
    app_js = patch(app_js, old_link, new_link, "학생 작품 메뉴")

    # 고른 회기만 담았으면 나머지 회기는 '준비 중' 으로 보여 줍니다
    if keep:
        lessons_js = limit_lessons(lessons_js, keep)

    # 바깥 CSS → 안으로
    page = patch(page, '<link rel="stylesheet" href="assets/css/styles.css">',
                 "<style>\n" + css + "\n</style>", "CSS 연결")
    # 바깥 JS → 자료 + 도우미 + 원래 코드 순서로 안에
    payload = ("<script>\nwindow.OFFLINE_ASSETS = " + js_literal(assets) + ";\n</script>\n"
               + HELPER % js_literal(SHIM)
               + "<script>\n" + inline_js(lessons_js) + "\n</script>\n"
               + "<script>\n" + inline_js(app_js) + "\n</script>")
    page = patch(page,
                 '<script src="assets/js/lessons.js"></script>\n<script src="assets/js/app.js"></script>',
                 payload, "JS 연결")

    page = patch(page, "<title>AI × BUILD |", "<title>AI × BUILD (보관본) |", "제목")
    if with_students:
        page = patch(page, '<p class="footer-sub">한국디지털미디어고등학교 · 웹 프로젝트 10회기</p>',
                     '<p class="footer-sub">한국디지털미디어고등학교 · 웹 프로젝트 10회기</p>\n'
                     '    <p class="footer-sub">수업용 보관본 — 학생 실명이 들어 있어 공개된 곳에 올리지 않습니다.</p>',
                     "안내 문구")

    name = os.path.basename(out_path)
    print("[4/4] 한 파일로 쓰기 — %s" % name)
    with io.open(out_path, "w", encoding="utf-8") as f:
        f.write(page)

    mb = os.path.getsize(out_path) / 1048576
    print("\n슬라이드 이미지 %d개 · 실습 도구 %d개 · 학생 작품 %s → %s (%.1f MB)"
          % (img_count, tool_count, "있음" if with_students else "없음", name, mb))
    print("원본 자료 %.1f MB 를 담았습니다. 더블클릭하면 바로 열립니다." % (total_bytes / 1048576))
    if with_students:
        print("학생 실명이 들어 있습니다 — 공개된 곳에는 올리지 마세요.")
    else:
        print("학생 실명이 없는 파일입니다 — 노션·드라이브에 올려도 됩니다.")


if __name__ == "__main__":
    main()
