# -*- coding: utf-8 -*-
"""
PRD/  ->  docs/PRD/   (배포용 복사 + '강의 홈페이지로 돌아가기' 버튼 자동 삽입)

사용법:  python tools/sync-prd.py

왜 필요한가
-----------
강의 홈페이지는 docs/ 폴더를 GitHub Pages 로 배포합니다.
최상위의 PRD/ 폴더는 배포에 포함되지 않으므로, docs/PRD/ 로 복사해야
학생이 웹에서 PRD 기획 페이지를 열 수 있습니다.

원본 보호
---------
- 최상위 PRD/ 원본은 읽기만 합니다. 절대 수정하지 않습니다.
- 돌아가기 버튼은 docs/PRD/ 의 "복사본"에만 삽입합니다.
- 여러 번 실행해도 버튼이 중복 삽입되지 않습니다.
- 원본 PRD 를 수정했다면 이 스크립트를 다시 실행하세요.
"""
import os
import re
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "03_실습도구_원본", "PRD")
DEST = os.path.join(ROOT, "docs", "PRD")

# 삽입 블록을 찾아내는 표식 (재실행 시 중복 방지 / 교체용)
MARK_START = "<!-- LECTURE-BACK-LINK:START (tools/sync-prd.py 가 자동 생성) -->"
MARK_END = "<!-- LECTURE-BACK-LINK:END -->"

BLOCK = MARK_START + """
<style>
.lecture-back-bar{
  background:#eef0fe;border-bottom:1px solid #d7daf5;
}
.lecture-back-bar .inner{
  max-width:940px;margin:0 auto;padding:9px 20px;
}
.lecture-back-link{
  display:inline-flex;align-items:center;gap:7px;
  font-family:'Pretendard','Pretendard Variable',-apple-system,BlinkMacSystemFont,
              'Segoe UI','Malgun Gothic','Apple SD Gothic Neo',system-ui,sans-serif;
  font-size:13.5px;font-weight:700;line-height:1;
  color:#3730a3;text-decoration:none;
  padding:9px 14px;border-radius:9px;border:1px solid #c7cbf2;background:#fff;
  min-height:38px;transition:background .15s,border-color .15s,color .15s;
}
.lecture-back-link:hover{background:#4f46e5;border-color:#4f46e5;color:#fff}
.lecture-back-link:focus-visible{outline:3px solid #4f46e5;outline-offset:2px}
@media (max-width:520px){
  .lecture-back-bar .inner{padding:8px 16px}
  .lecture-back-link{font-size:13px;width:100%;justify-content:center}
}
</style>
<div class="lecture-back-bar">
  <div class="inner">
    <a class="lecture-back-link" href="__HREF__">
      <span aria-hidden="true">&#8592;</span> 강의 홈페이지로 돌아가기
    </a>
  </div>
</div>
""" + MARK_END


# 배포본에 복사하지 않을 파일 (웹 실행에 필요 없거나 올리면 안 되는 것)
SKIP_EXT = (
    ".ppt", ".pptx", ".pptm", ".pot", ".potx",     # PowerPoint 원본
    ".doc", ".docx", ".xls", ".xlsx", ".hwp",      # 기타 원본 문서
    ".tmp", ".temp", ".asd", ".wbk", ".bak",       # 임시 · 자동저장 · 복구
    ".pyc", ".log", ".zip",
)
SKIP_DIRS = {"node_modules", ".git", ".vscode", ".idea", "__pycache__"}


def skip(dirpath, names):
    """shutil.copytree 에 넘길 제외 목록을 만든다."""
    out = []
    for n in names:
        low = n.lower()
        full = os.path.join(dirpath, n)
        if os.path.isdir(full):
            if n in SKIP_DIRS or n.startswith("."):
                out.append(n)
            continue
        if low.endswith(SKIP_EXT):
            out.append(n)
        elif n.startswith("~$") or n.startswith("."):   # 오피스 잠금 파일 · 숨김 파일
            out.append(n)
        elif low in ("thumbs.db", "desktop.ini", ".ds_store"):
            out.append(n)
    return out


# ---------------------------------------------------------------------
#  배포본에만 적용하는 치환 : 다운로드 확장자 .txt -> .md
#
#  학생이 받은 파일을 Claude Code 프로젝트 폴더에 PRD.md 로 바로 쓰기 위해서입니다.
#  원본 PRD/ 는 건드리지 않고, docs/PRD/ 복사본에만 적용합니다.
#  (CSS 클래스 ".txt" 같은 것을 건드리지 않도록 문장 단위로 정확히 바꿉니다)
# ---------------------------------------------------------------------
# HTML 소스 안에는 "실제 BOM 문자"가 아니라
# 백슬래시 + uFEFF 라는 "글자 6개"가 들어 있습니다. 그것과 똑같이 만듭니다.
BOM_ESCAPE = "\\uFEFF"

MD_REPLACEMENTS = [
    # 저장될 파일 이름
    ("return name + '_PRD.txt';",
     "return name + '_PRD.md';"),

    # Blob MIME 타입 + BOM 제거
    #
    #  .txt 였을 때는 Windows 메모장에서 한글이 깨지지 않도록
    #  맨 앞에 UTF-8 BOM 을 붙였습니다.
    #  그런데 .md 에서는 그 BOM 이 첫 줄의 "# 제목" 바로 앞에 붙어서
    #  Markdown 파서가 이것을 제목(h1)으로 인식하지 못합니다.
    #  (marked 파서로 실제 확인함 - BOM 이 있으면 h1 이 생성되지 않음)
    #
    #  .md 는 VS Code · Claude Code · GitHub 에서 여는 파일이고
    #  이들 모두 BOM 없는 UTF-8 을 정상 인식하므로 BOM 을 뺍니다.
    #  Windows 10 1903 이상의 메모장도 BOM 없는 UTF-8 을 올바로 읽습니다.
    ("return new Blob(['" + BOM_ESCAPE + "' + text], { type: 'text/plain;charset=utf-8' });",
     "return new Blob([text], { type: 'text/markdown;charset=utf-8' });"),

    # 저장 대화상자에서 고를 수 있는 형식
    ("types: [{ description: '텍스트 파일', accept: { 'text/plain': ['.txt'] } }]",
     "types: [{ description: 'Markdown 파일', accept: { 'text/markdown': ['.md'] } }]"),

    # 버튼 문구
    (">PRD.txt 다운로드<", ">PRD.md 다운로드<"),

    # 안내 메시지
    ("toast('PRD.txt 파일을 저장했습니다.', 'ok');",
     "toast('PRD.md 파일을 저장했습니다.', 'ok');"),
    ("toast('PRD.txt 파일을 내려받았습니다.', 'ok');",
     "toast('PRD.md 파일을 내려받았습니다.', 'ok');"),

    # 주석
    ("/* 일반 텍스트 파일로 만든다.",
     "/* Markdown 파일로 만든다."),
]


def to_markdown(html, rel):
    """학생용 PRD 페이지의 다운로드 확장자를 .md 로 바꾼다."""
    n = 0
    for old, new in MD_REPLACEMENTS:
        if old in html:
            html = html.replace(old, new)
            n += 1
    return html, n


def back_href(rel_path):
    """docs/PRD 기준 깊이에 맞춰 docs/index.html 로 가는 상대 경로를 만든다."""
    depth = len(rel_path.replace("\\", "/").split("/")) - 1   # 파일명 제외한 폴더 깊이
    return "../" * (depth + 1) + "index.html#/lesson-02"


def inject(html, href):
    # 이미 삽입된 블록이 있으면 통째로 걷어내고 새로 넣는다 (중복 방지)
    html = re.sub(re.escape(MARK_START) + r".*?" + re.escape(MARK_END),
                  "", html, flags=re.S)

    block = BLOCK.replace('__HREF__', href)
    m = re.search(r"<body[^>]*>", html, flags=re.I)
    if not m:
        return html, False
    return html[:m.end()] + "\n" + block + html[m.end():], True


def main():
    if not os.path.isdir(SRC):
        print("[오류] PRD 폴더가 없습니다: %s" % SRC)
        sys.exit(1)

    if os.path.isdir(DEST):
        shutil.rmtree(DEST)
    shutil.copytree(SRC, DEST, ignore=skip)

    n_html = 0
    n_other = 0
    for dirpath, _dirnames, filenames in os.walk(DEST):
        for fn in filenames:
            full = os.path.join(dirpath, fn)
            rel = os.path.relpath(full, DEST)
            if fn.lower().endswith((".html", ".htm")):
                with open(full, encoding="utf-8") as f:
                    html = f.read()
                html, n_md = to_markdown(html, rel)
                out, ok = inject(html, back_href(rel))
                if ok:
                    with open(full, "w", encoding="utf-8", newline="") as f:
                        f.write(out)
                    print("  버튼 삽입  docs/PRD/%s   ->  %s"
                          % (rel.replace("\\", "/"), back_href(rel)))
                    if n_md:
                        print("  .md 변환   docs/PRD/%s   (%d곳)"
                              % (rel.replace("\\", "/"), n_md))
                    n_html += 1
                else:
                    print("  [건너뜀] <body> 를 찾지 못함: %s" % rel)
            else:
                n_other += 1

    print("\n완료 — HTML %d개에 '강의 홈페이지로 돌아가기' 버튼 삽입, 기타 정적 파일 %d개 복사"
          % (n_html, n_other))
    print("PPT · 임시 · 자동저장 · 캐시 파일은 복사하지 않았습니다.")
    print("원본 PRD/ 폴더는 수정하지 않았습니다.")


if __name__ == "__main__":
    main()
