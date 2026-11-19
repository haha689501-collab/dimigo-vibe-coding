# AI × BUILD | Claude Vibe Coding

한국디지털미디어고등학교 · 웹 프로젝트 10회기 강의자료 웹사이트

학생이 메인 화면에서 전체 10회기 커리큘럼을 보고, 원하는 회기를 클릭하면
그 회기의 PPT 슬라이드를 웹에서 바로 넘겨 볼 수 있습니다.

---

## 1. 빠르게 실행해 보기

이 사이트는 **빌드(build)가 필요 없는 정적 웹사이트**입니다.
설치할 것도, 컴파일할 것도 없습니다.

### 방법 A — 가장 간단 (권장)

`docs` 폴더 안의 `index.html` 파일을 **더블클릭**하면 브라우저에서 열립니다.

> 단, 일부 브라우저는 로컬 파일(`file://`)에서 이미지 로딩을 제한할 수 있습니다.
> 화면이 이상하면 아래 방법 B를 쓰세요.

### 방법 B — 로컬 서버로 열기 (실제 배포와 동일한 환경)

터미널(PowerShell)에서 프로젝트 폴더로 이동한 뒤:

```powershell
cd docs
npx http-server -p 8765 -c-1
```

그리고 브라우저에서 <http://127.0.0.1:8765> 로 접속합니다.

Python이 설치되어 있다면 이 방법도 됩니다:

```powershell
cd docs
python -m http.server 8765
```

멈출 때는 터미널에서 `Ctrl + C` 를 누릅니다.

---

## 2. 폴더 구조

> **GitHub에 올라가는 것은 `docs/` 하나뿐입니다.** (그중 `docs/students/` 학생 작품은 실명이 있어 제외)
> 나머지 폴더는 모두 이 컴퓨터에만 있습니다. 옮긴 기록은 `정리기록.md`에 있습니다.

```text
claude_/
├── 강의용홈페이지.html            ← 더블클릭하면 docs/index.html 이 열리는 바로가기
├── README.md · 정리기록.md
│
├── docs/                          ← ★ GitHub에 올라가는 강의 홈페이지 (GitHub Pages)
│   ├── index.html
│   ├── assets/js/lessons.js       ←   회기 정보는 전부 여기 한 파일에만 있습니다
│   ├── lessons/lesson-0X/         ←   슬라이드 이미지(webp) · 썸네일 · slides.pdf
│   ├── PRD/ · MVP/                ←   실습 도구 배포본 (03_실습도구_원본 에서 복사)
│   ├── demo/                      ←   수업 시연 작품
│   └── students/                  ←   학생 작품 (로컬 전용 · GitHub에 안 올라감)
│
├── tools/                         ← 변환·복사·빌드 스크립트 (아래 표 참고)
│
├── 01_수업PPT/                    ← 최종본 PPT (export-slides.ps1 이 여기서 슬라이드를 만듭니다)
│   └── 지난버전/                  ←   이전 버전 · 초안 · restyle 원본
├── 02_교안·스크립트/              ← 교안 · 발표 스크립트 · 커리큘럼 문서 (.md)
├── 03_실습도구_원본/              ← PRD 기획 도우미 · MVP 비교검증 도우미 · 용어사전 원본
├── 04_수업자료·시연/              ← 우주탈출 시연 · 프로젝트 주제카드 · 7-8차시 쉼터·미션카드
├── 05_모션그래픽/                 ← 모션그래픽 프롬프트 · 샘플 · 깃허브용 한 파일
├── 06_참고·조사/                  ← 커리큘럼 PDF · 만족도 조사 · 유튜브 참고 · 노션 보관본
├── _보관/                         ← 예전 버전 · 복사본
└── _삭제후보/                     ← 지워도 되는 것 (확인 후 직접 삭제)
```

| 스크립트 | 하는 일 |
|---|---|
| `tools/export-slides.ps1` | `01_수업PPT/` → PNG (+PDF) |
| `tools/make-webp.py` | PNG → `docs/lessons/` WebP |
| `tools/sync-prd.py` | `03_실습도구_원본/PRD/` → `docs/PRD/` 복사 |
| `tools/sync-students.py` | 학생 폴더 → `docs/students/data.js` 작품 목록 |
| `tools/build-offline.py` · `build-site-offline.py` | 나눠 주기용 한 파일 보관본 |

> PPT 생성 스크립트(`tools/make-lesson*.py`)는 새 PPT를 **최상위**에 만듭니다. 확인 후 `01_수업PPT/`로 옮기세요.

---

## 3. 왜 이 기술을 선택했나

| 선택 | 이유 |
|---|---|
| **HTML + CSS + 순수 JavaScript** (React·빌드 도구 없음) | 빌드 단계가 없어 파일을 고치고 바로 새로고침하면 끝입니다. 수업 자료는 자주 추가·수정되므로, 선생님이 `npm install` 없이도 유지보수할 수 있어야 합니다. |
| **해시 라우팅** (`#/lesson-01`) | GitHub Pages는 서버 설정을 바꿀 수 없습니다. 일반 경로(`/lesson-01`)를 쓰면 새로고침할 때 404가 납니다. 해시 방식은 **새로고침·뒤로가기·주소 직접 입력이 모두 정상 동작**합니다. |
| **모든 경로를 상대 경로로** | `github.io/저장소이름/` 처럼 **하위 경로**에 배포해도 이미지와 링크가 깨지지 않습니다. |
| **슬라이드를 WebP 이미지로 변환** | 브라우저에서 PPTX를 직접 해석하는 방식은 글꼴이 깨지고 불안정합니다. 이미지로 미리 바꿔 두면 어떤 기기에서도 PPT 원본 그대로 보입니다. (슬라이드 78장 전체가 4.7MB 로 가볍습니다) |
| **썸네일은 지연 로딩(`loading="lazy"`)** | 처음 들어올 때 이미지 78장을 한꺼번에 받지 않습니다. 보고 있는 슬라이드의 앞·뒤 1장만 미리 받아 둡니다. |

로그인·데이터베이스·외부 유료 API·API 키를 전혀 사용하지 않습니다.

---

## 4. PPT를 슬라이드 이미지와 PDF로 바꾸는 방법

### 이미 변환은 끝나 있습니다

1~3회기 슬라이드 **78장**과 교안 **PDF 3개**는 이미 `docs/lessons/` 안에 들어 있습니다.
아래 내용은 **4회기 이후 새 PPT를 추가할 때** 필요합니다.

### 변환 순서 (2단계)

```text
                                ┌─▶ PNG 1920x1080 ──(2단계: Python)──▶ WebP 1600x900 + 썸네일
PPTX ──(1단계: PowerPoint)──────┤
                                └─▶ PDF (학생 다운로드용 교안)
```

#### 1단계 — PPTX를 PNG + PDF로

> PowerPoint가 켜져 있어도 실행할 수 있습니다.
> **이미 열려 있던 PowerPoint 창은 건드리지 않고, 작업이 끝나도 종료하지 않습니다.**
> (스크립트가 직접 켠 경우에만 종료합니다)

```powershell
powershell -ExecutionPolicy Bypass -File tools\export-slides.ps1
```

이 스크립트는 안전하게 동작합니다.

- 원본 PPTX를 **직접 열지 않습니다.** `tools\_work\` 로 복사한 뒤 **복사본만** 엽니다.
- 복사본은 창 없이(비표시) 읽기 전용으로 열고 바로 닫습니다.
- 파일명에 한글이 있어도 되고, 비슷한 이름이 여러 개면 **가장 최근에 수정된 파일**을 씁니다.
- 결과는 `tools\_png\lesson-0X\` 에 PNG로, `tools\_png\lesson-0X.pdf` 에 PDF로 저장됩니다.
- PDF에는 **강사용 발표자 노트가 포함되지 않습니다.** (슬라이드만 출력)

#### 2단계 — PNG를 WebP로

```powershell
python tools\make-webp.py
```

`docs/lessons/lesson-0X/slides/001.webp ...` 와 `thumbs/001.webp ...` 가 만들어지고,
**슬라이드가 몇 장인지** 화면에 표시됩니다. 이 숫자를 5장에서 사용합니다.

Pillow가 없다면 먼저 설치하세요:

```powershell
pip install pillow
```

### PowerPoint가 없는 컴퓨터라면 — LibreOffice 사용

무료 프로그램인 LibreOffice로도 같은 작업을 할 수 있습니다.

1. <https://ko.libreoffice.org/download/> 에서 내려받아 설치합니다.
   (또는 `winget install TheDocumentFoundation.LibreOffice`)
2. PPTX를 PDF로 바꿉니다.

   ```powershell
   & "C:\Program Files\LibreOffice\program\soffice.exe" --headless `
     --convert-to pdf --outdir tools\_png "회기별_PPT파일.pptx"
   ```

3. PDF의 각 쪽을 PNG로 바꿉니다. (Poppler 설치 필요: `winget install oschwartz10612.Poppler`)

   ```powershell
   pdftoppm -png -r 144 tools\_png\파일이름.pdf tools\_png\lesson-04\슬라이드
   ```

4. 그다음 위의 **2단계**(`python tools\make-webp.py`)를 그대로 실행합니다.

### 도구가 하나도 없다면 — 손으로 저장하기

PowerPoint에서 **파일 → 내보내기 → 파일 형식 변경 → PNG** 를 선택하고
"모든 슬라이드"를 저장하면 됩니다. 그다음 2단계를 실행하세요.

> 슬라이드 이미지는 반드시 **16:9 비율**을 유지해야 화면에서 잘리거나 찌그러지지 않습니다.

---

## 5. 4~10회기 자료를 추가하는 방법 ★

새 회기를 공개하려면 **파일을 넣고, 데이터 파일 한 곳에서 숫자 두 개만 바꾸면** 됩니다.
코드를 새로 짤 필요가 없습니다.

아래는 **4회기 PPT가 완성되었다고 가정한** 예시입니다.

### ① PPT 파일을 프로젝트 폴더에 넣기

완성한 PPT를 프로젝트 맨 위 폴더에 넣습니다.
파일 이름에 `_4` 가 들어가게 하세요. (변환 스크립트가 이 패턴으로 찾습니다)

```text
claude_/디미고_바이브코딩_4차시_클로드코드입문.pptx
```

`tools\export-slides.ps1` 의 `$map` 에 4회기 줄을 추가합니다.

```powershell
$map = @(
  @{ id = 1; pattern = '*_1*.pptx' },
  @{ id = 2; pattern = '*_2*.pptx' },
  @{ id = 3; pattern = '*_3*.pptx' },
  @{ id = 4; pattern = '*_4*.pptx' }      # ← 이 줄 추가
)
```

### ② 슬라이드 이미지와 PDF 만들기

아래 두 줄을 차례로 실행합니다.

```powershell
powershell -ExecutionPolicy Bypass -File tools\export-slides.ps1
python tools\make-webp.py
```

실행이 끝나면 아래처럼 슬라이드 장수가 출력됩니다. **이 숫자를 기억하세요.**

```text
lesson-04 -> slideCount: 18
```

이미지는 자동으로 여기에 들어갑니다.

```text
docs/lessons/lesson-04/slides/001.webp ~ 018.webp
docs/lessons/lesson-04/thumbs/001.webp ~ 018.webp
```

### ③ 다운로드용 PDF 넣기

학생이 교안을 내려받을 수 있게 PDF를 복사합니다.
(②단계에서 `tools\_png\lesson-04.pdf` 가 이미 만들어져 있습니다)

```powershell
Copy-Item "tools\_png\lesson-04.pdf" "docs\lessons\lesson-04\slides.pdf"
```

> 교안은 **PDF로만** 제공합니다. PPT 원본 파일은 웹사이트에 올리지 않습니다.

### ④ 데이터 파일 한 곳만 수정 — `docs/assets/js/lessons.js`

메모장이나 VS Code로 열어서 `id: 4` 부분을 찾습니다.
**바꿀 곳은 주석(`← ①②③`)이 붙은 네 줄뿐입니다.**

```javascript
  {
    id: 4,
    slug: 'lesson-04',
    phase: 'BUILD',
    title: 'Claude Code 입문',
    shortTitle: 'Claude Code 입문',
    summary: 'AI가 브라우저 밖으로 나와 내 폴더를 직접 엽니다.',
    keywords: ['VS Code', '프로젝트 폴더', 'AI Agent'],
    goals: [ '...' ],
    outcome: 'Claude Code 개발환경 구성',

    status: 'open',            // ← ① 'coming-soon' 을 'open' 으로 (준비 중 → 자료 공개)
    slideCount: 18,            // ← ② 위 ②단계에서 나온 숫자를 입력

    slidesPath: 'lessons/lesson-04/slides/',
    thumbsPath: 'lessons/lesson-04/thumbs/',
    pdfFile: 'lessons/lesson-04/slides.pdf',        // ← ③ null 이었다면 이 경로로 바꾸기
    pdfName: '디미고_바이브코딩_4회기.pdf',           //     내려받을 때 저장될 파일 이름
  },
```

정리하면 이 네 가지입니다.

| 바꿀 항목 | 바꾸기 전 | 바꾼 후 |
|---|---|---|
| `status` | `'coming-soon'` (준비 중) | `'open'` (자료 공개) |
| `slideCount` | `0` | `18` (실제 슬라이드 장수) |
| `pdfFile` | `null` | `'lessons/lesson-04/slides.pdf'` |
| `pdfName` | `null` | `'디미고_바이브코딩_4회기.pdf'` |

> 제목이나 키워드도 실제 PPT에 맞게 자유롭게 고치면 됩니다.
> 이 파일만 고치면 **메인 카드 · 커리큘럼 표 · 뷰어 · 이전/다음 회기 버튼**이 전부 함께 바뀝니다.

### ⑤ 저장하고 확인

파일을 저장한 뒤 브라우저를 **새로고침**하면 끝입니다.
4회기 카드가 `준비 중` 에서 `자료 공개` 로 바뀌고, 클릭하면 슬라이드 뷰어가 열립니다.

> `pdfFile` 이 `null` 이면 다운로드 버튼은 아예 표시되지 않습니다.
> (작동하지 않는 버튼을 만들지 않기 위해서입니다)

---

## 6. GitHub에 올리고 GitHub Pages로 공개하기

### ① 처음 한 번만 — 저장소 만들기

1. <https://github.com> 에 로그인하고 오른쪽 위 **+ → New repository** 를 누릅니다.
2. 저장소 이름을 정합니다. 예: `dimigo-vibe-coding`
3. **Public** 을 선택하고 **Create repository** 를 누릅니다.
   (GitHub Pages 무료 공개는 Public 저장소에서 됩니다)

### ② 내 컴퓨터에서 업로드

프로젝트 폴더에서 터미널을 열고 아래를 차례로 실행합니다.
`내아이디` 와 `저장소이름` 은 본인 것으로 바꾸세요.

```powershell
git init
git add .
git commit -m "디미고 바이브코딩 강의자료 웹사이트"
git branch -M main
git remote add origin https://github.com/내아이디/저장소이름.git
git push -u origin main
```

### ②-1 GitHub 에 올라가는 것 / 올라가지 않는 것

`.gitignore` 에 이미 설정되어 있습니다. 따로 손댈 필요는 없습니다.

| | 대상 |
|---|---|
| ✅ **올라감** | `docs/` 폴더만 (홈페이지 · 슬라이드 이미지 · 교안 PDF · `docs/PRD/` · `docs/MVP/`). 단 `docs/students/` 학생 작품은 제외 |
| 🚫 **안 올라감** | **PPT 원본 전부** (`*.ppt`, `*.pptx`), **원본 PDF** (`*.pdf` — 최상위 커리큘럼 PDF 등), 오피스 임시·자동저장·복구 파일 (`~$*`, `*.tmp`, `*.asd`, `*.wbk`), 중간 변환물 `tools/_png/`, `node_modules/`, `.vscode/`, `__pycache__/`, `Thumbs.db`, `.DS_Store` |

> **수업 원본 PPT 와 원본 PDF 는 GitHub 에 올라가지 않습니다.**
> 학생에게는 `docs/lessons/*/slides/` 의 이미지와 `slides.pdf` 로만 제공됩니다.
>
> `*.pdf` 를 전부 제외하면 학생용 교안 PDF 까지 빠지므로,
> `.gitignore` 에 아래 한 줄로 교안 PDF 만 되살려 두었습니다.
> 이 줄을 지우면 강의자료 페이지의 "PDF 다운로드" 버튼이 404 가 됩니다.
>
> ```gitignore
> !docs/lessons/*/slides.pdf
> ```

### ③ GitHub Pages 켜기

1. GitHub 저장소 페이지에서 **Settings** 탭을 누릅니다.
2. 왼쪽 메뉴에서 **Pages** 를 누릅니다.
3. **Source** 를 `Deploy from a branch` 로 둡니다.
4. **Branch** 를 `main` 으로, 폴더를 **`/docs`** 로 선택하고 **Save** 를 누릅니다.

> 이 프로젝트가 웹사이트를 `docs/` 폴더에 만든 이유가 바로 이것입니다.
> GitHub Pages는 저장소 최상위(`/`) 또는 `/docs` 만 공개할 수 있습니다.

### ④ 1~3분 기다리면 공개됩니다

주소는 다음과 같습니다.

```text
https://내아이디.github.io/저장소이름/
```

Settings → Pages 화면 위쪽에도 완성된 주소가 표시됩니다.
초록색 체크 표시가 뜨면 배포가 끝난 것입니다.

### ⑤ 이후 자료를 추가했을 때 (반영 과정)

새 회기를 추가하거나 PRD 원본을 고친 뒤에는 이 세 줄만 실행하면 됩니다.
(PRD 를 고쳤다면 먼저 `python tools\sync-prd.py` 를 실행하세요)

```powershell
git add .
git commit -m "4회기 강의자료 추가"
git push
```

`push` 하면 GitHub가 자동으로 다시 배포합니다.
**Actions** 탭에서 진행 상황을 볼 수 있고, 보통 **1~3분** 뒤 사이트에 반영됩니다.
화면이 그대로면 브라우저에서 `Ctrl + F5` 로 강력 새로고침을 하세요.

---

## 6-2. PRD 기획 페이지 (2회기 실습 도구)

2회기 강의자료 페이지 오른쪽에 **`PRD 기획 시작하기`** 버튼이 있습니다.
누르면 학생이 단계별로 자기 PRD를 작성하는 웹페이지가 열립니다.

| 항목 | 내용 |
|---|---|
| 원본 위치 | `03_실습도구_원본/PRD/index.html` (학생용) · `…/PRD/reviewer/index.html` (교사용) |
| 배포 위치 | `docs/PRD/index.html` · `docs/PRD/reviewer/index.html` |
| 연결 경로 | `PRD/index.html` (강의 홈페이지 기준 상대 경로) |
| 돌아가기 | PRD 페이지 맨 위 `← 강의 홈페이지로 돌아가기` → 2회기 강의자료로 복귀 |
| 다운로드 | `프로젝트명_PRD.md` (Markdown · BOM 없는 UTF-8) |

### 왜 복사본이 필요한가

GitHub Pages 는 `docs/` 폴더만 공개합니다.
`03_실습도구_원본/PRD/` 폴더는 배포에 포함되지 않으므로, `docs/PRD/` 로 복사해야
학생이 인터넷에서 열 수 있습니다.

**원본 `03_실습도구_원본/PRD/` 폴더는 그대로 두고, 복사본만 만듭니다.**

### PRD 원본을 수정했다면

원본을 고친 뒤에는 아래 한 줄을 실행해 배포본에 반영하세요.

```powershell
python tools\sync-prd.py
```

이 스크립트는

- `03_실습도구_원본/PRD/` 를 `docs/PRD/` 로 복사합니다. (PPT·임시·캐시 파일은 제외)
- 각 HTML 맨 위에 `← 강의 홈페이지로 돌아가기` 버튼을 자동으로 넣습니다.
- 다운로드 확장자를 `.txt` → **`.md`** 로 바꿉니다.
- 여러 번 실행해도 버튼이 중복으로 생기지 않습니다.
- **원본 `03_실습도구_원본/PRD/` 폴더는 절대 수정하지 않습니다.**

> **왜 배포본에서만 `.md` 로 바꾸나요?**
> 원본 PRD 페이지는 `프로젝트명_PRD.txt` 로 저장하도록 만들어져 있습니다.
> 학생이 받은 파일을 Claude Code 프로젝트에 `PRD.md` 로 바로 쓰려면 `.md` 가 편합니다.
> 원본을 고치지 않고 배포본에만 적용하기 위해, 이 스크립트가 복사하면서 바꿔 줍니다.
>
> 이때 파일 맨 앞의 **UTF-8 BOM 도 함께 제거**합니다.
> BOM 이 남아 있으면 첫 줄 `# 제목` 이 Markdown 제목으로 인식되지 않기 때문입니다.

### 교사용 분석 페이지

학생들이 낸 PRD 파일을 모아서 볼 수 있는 페이지도 함께 배포됩니다.
강의 홈페이지에는 링크를 넣지 않았으므로, 필요할 때 주소로 직접 여세요.

```text
https://내아이디.github.io/저장소이름/PRD/reviewer/
```

파일을 드래그앤드롭하거나 선택해서 올릴 수 있고, 여러 번 나눠 올려도
먼저 올린 학생 자료가 그대로 유지됩니다.

---

## 6-3. 학생 작품 갤러리 (완성작만)

학생 작품은 **실명이 있어 GitHub에 올리지 않습니다.** 내 컴퓨터에서 `docs/students/index.html`을 열어 봅니다.
강의 홈페이지를 로컬에서 열면 위쪽 메뉴에 **학생 작품**이 보입니다.

### 완성작 넣는 순서

1. 학생 폴더의 알맞은 칸에 파일을 넣습니다. 모든 학생이 같은 다섯 칸을 갖습니다.

   ```text
   docs/students/
   └── 홍길동/
       ├── 01_기획서/           ← PRD
       ├── 02_웹서비스/         ← 익명게시판 완성
       ├── 03_인터랙티브웹/
       ├── 04_개인프로젝트/
       └── 05_팀프로젝트/
   ```

   - 작품 설명은 같은 폴더의 `설명.txt`에 `파일이름 : 설명` 형식으로 한 줄씩 적습니다. (없어도 됩니다)
   - 칸을 추가하려면 `06_이름` 폴더를 만들고 `tools/sync-students.py` 맨 위 `DAYS` · `LABELS`에 추가합니다.
   - 작품에 그림 폴더(`assets/`)가 있으면 폴더째 넣으세요.
2. `python tools/sync-students.py` 를 실행합니다. → `docs/students/data.js`가 새로 만들어집니다.
   - 빈 칸은 갤러리에 나오지 않습니다.
   - `[개인정보·보안 경고]`가 나오면 그 파일을 열어 확인합니다.
3. `docs/students/index.html`을 새로고침해서 확인합니다.
4. 나눠 줄 때는 `python tools/build-offline.py` → `학생작품_모아보기.html` 한 파일을 씁니다.

`docs/students/00_예시학생/`은 폴더 규칙 안내용 예시이며 갤러리에는 나오지 않습니다.

---

## 7. 웹사이트에서 할 수 있는 것

### 메인 화면

- 10회기 카드 (PC 3열 · 태블릿 2열 · 모바일 1열)
- 회기별 핵심 키워드와 결과물, 자료 상태(`자료 공개` / `준비 중`)
- 전체 커리큘럼 표 · 프로젝트 과정(IDEA → SPEC → BUILD → TEST → RELEASE)

### 슬라이드 뷰어

| 기능 | 사용법 |
|---|---|
| 이전 / 다음 슬라이드 | 슬라이드 좌우 화살표, 아래쪽 `이전` `다음` 버튼 |
| 키보드 이동 | `←` `→` (맨 앞·맨 뒤로는 `Home` `End`) |
| 모바일 | 슬라이드를 좌우로 밀기(스와이프) |
| 슬라이드 번호 | 아래쪽 가운데 `3 / 22` 표시 |
| 썸네일 이동 | 아래 썸네일 목록을 클릭 |
| 전체화면 | `전체화면` 버튼 (다시 누르면 종료) |
| 자료 내려받기 | 맨 아래 `강의자료 PDF 다운로드` |

- 첫 슬라이드에서는 `이전`, 마지막 슬라이드에서는 `다음` 버튼이 자동으로 비활성화됩니다.
- 강사용 **발표자 노트는 학생 화면에 표시되지 않습니다.** (슬라이드 이미지만 사용)

---

## 8. 자주 생기는 문제

| 증상 | 원인과 해결 |
|---|---|
| 슬라이드가 안 보이고 빈 칸만 나옴 | `lessons.js` 의 `slideCount` 숫자가 실제 이미지 장수와 다릅니다. `docs/lessons/lesson-0X/slides/` 안의 파일 개수를 세어 맞추세요. |
| 이미지 파일 이름 문제 | 반드시 `001.webp`, `002.webp` 처럼 **세 자리 숫자**여야 합니다. `1.webp` 는 인식되지 않습니다. |
| 카드가 여전히 `준비 중` | `status` 를 `'open'` 으로 바꿨는지, `slideCount` 가 `0` 이 아닌지 확인하세요. 둘 다 만족해야 공개됩니다. |
| PDF 다운로드 버튼이 안 보임 | `lessons.js` 의 `pdfFile` 이 `null` 입니다. PDF 경로를 넣어 주세요. |
| GitHub Pages에서 이미지가 안 뜸 | Settings → Pages 에서 폴더가 `/docs` 로 되어 있는지 확인하세요. |
| 수정했는데 사이트가 그대로 | 배포에 1~3분 걸립니다. `Ctrl + F5` 로 새로고침하세요. |

---

## 9. 확인한 커리큘럼 (폴더 자료 기준)

`claude 차시정리 .pptx` 의 커리큘럼 표와, 1~3회기 PPT 실제 내용을 반영했습니다.

| 회기 | 주제 | 결과물 | 자료 |
|---|---|---|---|
| 1 | 바이브코딩이란? 그리고 프롬프트 | 할 일 관리 앱 개조 실습 | 공개 (22장) |
| 2 | PRD · 아이디어를 설계도로 | 내 프로젝트 PRD.md | 공개 (31장) |
| 3 | 코드를 내 파일로 · PRD를 첫 화면으로 | 내 폴더 / index.html | 공개 (25장) |
| 4 | Claude Code 입문 | Claude Code 개발환경 구성 | 준비 중 |
| 5 | AI와 UI 만들기 | 디자인이 적용된 웹페이지 | 준비 중 |
| 6 | 동작하는 웹앱 | 실제로 동작하는 핵심 기능 | 준비 중 |
| 7 | 게임 만들기 | 동작하는 미니게임 | 준비 중 |
| 8 | 도구 만들기와 디버깅 | 생활 도구 웹앱과 오류 수정 기록 | 준비 중 |
| 9 | 나만의 프로젝트 | 내 프로젝트 V1 | 준비 중 |
| 10 | 프로젝트 완성과 발표 | 배포된 웹사이트와 최종 발표 | 준비 중 |

> **4회기 이후 제목에 대한 참고**
> `claude 차시정리 .pptx` 의 표에서는 3회기가 "Claude Code 입문"이었습니다.
> 그런데 실제 3회기 PPT(가장 최근 자료)는 "코드를 내 파일로"이고,
> 마지막 장에서 **"NEXT · 4차시 — VS Code, Claude Code"** 라고 예고하고 있습니다.
> 그래서 4회기부터 한 칸씩 밀어 반영했습니다.
> 4~10회기는 아직 `준비 중` 이므로, 실제 PPT가 나오면 `lessons.js` 에서 제목만 고치면 됩니다.

---

## 10. 원본 자료 보호에 대하여

이 프로젝트는 다음 규칙을 지키도록 만들어져 있습니다.

- 프로젝트 폴더 안의 원본 PPTX·PDF는 **읽기 전용**으로만 사용합니다.
- 변환 스크립트는 원본을 열지 않고 **복사본만** 엽니다.
- PowerPoint가 실행 중이어도 **열려 있는 창은 건드리지 않고 종료하지도 않습니다.**
- 웹사이트용 파일은 전부 `docs/` 라는 **새 폴더**에만 만듭니다.
- 프로젝트 폴더 밖의 파일이나 자동 복구 파일은 사용하지 않습니다.
