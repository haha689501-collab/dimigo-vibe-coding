/* =====================================================================
 *  회기 데이터 — 이 파일 하나만 수정하면 사이트 전체가 바뀝니다.
 *
 *  [새 회기를 공개하는 방법]  자세한 설명은 README.md 참고
 *   1. docs/lessons/lesson-0X/slides/ 에 001.webp, 002.webp ... 넣기
 *   2. docs/lessons/lesson-0X/thumbs/ 에 같은 이름의 썸네일 넣기
 *   3. 아래에서 slideCount 를 실제 장수로 바꾸기
 *   4. status 를 "coming-soon" -> "open" 으로 바꾸기
 *   5. 저장 후 GitHub 에 올리기
 *
 *  강의자료는 보기만 됩니다. PDF 는 docs/ 에 넣지 마세요 (GitHub Pages 에 그대로 공개됩니다).
 *  pdfFile / pdfName 은 null 로 둡니다.
 *
 *  status : "open"        = 자료 공개 (클릭하면 뷰어로 이동)
 *           "coming-soon" = 준비 중  (클릭 불가)
 * ===================================================================== */

const COURSE = {
  school: '한국디지털미디어고등학교',
  className: 'AI 노코드 창작 클래스',
  title: 'AI × BUILD',
  subtitle: 'Claude Vibe Coding',
  heroTitle: '아이디어를, 실제로 작동하는 웹 작품으로',
  heroSub: 'Claude Max로 나만의 작동하는 결과물 만들기',
  tagline: '한국디지털미디어고등학교 · 웹 프로젝트 10회기',
  /* 첫 화면 설명 — 줄 단위로 끊어 보여 줍니다 */
  leadLines: [
    'AI와 함께 만들고, 실행하고, 수정하며 개인 작품을 완성한 뒤',
    '친구들과 강점을 모아 완성도 높은 조별 대표작으로 발전시킵니다.',
  ],
  note: '10회기 프로젝트 · 개인 작품 + 조별 대표작',
  totalSessions: 10,
};

/* 수업 전체 흐름 : IDEA → SPEC → BUILD → TEST → RELEASE */
const PHASES = {
  IDEA: { label: 'IDEA', desc: '바이브코딩과 프롬프트의 기초' },
  SPEC: { label: 'SPEC', desc: 'PRD와 MVP로 판단 기준 세우기' },
  BUILD: { label: 'BUILD', desc: 'Claude Code로 개인 작품을 만들어 배포하고, 조 작품으로 한 바퀴 완성' },
  TEST: { label: 'TEST', desc: '사용자 검증과 UI 완성' },
  RELEASE: { label: 'RELEASE', desc: '안정화 · 배포 · 작품 쇼케이스' },
};

const LESSONS = [
  {
    id: 1,
    slug: 'lesson-01',
    phase: 'IDEA',
    title: '바이브코딩과 프롬프트',
    shortTitle: '바이브코딩과 프롬프트',
    summary: 'AI와 함께 만드는 개발 방식을 이해하고, 결과를 판단할 수 있는 프롬프트를 씁니다.',
    keywords: ['Vibe Coding', '프롬프트 5W1H', 'Claude Artifacts'],
    goals: [
      '바이브코딩이 기존 개발과 무엇이 같고 다른지 설명할 수 있다.',
      '명확성·구체성·맥락·제약을 넣어 프롬프트를 작성할 수 있다.',
      'Claude Artifacts로 웹앱을 즉시 만들고 수정할 수 있다.',
    ],
    outcome: '할 일 관리 앱 개조 실습 (LEVEL 1~4)',
    status: 'open',
    slideCount: 18,
    slidesPath: 'lessons/lesson-01/slides/',
    thumbsPath: 'lessons/lesson-01/thumbs/',
    pdfFile: null,
    pdfName: null,
  },
  {
    id: 2,
    slug: 'lesson-02',
    phase: 'SPEC',
    title: '아이디어를 설계도로 (PRD)',
    shortTitle: '아이디어와 PRD',
    summary: '컨텍스트와 MVP를 이해하고, 내가 내린 결정을 PRD.md 한 장으로 정리합니다.',
    keywords: ['컨텍스트', 'MVP · MUST/LATER/NON-GOAL', 'PRD', 'Markdown'],
    goals: [
      'AI가 무엇을 참고해 답하는지(컨텍스트) 설명할 수 있다.',
      '컨텍스트가 유실되면 왜 결과가 흔들리는지 이해한다.',
      '제품의 목적과 첫 버전의 범위(MVP)를 결정할 수 있다.',
      '결정을 PRD.md로 구조화해 전달할 수 있다.',
    ],
    outcome: '내 프로젝트 PRD.md',
    status: 'open',
    slideCount: 29,
    slidesPath: 'lessons/lesson-02/slides/',
    thumbsPath: 'lessons/lesson-02/thumbs/',
    pdfFile: null,
    pdfName: null,
  },
  {
    id: 3,
    slug: 'lesson-03',
    phase: 'SPEC',
    title: '같은 설계도, 다른 작품',
    shortTitle: '같은 설계도, 다른 작품',
    summary: '같은 PRD로 만든 두 결과를 비교해 AI의 실수와 설계의 빈칸을 구분하고, 반 완성 기준을 세웁니다.',
    keywords: ['같은 PRD 다른 결과', 'MVP'],
    goals: [
      '같은 PRD에서 서로 다른 결과가 나오는 이유를 설명할 수 있다.',
      '차이를 PRD 위반과 PRD의 빈칸으로 나누어 판단할 수 있다.',
      '기본판과 완성판을 비교해 완성돼 보이는 이유를 찾는다.',
      '반 완성 기준 초안과 조별 후보를 정한다.',
    ],
    outcome: 'A·B 검증 비교 · 검증 기록',
    status: 'open',
    slideCount: 18,
    slidesPath: 'lessons/lesson-03/slides/',
    thumbsPath: 'lessons/lesson-03/thumbs/',
    pdfFile: null,
    pdfName: null,
  },
  {
    id: 4,
    slug: 'lesson-04',
    phase: 'BUILD',
    title: '인터랙티브 웹 프로젝트 · Claude Code로 구현하고 개선하기',
    shortTitle: '인터랙티브 웹 프로젝트',
    summary: 'Claude Code로 내 아이디어를 직접 만들고, 계획과 결과를 비교하며 개선합니다.',
    keywords: ['Claude Code', '계획 승인', 'Git 커밋', '개발일지'],
    goals: [
      '아이디어가 실행 가능한 웹으로 완성되는 과정을 분석한다.',
      '계획을 먼저 받아 읽고 승인하거나 거절할 수 있다.',
      '본 것과 바라는 것을 나누어 말해 결과를 고칠 수 있다.',
      '잘 된 지점을 커밋으로 남기고 되돌릴 수 있다.',
    ],
    outcome: '개인 작품 index.html · 개발일지.md',
    status: 'open',
    slideCount: 15,
    slidesPath: 'lessons/lesson-04/slides/',
    thumbsPath: 'lessons/lesson-04/thumbs/',
    pdfFile: null,
    pdfName: null,
  },
  {
    id: 5,
    slug: 'lesson-05',
    phase: 'BUILD',
    title: '인터랙티브 웹 프로젝트 2 · 개선하고 배포하기',
    shortTitle: '개선하고 배포하기',
    summary: '4회기 작품을 이어서 개선하고, 규칙을 고정해 인터넷에 배포합니다.',
    keywords: ['Claude Code', '개선 반복', 'CLAUDE.md', '배포 URL'],
    goals: [
      '앞 회기 작품을 이어받아 부족한 부분을 골라 개선한다.',
      '프로젝트 규칙을 CLAUDE.md로 고정해 매번 읽히게 한다.',
      '작품을 인터넷에 배포하고 주소를 확보한다.',
      '휴대폰과 다른 기기에서 열어 동작을 확인한다.',
    ],
    outcome: 'CLAUDE.md · 배포 URL',
    status: 'coming-soon',
    slideCount: 0,
    slidesPath: 'lessons/lesson-05/slides/',
    thumbsPath: 'lessons/lesson-05/thumbs/',
    pdfFile: null,
    pdfName: null,
  },
  {
    id: 6,
    slug: 'lesson-06',
    phase: 'BUILD',
    title: '우리 조 작품 정하기',
    shortTitle: '우리 조 작품 정하기',
    summary: '개인 작품을 시연해 조 대표작과 대표 주제를 정하고, 작품 기준과 역할을 확정합니다.',
    keywords: ['조 작품 결정', '대표 주제', '작품 기준', '핵심 흐름', '역할표'],
    goals: [
      '개인 작품을 시연하고 근거를 들어 조 대표작을 결정한다.',
      '작품 기준을 확정하고 이번 버전의 범위를 자른다.',
      '첫 화면과 핵심 흐름을 조원이 같은 그림으로 이해한다.',
      '역할을 나누고 조별 작업 방식을 정한다.',
    ],
    outcome: '조편성 · 대표 주제 · 작품 기준 · 역할표',
    status: 'coming-soon',
    slideCount: 0,
    slidesPath: 'lessons/lesson-06/slides/',
    thumbsPath: 'lessons/lesson-06/thumbs/',
    pdfFile: null,
    pdfName: null,
  },
  {
    id: 7,
    slug: 'lesson-07',
    phase: 'BUILD',
    title: '조별 작품 한 바퀴 완성',
    shortTitle: '조별 작품 한 바퀴 완성',
    summary: '시작부터 결과까지 조 작품이 끊기지 않고 동작하는가?',
    keywords: ['기능 완결', '데이터와 상태', '예외 처리', '버그 기록'],
    goals: [
      '기능·데이터·상태를 연결해 사용 흐름을 끝까지 완결한다.',
      '입력과 결과가 저장되고 다시 열어도 유지되게 만든다.',
      '발견한 버그를 기록하고 원인을 찾아 수정한다.',
      '잘 동작하는 지점마다 커밋으로 되돌릴 수 있게 남긴다.',
    ],
    outcome: '조별 기능 완결 · 버그 기록',
    status: 'coming-soon',
    slideCount: 0,
    slidesPath: 'lessons/lesson-07/slides/',
    thumbsPath: 'lessons/lesson-07/thumbs/',
    pdfFile: null,
    pdfName: null,
  },
  {
    id: 8,
    slug: 'lesson-08',
    phase: 'TEST',
    title: '사용자 검증과 UI 완성',
    shortTitle: '사용자 검증과 UI 완성',
    summary: '설명 없이도 사용할 수 있는가? 무엇이 이 작품을 기억하게 만드는가?',
    keywords: ['사용자 테스트', '관찰 기록', 'UI 폴리싱', '대표 연출', '크레딧'],
    goals: [
      '설명 없이 다른 사람이 사용하게 하고 행동을 관찰해 기록한다.',
      '관찰 결과를 지금 고칠 것과 나중에 할 것으로 분류해 반영한다.',
      '간격·정렬·색을 정리하고 대표 연출 하나를 완성한다.',
      '그림과 소리의 출처를 크레딧으로 남기고 전후를 비교한다.',
    ],
    outcome: 'UI 폴리싱 · 크레딧 · Before/After',
    status: 'coming-soon',
    slideCount: 0,
    slidesPath: 'lessons/lesson-08/slides/',
    thumbsPath: 'lessons/lesson-08/thumbs/',
    pdfFile: null,
    pdfName: null,
  },
  {
    id: 9,
    slug: 'lesson-09',
    phase: 'RELEASE',
    title: '출시 후보 완성',
    shortTitle: '출시 후보 완성',
    summary: '예상하지 못한 상황과 다른 기기에서도 작동하는가?',
    keywords: ['기능 동결', '예외 상황', '기기 호환', 'TEST.md · README.md'],
    goals: [
      '새 기능을 멈추고 작품 기준 전 항목을 점검한다.',
      '잘못된 입력과 예외 상황에서도 멈추지 않게 보완한다.',
      'PC와 휴대폰 등 다른 기기에서 동작을 확인한다.',
      '점검 결과와 사용법을 TEST.md · README.md로 남긴다.',
    ],
    outcome: '출시 후보본 · TEST.md · README.md',
    status: 'coming-soon',
    slideCount: 0,
    slidesPath: 'lessons/lesson-09/slides/',
    thumbsPath: 'lessons/lesson-09/thumbs/',
    pdfFile: null,
    pdfName: null,
  },
  {
    id: 10,
    slug: 'lesson-10',
    phase: 'RELEASE',
    title: '배포와 작품 쇼케이스',
    shortTitle: '배포와 작품 쇼케이스',
    summary: '완성 결과와 발전 과정을 어떻게 보여 줄까?',
    keywords: ['배포', 'URL · QR', '작품 갤러리', '회고'],
    goals: [
      '최종 버전을 배포하고 URL과 QR로 공유한다.',
      '조별로 작품과 개발 과정을 정해진 시간 안에 발표한다.',
      '작품 기준으로 서로의 결과를 평가한다.',
      '만들며 배운 것을 회고로 정리한다.',
    ],
    outcome: '정식 배포 · URL·QR · 작품 갤러리 · 회고',
    status: 'coming-soon',
    slideCount: 0,
    slidesPath: 'lessons/lesson-10/slides/',
    thumbsPath: 'lessons/lesson-10/thumbs/',
    pdfFile: null,
    pdfName: null,
  },
];

/* 실습 도구 — 강의자료 섹션의 '실습도구' 탭에 모아서 보여 줍니다.
 *  title : 도구 이름   lesson : 관련 회기 번호 (없으면 null)
 *  href  : 열 주소     label  : 버튼 글자       note : 한 줄 설명 */
const TOOLS = [
  {
    title: 'PRD 기획 도우미',
    lesson: 2,
    href: './PRD/index.html',
    label: 'PRD 기획 시작하기',
    note: '11단계를 따라가며 내 PRD를 완성하고 파일로 내려받습니다.',
  },
  {
    title: 'MVP 비교·검증 도우미',
    lesson: 3,
    href: './MVP/index.html',
    label: 'MVP 비교 시작하기',
    note: '같은 PRD로 두 버전을 만들어 비교하고, PRD의 빈칸을 찾아 채웁니다.',
  },
];
