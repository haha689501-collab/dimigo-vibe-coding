/* tools/build-gallery.py 가 만든 공개용 파일입니다. 직접 고치지 마세요. */
window.STUDENT_WORKS = {
 "updated": "2026-10-17 23:07",
 "days": 10,
 "labels": {
  "1": "익명게시판 PRD",
  "2": "익명게시판 완성",
  "3": "MVP 비교",
  "4": "인터랙티브 웹 프로젝트",
  "5": "기획 한 장 · V0.1",
  "6": "한 판 완성 · V0.5",
  "7": "플레이테스트",
  "8": "그림·소리 · V0.9",
  "9": "출시 · V1.0",
  "10": "데모데이"
 },
 "students": [
  {
   "id": "s01",
   "name": "고*형",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명_계시판_PRD.md",
      "ext": "md",
      "path": "s01/day1/익명_계시판_PRD.md",
      "size": 3870,
      "title": "익명 게시판",
      "text": "# 익명 게시판\n\n## 1. 프로젝트 개요\n\n교내 모든 학생들이 사용할 수 있는 익명 게시판\n\n- 카테고리: 생활\n\n## 2. 해결하려는 문제\n\n신입생 등의 학교에 적응이 잘 안된 사람, 혹은 낯을 가리는 사람들은 교내에서 궁금한 점이나 어려운 점 들을 해결할 때 조언을 구할만한 방법이 마땅치 않음.\n\n## 3. 주요 사용자\n\n- 사용자: 전교생\n- 사용 상황: 교내가 떠들썩한데 뭔 일이 일어났는지 모를 때\n- 사용 기기: 스마트폰\n\n## 4. User Flow\n\n1. 앱을 연다\n2. 로그인을 한다\n3. 원하는 카테고리에 들어간다\n4. 글쓰기 버튼을 눌러서 글을 쓴다 (제목, 내용+첨부된 사진,동영상,링크)\n5. 글쓰기 완료를 눌러 글을 올린다\n6. 글을 읽고 싶다면 카테고리 내에서 원하는 글을 눌러 들어간다\n7. 좋아요/싫어요 버튼이 있어 누를 수 있다. (두 버튼 중복 가능, 한 버튼은 한번만 누르고 다시 누르면 취소)\n8. 댓글 작성란에 댓글을 작성 후 작성하기 버튼을 눌러서 댓글을 작성한다.\n\n## 5. MVP\n\n### 이번에 만들 기능\n\n- 카테고리 탭 만들기 (전체/잡담/공부/IT)\n- 글 작성,수정,삭제\n- 글 보기 (좋아요, 싫어요)\n- 댓글 작성,수정,삭제\n- 글 검색\n\n### 나중에 만들 기능\n\n- 회원가입\n- 로그인/로그아웃\n\n## 6. UI 구성\n\n- 목록\n- 버튼\n- 검색\n\n- 가장 먼저 보여야 하는 것: 작성된 글들이 최신순으로 나열되어 있는 전체 탭 (글들이 제목, 작성일, 작성자 형태로 리스트 형식으로 구성되어 되어있음)\n- 가장 중요한 버튼: 글쓰기 버튼\n- 추가 메모: 전체적으로밝은 톤에다가 버튼이나 강조 색상은 연한 남색, 왼쪽 위에는 누르면 전체 탭으로 이동하는 Memo라고 적힌 로고가 있음, 오른쪽 위에는 사용자 프로필 버튼이 있고 누르면 설정과 로그아웃 버튼이 보임(설정이랑 로그아웃은 버튼만 만들고 구현은 나중에 구현하기), 로고랑 프로필 아래에는 전체/공부/IT 라는 카테고리가 있는 탭이 있고 탭을 누르면 해당 탭의 글 리스트로 이동함(처음에는 전체 탭만 만들고 다른 탭들 상호작용은 나중에 구현하기), 글쓰기 버튼은 오른쪽 아래에 있고 전체적으로 강조 색상으로 칠해져 있음\n\n## 7. Acceptance Criteria\n\n- 접속 시 '전체'탭으로 이동하여 모든 글들이 리스트로 보인다\n- 왼쪽 위에는 로고가 있고 오른쪽엔 프로필이 있다. 그 아래로는 카테고리 탭과 글들\n- 글쓰기 버튼을 누르면 글쓰기 화면으로 이동한다\n- 글쓰기 화면에선 제목, 내용을 입력받고 내용에는 사진, 동영상, 링크를 첨부할 수 있다.\n- 글쓰기 완료를 누르면 글이 리스트에 추가된다, 이때 최신순 나열을 하여 최근에 쓴 글이 가장 위로 오게 한다.\n- 리스트에서 각 항목은 제목, 작성일, 작성자가 적혀 있고 해당 항목을 누를 시 해당 글로 이동하여 읽을 수 있음\n- 저장된 글들은 글쓴이가 삭제를 하지 않는 한 사라지면 안됨. 새로고침, 서버 재시작 등의 이밴트가 발생하여도 저장된 내용이 유지되어야 한다.\n\n## 8. 수정 기록\n\n- 회원가입,로그인 기능을 나중으로 밀기로\n- 카테고리 목록 불일치, 최종 희망 카테고리는 (전체, 잡담, 공부, IT)임\n- 익명 게시판인데 작성자 표시를 요구함, 작성자는 사용자가 임의로 정한 자신의 닉네임이며 익명성이 보장됨\n- 글쓰기에서의 첨부파일 규격 제한이 없음 (최대 200MB)\n- 작성일 표시 형식은 ~전으로 표시 (예:3시간전, 7일전, 3개월전)\n"
     }
    ],
    "2": [
     {
      "name": "index.html",
      "ext": "html",
      "path": "s01/day2/index.html",
      "size": 50961,
      "title": "Memo · 익명 게시판"
     }
    ],
    "4": [
     {
      "name": "index.html",
      "ext": "html",
      "path": "s01/day4/index.html",
      "size": 49616,
      "title": "심야 카드룸"
     }
    ]
   }
  },
  {
   "id": "s02",
   "name": "국*연",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명게시판_PRD.txt",
      "ext": "txt",
      "path": "s02/day1/익명게시판_PRD.txt",
      "size": 2539,
      "title": "익명게시판_PRD"
     }
    ],
    "2": [
     {
      "name": "index.html",
      "ext": "html",
      "path": "s02/day2/index.html",
      "size": 27256,
      "title": "유리병 편지 — 익명 게시판"
     }
    ]
   }
  },
  {
   "id": "s03",
   "name": "김*현",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명게시판_PRD.txt",
      "ext": "txt",
      "path": "s03/day1/익명게시판_PRD.txt",
      "size": 1866,
      "title": "익명게시판_PRD"
     }
    ],
    "2": [
     {
      "name": "익명게시판-final.html",
      "ext": "html",
      "path": "s03/day2/익명게시판-final.html",
      "size": 43237,
      "title": "익명게시판"
     }
    ],
    "4": [
     {
      "name": "노을 옥상 대결 (Copy).html",
      "ext": "html",
      "path": "s03/day4/노을 옥상 대결 (Copy).html",
      "size": 262147,
      "title": "노을 옥상 대결 (Copy)"
     }
    ]
   }
  },
  {
   "id": "s04",
   "name": "김*온",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익게.txt",
      "ext": "txt",
      "path": "s04/day1/익게.txt",
      "size": 1846,
      "title": "익게"
     }
    ],
    "4": [
     {
      "name": "index.html",
      "ext": "html",
      "path": "s04/day4/index.html",
      "size": 68687,
      "title": "산길 드라이브 3D"
     }
    ]
   }
  },
  {
   "id": "s05",
   "name": "서*우",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명_게시판_PRD.txt",
      "ext": "txt",
      "path": "s05/day1/익명_게시판_PRD.txt",
      "size": 1723,
      "title": "익명_게시판_PRD"
     }
    ]
   }
  },
  {
   "id": "s06",
   "name": "안*호",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명게시판_PRD.txt",
      "ext": "txt",
      "path": "s06/day1/익명게시판_PRD.txt",
      "size": 1777,
      "title": "익명게시판_PRD"
     }
    ]
   }
  },
  {
   "id": "s07",
   "name": "오*현",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명게시판_PRD.txt",
      "ext": "txt",
      "path": "s07/day1/익명게시판_PRD.txt",
      "size": 3497,
      "title": "익명게시판_PRD"
     }
    ],
    "2": [
     {
      "name": "index.html",
      "ext": "html",
      "path": "s07/day2/index.html",
      "size": 35279,
      "title": "익명게시판"
     }
    ],
    "4": [
     {
      "name": "abyss-tower.html",
      "ext": "html",
      "path": "s07/day4/abyss-tower.html",
      "size": 164862,
      "title": "심연의 탑"
     }
    ]
   }
  },
  {
   "id": "s08",
   "name": "유*람",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명게시판_PRD.txt",
      "ext": "txt",
      "path": "s08/day1/익명게시판_PRD.txt",
      "size": 3579,
      "title": "익명게시판_PRD"
     }
    ],
    "2": [
     {
      "name": "index.html",
      "ext": "html",
      "path": "s08/day2/index.html",
      "size": 47165,
      "title": "익명게시판"
     }
    ],
    "4": [
     {
      "name": "마우스게임.html",
      "ext": "html",
      "path": "s08/day4/마우스게임.html",
      "size": 130087,
      "title": "Neon Soul"
     }
    ]
   }
  },
  {
   "id": "s09",
   "name": "이*서",
   "team": "",
   "days": {
    "2": [
     {
      "name": "index.html",
      "ext": "html",
      "path": "s09/day2/index.html",
      "size": 32518,
      "title": "익명 게시판"
     }
    ],
    "4": [
     {
      "name": "mubeop-delivery.html",
      "ext": "html",
      "path": "s09/day4/mubeop-delivery.html",
      "size": 70134,
      "title": "무법도시 배달"
     }
    ]
   }
  },
  {
   "id": "s10",
   "name": "이*수",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명_게시판_PRD.txt",
      "ext": "txt",
      "path": "s10/day1/익명_게시판_PRD.txt",
      "size": 2319,
      "title": "익명_게시판_PRD"
     }
    ],
    "2": [
     {
      "name": "anonymous-board-final.html",
      "ext": "html",
      "path": "s10/day2/anonymous-board-final.html",
      "size": 46499,
      "title": "조용한 게시판"
     }
    ],
    "4": [
     {
      "name": "exam-boss-raid.html",
      "ext": "html",
      "path": "s10/day4/exam-boss-raid.html",
      "size": 625422,
      "title": "시험 보스 레이드"
     },
     {
      "name": "game (4).html",
      "ext": "html",
      "path": "s10/day4/game (4).html",
      "size": 208257,
      "title": "야간 자습"
     },
     {
      "name": "tetris.html",
      "ext": "html",
      "path": "s10/day4/tetris.html",
      "size": 64332,
      "title": "NEON RIFT - 격변 테트리스"
     }
    ]
   }
  },
  {
   "id": "s11",
   "name": "이*후",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명게시판_PRD.txt",
      "ext": "txt",
      "path": "s11/day1/익명게시판_PRD.txt",
      "size": 1580,
      "title": "익명게시판_PRD"
     }
    ]
   }
  },
  {
   "id": "s12",
   "name": "이*호",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명게시판_PRD.txt",
      "ext": "txt",
      "path": "s12/day1/익명게시판_PRD.txt",
      "size": 4276,
      "title": "익명게시판_PRD"
     }
    ]
   }
  },
  {
   "id": "s13",
   "name": "이*연",
   "team": "",
   "days": {
    "4": [
     {
      "name": "index.html",
      "ext": "html",
      "path": "s13/day4/index.html",
      "size": 27993,
      "title": "DEAD SIGNAL — FPS"
     }
    ]
   }
  },
  {
   "id": "s14",
   "name": "정*욱",
   "team": "",
   "days": {
    "1": [
     {
      "name": "PRD.md",
      "ext": "md",
      "path": "s14/day1/PRD.md",
      "size": 3819,
      "title": "익명 게시판",
      "text": "# 익명 게시판\n\n## 1. 프로젝트 개요\n\n로그인은 하되 글에는 이름이 남지 않는, 누구나 참여할 수 있는 공개 익명 커뮤니티\n\n- 카테고리: 커뮤니티\n\n## 2. 해결하려는 문제\n\n익명성을 높이면 솔직해지지만 책임이 사라지고, 책임을 물으려 실명을 요구하면 아무도 말하지 않는다.\n\n### 해결 방식\n\n화면에는 작성자를 전혀 노출하지 않되, 서버에는 계정 식별자를 해시한 값만 저장한다.\n이용자는 익명으로 말할 수 있고, 운영자는 신고가 누적된 계정에 제재를 가할 수 있다.\n단, 이는 완전한 익명이 아니라 '운영자에게만 열린 익명'이므로, 서비스 안내 문구에서 절대 익명처럼 표현하지 않는다.\n\n## 3. 주요 사용자\n\n- 사용자: 이름이 드러나는 것이 부담스러워 말을 꺼내지 못하는, 10~30대 온라인 커뮤니티 이용자\n- 사용 상황: 실명으로는 꺼내기 어려운 이야기를, 주로 심야·이동 중에 짧게 쓰고 반응을 확인하는 상황\n- 사용 기기: 스마트폰 중심의 모바일 웹, 데스크톱 브라우저 병행 지원 (반응형)\n\n## 4. User Flow\n\n1. 웹사이트에 접속해 로그인한다.\n2. 게시글 목록에서 최근 글을 확인한다.\n3. 읽고 싶은 글을 눌러 본문과 댓글을 확인한다.\n4. 글쓰기 버튼을 눌러 익명으로 글을 작성한다.\n5. 작성한 글에 달린 댓글과 반응을 확인한다.\n6. 필요하면 '내가 쓴 글'에서 글을 삭제한다.\n\n## 5. MVP\n\n### 이번에 만들 기능\n\n- 회원가입 및 로그인\n- 익명 글 작성\n- 댓글 작성\n- 내가 쓴 글 확인 및 삭제\n- 게시글 신고\n\n### 나중에 만들 기능\n\n- 게시판 카테고리 분류\n- 관리자 숨김 처리\n\n## 6. UI 구성\n\n- 제목\n- 목록\n- 카드\n- 버튼\n\n- 가장 먼저 보여야 하는 것: 최신순으로 정렬된 익명 게시글 목록\n- 가장 중요한 버튼: 글쓰기\n- 추가 메모: 모바일 화면을 기준으로 설계하고, 데스크톱은 가운데 정렬로 폭만 제한한다. / 첫 화면은 상단 제목, 그 아래 게시글 카드 목록, 우측 하단에 고정된 글쓰기 버튼으로 구성한다. / 각 카드에는 글 제목, 본문 첫 두 줄, 작성 시각, 댓글 수, 회색 '익명' 라벨만 표시하고 작성자 정보는 어디에도 노출하지 않는다. / 밤에 쓰는 이용자가 많을 것으로 보아 다크 단일 테마로 구성하고, 강조색은 글쓰기 버튼 한 곳에만 사용한다. / 빈 화면일 때는 \"아직 글이 없습니다. 첫 글을 남겨보세요.\"를 안내한다.\n\n## 7. Acceptance Criteria\n\n- 로그인하지 않으면 글쓰기와 댓글 작성이 차단된다.\n- 작성한 글이 목록 맨 위에 나타난다.\n- 게시글과 댓글 어디에도 작성자의 아이디나 이름이 표시되지 않는다.\n- 내가 쓴 글 목록에서 본인 글만 확인하고 삭제할 수 있다.\n- 다른 사용자의 글 번호로 삭제를 요청해도 처리되지 않는다.\n- 댓글을 작성하면 해당 글의 댓글 수가 1 증가한다.\n- 제목이나 본문이 비어 있으면 글이 저장되지 않는다.\n- 신고가 3회 누적되면 해당 글이 목록에서 자동으로 숨겨진다.\n- 새로고침하거나 다시 접속해도 작성한 글과 댓글이 그대로 남아 있다.\n\n## 8. 기술 및 운영 조건\n\n- 기술 스택: Express + SQLite, 서버 렌더링 또는 REST API 방식\n- 개인정보: IP는 해시 후 저장하고 보관 기간 경과 시 파기, 게시물 테이블에 계정 원본 식별자를 저장하지 않음\n- 도배 방지: 계정 단위 작성 간격 제한\n- 비밀번호 저장: bcrypt 등 단방향 해시 사용"
     }
    ],
    "4": [
     {
      "name": "index.html",
      "ext": "html",
      "path": "s14/day4/index.html",
      "size": 54890,
      "title": "마켓보드 · 실시간 시세"
     }
    ]
   }
  },
  {
   "id": "s15",
   "name": "차*린",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명_게시판_PRD.txt",
      "ext": "txt",
      "path": "s15/day1/익명_게시판_PRD.txt",
      "size": 4093,
      "title": "익명_게시판_PRD"
     }
    ],
    "2": [
     {
      "name": "anonymous-board.html",
      "ext": "html",
      "path": "s15/day2/anonymous-board.html",
      "size": 26257,
      "title": "익명 게시판"
     }
    ]
   }
  },
  {
   "id": "s16",
   "name": "하*린",
   "team": "",
   "days": {
    "1": [
     {
      "name": "수정_익명계시판_PRD.txt",
      "ext": "txt",
      "path": "s16/day1/수정_익명계시판_PRD.txt",
      "size": 2331,
      "title": "수정_익명계시판_PRD"
     }
    ],
    "2": [
     {
      "name": "디미고_사랑방.html",
      "ext": "html",
      "path": "s16/day2/디미고_사랑방.html",
      "size": 27800,
      "title": "디미고 익명 사랑방"
     }
    ],
    "4": [
     {
      "name": "index-2.html",
      "ext": "html",
      "path": "s16/day4/index-2.html",
      "size": 161053,
      "title": "ORBIT LAB — 우주·로켓 3D 학습 플랫폼"
     }
    ]
   }
  },
  {
   "id": "s17",
   "name": "한*윤",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명게시판_PRD.txt",
      "ext": "txt",
      "path": "s17/day1/익명게시판_PRD.txt",
      "size": 2320,
      "title": "익명게시판_PRD"
     }
    ],
    "2": [
     {
      "name": "dimigo_board_final.html",
      "ext": "html",
      "path": "s17/day2/dimigo_board_final.html",
      "size": 18120,
      "title": "디미고 민원 게시판"
     }
    ],
    "4": [
     {
      "name": "moru-introduction.html",
      "ext": "html",
      "path": "s17/day4/moru-introduction.html",
      "size": 638190,
      "title": "모루 · 예시 데이터 체험"
     }
    ]
   }
  },
  {
   "id": "s18",
   "name": "한*화",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명_게시판_PRD.txt",
      "ext": "txt",
      "path": "s18/day1/익명_게시판_PRD.txt",
      "size": 1577,
      "title": "익명_게시판_PRD"
     }
    ],
    "2": [
     {
      "name": "index (2).html",
      "ext": "html",
      "path": "s18/day2/index (2).html",
      "size": 32518,
      "title": "익명 게시판"
     }
    ]
   }
  },
  {
   "id": "s19",
   "name": "홍*준",
   "team": "",
   "days": {
    "1": [
     {
      "name": "익명게시판_PRD.txt",
      "ext": "txt",
      "path": "s19/day1/익명게시판_PRD.txt",
      "size": 2242,
      "title": "익명게시판_PRD"
     }
    ],
    "2": [
     {
      "name": "index.html",
      "ext": "html",
      "path": "s19/day2/index.html",
      "size": 45541,
      "title": "익명게시판"
     }
    ],
    "4": [
     {
      "name": "Shard.html",
      "ext": "html",
      "path": "s19/day4/Shard.html",
      "size": 12024576,
      "title": "Shard",
      "desc": "AI 학습 데이터를 공유하고 좋더라·별로더라를 나누는 커뮤니티"
     }
    ]
   }
  }
 ]
};
