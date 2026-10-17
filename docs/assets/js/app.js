/* =====================================================================
 *  AI × BUILD | Claude Vibe Coding — 라우터 + 슬라이드 뷰어
 *
 *  · 해시 라우팅(#/lesson-01)을 사용합니다.
 *    GitHub Pages 의 저장소 하위 경로에서도 새로고침·뒤로가기가 정상 동작하고,
 *    별도의 서버 설정이 필요 없기 때문입니다.
 *  · 모든 경로는 상대 경로입니다. (하위 경로 배포 시 이미지가 깨지지 않도록)
 * ===================================================================== */
(function () {
  'use strict';

  var main = document.getElementById('main');
  var navEl = document.getElementById('siteNav');
  var navToggle = document.getElementById('navToggle');

  var SECTIONS = ['curriculum', 'lessons', 'process'];

  /* ---------- 유틸 ---------- */
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function pad3(n) { return ('00' + n).slice(-3); }
  function byId(id) { return LESSONS.filter(function (l) { return l.id === id; })[0]; }
  function bySlug(slug) { return LESSONS.filter(function (l) { return l.slug === slug; })[0]; }
  function isOpen(l) { return !!l && l.status === 'open' && l.slideCount > 0; }

  var openCount = LESSONS.filter(isOpen).length;

  /* ================================================================
   *  홈 화면
   * ================================================================ */
  function renderHome(tab) {
    var pct = Math.round((openCount / COURSE.totalSessions) * 100);

    var hero = '' +
      '<section class="hero">' +
        '<div class="wrap hero-inner">' +
          '<div>' +
            '<p class="eyebrow">' + esc(COURSE.school) + ' · ' + esc(COURSE.className) + '</p>' +
            '<h1 class="hero-title">' + esc(COURSE.heroTitle) + '</h1>' +
            '<p class="hero-subtitle">' + esc(COURSE.heroSub) + '</p>' +
            '<p class="hero-brand">AI <span class="hero-x">×</span> BUILD · ' +
              '<span class="hero-sub">' + esc(COURSE.subtitle) + '</span></p>' +
            '<p class="hero-lead">' + COURSE.leadLines.map(esc).join('<br>') + '</p>' +
            '<p class="hero-note">' + esc(COURSE.note) + '</p>' +
            '<div class="hero-actions">' +
              '<a class="btn btn-primary" href="#lessons">강의자료 보기</a>' +
              '<a class="btn btn-ghost" href="#curriculum">전체 커리큘럼</a>' +
            '</div>' +
          '</div>' +
          '<div class="progress-card">' +
            '<div class="progress-top">' +
              '<span class="progress-num"><em>' + openCount + '</em> / ' + COURSE.totalSessions + '</span>' +
              '<span class="progress-label">회기 자료 등록</span>' +
            '</div>' +
            '<div class="progress-bar">' +
              '<div class="progress-fill" style="width:' + pct + '%"></div>' +
            '</div>' +
            '<p class="progress-note">' + (openCount ? '1~' + openCount + '회기 자료 공개' : '자료 준비 중') +
              (openCount < COURSE.totalSessions ? ' · ' + (openCount + 1) + '~' + COURSE.totalSessions + '회기 준비 중' : '') + '</p>' +
          '</div>' +
        '</div>' +
      '</section>';

    /* --- 회기 카드 --- */
    var cards = LESSONS.map(function (l) {
      var open = isOpen(l);
      var keys = l.keywords.map(function (k) { return '<span class="key">' + esc(k) + '</span>'; }).join('');

      var foot = open
        ? '<a class="btn btn-primary card-link" href="#/' + l.slug + '">' +
            '강의자료 보기<span aria-hidden="true">→</span>' +
            '<span class="sr-only"> : ' + esc(l.shortTitle) + '</span></a>' +
          '<span class="card-slides">' + l.slideCount + '장</span>'
        : '<button class="btn btn-quiet" type="button" disabled ' +
            'aria-label="' + l.id + '회기 ' + esc(l.shortTitle) + ' 자료 준비 중">준비 중</button>';

      return '' +
        '<article class="lesson-card ' + (open ? 'is-open' : 'is-soon') + '">' +
          '<div class="card-top">' +
            '<span class="card-no"><b>' + l.id + '</b>회기</span>' +
            '<span class="badge ' + (open ? 'badge-open' : 'badge-soon') + '">' +
              (open ? '자료 공개' : '준비 중') + '</span>' +
          '</div>' +
          '<h3 class="card-title">' + esc(l.shortTitle) + '</h3>' +
          '<p class="card-summary">' + esc(l.summary) + '</p>' +
          '<div class="card-keys">' + keys + '</div>' +
          '<div class="card-outcome">' +
            '<span class="lbl">결과물</span>' +
            '<span class="val">' + esc(l.outcome) + '</span>' +
          '</div>' +
          '<div class="card-foot">' + foot + '</div>' +
        '</article>';
    }).join('');

    /* --- 실습 도구 카드 --- */
    var toolCards = TOOLS.map(function (t) {
      return '' +
        '<article class="lesson-card is-open tool-item">' +
          '<div class="card-top">' +
            '<span class="tool-label">실습 도구</span>' +
            (t.lesson ? '<span class="badge badge-phase">' + t.lesson + '회기</span>' : '') +
          '</div>' +
          '<h3 class="card-title">' + esc(t.title) + '</h3>' +
          '<p class="card-summary">' + esc(t.note || '') + '</p>' +
          '<div class="card-foot">' +
            '<a class="btn btn-primary card-link" href="' + esc(t.href) + '">' +
              esc(t.label) + '<span aria-hidden="true">→</span></a>' +
          '</div>' +
        '</article>';
    }).join('');

    var onTools = tab === 'tools';
    function tabBtn(key, label) {
      var on = (key === 'tools') === onTools;
      return '<a class="tab' + (on ? ' is-on' : '') + '" role="tab" id="tab-' + key + '" ' +
        'href="#/' + key + '" aria-selected="' + on + '" aria-controls="panel-' + key + '">' + label + '</a>';
    }

    var lessonsSection = '' +
      '<section class="section" id="lessons">' +
        '<div class="wrap">' +
          '<div class="section-head">' +
            '<p class="eyebrow">강의자료</p>' +
            '<h2>' + (onTools ? '실습도구' : '회기별 강의자료') + '</h2>' +
            '<p>' + (onTools
              ? '수업 중 직접 해 보는 실습 도구입니다. 버튼을 누르면 바로 시작합니다.'
              : '카드를 클릭하면 해당 회기의 PPT 슬라이드를 웹에서 바로 볼 수 있습니다.') + '</p>' +
          '</div>' +
          '<div class="tabs" role="tablist" aria-label="강의자료 보기 방식">' +
            tabBtn('lessons', '강의자료 보기') + tabBtn('tools', '실습도구') +
          '</div>' +
          '<div role="tabpanel" id="panel-' + (onTools ? 'tools' : 'lessons') + '" ' +
            'aria-labelledby="tab-' + (onTools ? 'tools' : 'lessons') + '">' +
            '<div class="lesson-grid">' + (onTools ? toolCards : cards) + '</div>' +
          '</div>' +
        '</div>' +
      '</section>';

    /* --- 전체 커리큘럼 표 --- */
    var rows = LESSONS.map(function (l) {
      var open = isOpen(l);
      return '' +
        '<tr>' +
          '<td class="c-no">' + l.id + '</td>' +
          '<td class="c-title">' + esc(l.title) + '</td>' +
          '<td class="c-key">' + l.keywords.map(esc).join(' · ') + '</td>' +
          '<td class="c-out">' + esc(l.outcome) + '</td>' +
          '<td class="c-st"><span class="badge ' + (open ? 'badge-open' : 'badge-soon') + '">' +
            (open ? '자료 공개' : '준비 중') + '</span></td>' +
        '</tr>';
    }).join('');

    var curriculumSection = '' +
      '<section class="section section-alt" id="curriculum">' +
        '<div class="wrap">' +
          '<div class="section-head">' +
            '<p class="eyebrow">전체 커리큘럼</p>' +
            '<h2>10회기 한눈에 보기</h2>' +
            '<p>기초(1~3회기) → 개인 작품 구현과 배포(4~5회기) → 조 작품 완성(6~7회기) → 사용자 검증과 UI 완성(8회기) → 안정화와 공개(9~10회기)로 이어집니다.</p>' +
          '</div>' +
          '<div class="table-scroll">' +
            '<table class="curriculum-table">' +
              '<caption class="sr-only">1회기부터 10회기까지 주제와 결과물, 자료 상태</caption>' +
              '<thead><tr>' +
                '<th scope="col">회기</th><th scope="col">주제</th>' +
                '<th scope="col">핵심 키워드</th><th scope="col">결과물</th>' +
                '<th scope="col">자료</th>' +
              '</tr></thead>' +
              '<tbody>' + rows + '</tbody>' +
            '</table>' +
          '</div>' +
        '</div>' +
      '</section>';

    /* --- 프로젝트 과정 --- */
    var order = ['IDEA', 'SPEC', 'BUILD', 'TEST', 'RELEASE'];
    var railItems = order.map(function (key) {
      var ids = LESSONS.filter(function (l) { return l.phase === key; })
                       .map(function (l) { return l.id; });
      var range = ids.length > 1 ? ids[0] + '~' + ids[ids.length - 1] + '회기' : ids[0] + '회기';
      return '' +
        '<li class="rail-item">' +
          '<span class="rail-step">STEP ' + (order.indexOf(key) + 1) + '</span>' +
          '<h3>' + PHASES[key].label + '</h3>' +
          '<p>' + esc(PHASES[key].desc) + '</p>' +
          '<p class="rail-sessions">' + range + '</p>' +
          '<div class="rail-line"></div>' +
        '</li>';
    }).join('');

    var processSection = '' +
      '<section class="section" id="process">' +
        '<div class="wrap">' +
          '<div class="section-head">' +
            '<p class="eyebrow">프로젝트 과정</p>' +
            '<h2>IDEA → SPEC → BUILD → TEST → RELEASE</h2>' +
            '<p>1~3회기에 기초를 익히고, 4~5회기에 각자 Claude Code로 만들어 배포합니다. 6회기부터 조 작품 하나를 골라 한 바퀴 완성하고, 사용자 검증과 UI 완성을 거쳐 배포와 작품 쇼케이스로 마칩니다.</p>' +
          '</div>' +
          '<ul class="rail">' + railItems + '</ul>' +
        '</div>' +
      '</section>';

    main.innerHTML = hero + lessonsSection + curriculumSection + processSection;
    document.title = 'AI × BUILD | Claude Vibe Coding · 한국디지털미디어고등학교';
  }

  /* ================================================================
   *  강의자료(슬라이드 뷰어) 화면
   * ================================================================ */
  var viewer = null;   // 현재 뷰어 상태

  function renderLesson(lesson, startIndex) {
    var prev = byId(lesson.id - 1);
    var next = byId(lesson.id + 1);

    function navBtn(target, dir) {
      var label = dir === 'prev' ? '이전 회기' : '다음 회기';
      var arrow = dir === 'prev' ? '←' : '→';
      if (!target) {
        return '<button class="btn btn-quiet" type="button" disabled>' +
               (dir === 'prev' ? arrow + ' ' + label : label + ' ' + arrow) + '</button>';
      }
      if (!isOpen(target)) {
        return '<button class="btn btn-quiet" type="button" disabled ' +
               'aria-label="' + label + '(' + target.id + '회기) 준비 중">' +
               '<span class="nav-soon">' + esc(label) +
               '<small>' + target.id + '회기 · 준비 중</small></span></button>';
      }
      return '<a class="btn btn-ghost" href="#/' + target.slug + '" ' +
             'aria-label="' + label + ' : ' + target.id + '회기 ' + esc(target.shortTitle) + '">' +
             (dir === 'prev' ? arrow + ' ' + label : label + ' ' + arrow) + '</a>';
    }

    var goals = lesson.goals.map(function (g) { return '<li>' + esc(g) + '</li>'; }).join('');

    var thumbs = '';
    for (var i = 1; i <= lesson.slideCount; i++) {
      thumbs += '<button class="thumb" type="button" data-index="' + (i - 1) + '" ' +
                'aria-label="' + i + '번째 슬라이드로 이동">' +
                '<img src="' + esc(lesson.thumbsPath) + pad3(i) + '.webp" alt="" ' +
                'loading="lazy" decoding="async" width="400" height="225">' +
                '<span class="n">' + i + '</span></button>';
    }

    main.innerHTML = '' +
      '<div class="wrap lesson-page">' +
        '<div class="lesson-bar">' +
          '<div class="lesson-bar-main">' +
            '<div class="crumb">' +
              '<a class="back-link" href="#/"><span aria-hidden="true">←</span> 전체 커리큘럼</a>' +
              '<span class="pos"><b>' + lesson.id + '회기</b> / ' + COURSE.totalSessions + '회기</span>' +
              '<span class="badge badge-phase">' + lesson.phase + '</span>' +
            '</div>' +
            '<h1 class="lesson-title">' + esc(lesson.title) + '</h1>' +
            '<div class="lesson-meta">' +
              '<details class="meta-box meta-goals" id="goalsBox">' +
                '<summary><span class="lbl">오늘의 학습 목표</span>' +
                  '<span class="chev" aria-hidden="true"></span></summary>' +
                '<ul>' + goals + '</ul>' +
              '</details>' +
              '<div class="meta-box">' +
                '<span class="lbl">오늘의 핵심 결과물</span>' +
                '<p class="outcome-val">' + esc(lesson.outcome) + '</p>' +
              '</div>' +
            '</div>' +
          '</div>' +
          '<div class="lesson-side">' +
            '<div class="lesson-nav">' + navBtn(prev, 'prev') + navBtn(next, 'next') + '</div>' +
          '</div>' +
        '</div>' +

        '<div class="viewer">' +
          '<div class="stage-shell">' +
            '<div class="stage" id="stage">' +
              '<img id="slideImg" alt="" decoding="async">' +
              '<button class="stage-arrow stage-prev" id="prevBtn" type="button" aria-label="이전 슬라이드">' +
                '<span aria-hidden="true">‹</span></button>' +
              '<button class="stage-arrow stage-next" id="nextBtn" type="button" aria-label="다음 슬라이드">' +
                '<span aria-hidden="true">›</span></button>' +
              '<div class="fs-counter" id="fsCounter" aria-hidden="true"></div>' +
            '</div>' +
          '</div>' +

          '<div class="controls">' +
            '<button class="ctrl-btn" id="prevBtn2" type="button">' +
              '<span aria-hidden="true">←</span> 이전</button>' +
            '<span class="counter" id="counter" role="status" aria-live="polite"></span>' +
            '<button class="ctrl-btn" id="nextBtn2" type="button">' +
              '다음 <span aria-hidden="true">→</span></button>' +
            '<button class="ctrl-btn" id="fsBtn" type="button">전체화면</button>' +
          '</div>' +
          '<p class="hint">키보드 <kbd>←</kbd> <kbd>→</kbd> 로 이동 · 모바일에서는 좌우로 밀어서 이동</p>' +

          '<div class="thumbs-head">' +
            '<h2>슬라이드 목록</h2>' +
            '<span class="cnt">전체 ' + lesson.slideCount + '장</span>' +
          '</div>' +
          '<div class="thumbs" id="thumbs">' + thumbs + '</div>' +

        '</div>' +
      '</div>';

    // 넓은 화면에서는 학습 목표를 펼쳐 두고, 모바일에서는 접어 슬라이드를 먼저 보여줍니다.
    var goalsBox = document.getElementById('goalsBox');
    if (goalsBox && window.matchMedia('(min-width: 861px)').matches) {
      goalsBox.open = true;
    }

    document.title = lesson.id + '회기 · ' + lesson.title + ' | AI × BUILD';
    setupViewer(lesson, startIndex);
  }

  /* ---------- 슬라이드 뷰어 동작 ---------- */
  function setupViewer(lesson, startIndex) {
    var img = document.getElementById('slideImg');
    var stage = document.getElementById('stage');
    var counter = document.getElementById('counter');
    var fsCounter = document.getElementById('fsCounter');
    var thumbsWrap = document.getElementById('thumbs');
    var prevBtns = [document.getElementById('prevBtn'), document.getElementById('prevBtn2')];
    var nextBtns = [document.getElementById('nextBtn'), document.getElementById('nextBtn2')];
    var fsBtn = document.getElementById('fsBtn');
    var thumbEls = Array.prototype.slice.call(thumbsWrap.querySelectorAll('.thumb'));

    var index = Math.min(Math.max(startIndex || 0, 0), lesson.slideCount - 1);

    // 다른 화면으로 이동할 때 정리할 전역 리스너 목록
    var cleanups = [];
    function on(target, type, fn) {
      target.addEventListener(type, fn);
      cleanups.push(function () { target.removeEventListener(type, fn); });
    }

    viewer = {
      lesson: lesson,
      get index() { return index; },
      go: go,
      destroy: function () {
        for (var i = 0; i < cleanups.length; i++) cleanups[i]();
        cleanups = [];
      }
    };

    function srcFor(i) { return lesson.slidesPath + pad3(i + 1) + '.webp'; }

    function preload(i) {
      if (i < 0 || i >= lesson.slideCount) return;
      var im = new Image();
      im.decoding = 'async';
      im.src = srcFor(i);
    }

    function go(next, pushUrl) {
      var target = Math.min(Math.max(next, 0), lesson.slideCount - 1);
      index = target;
      render();
      if (pushUrl !== false) updateHash();
    }

    function updateHash() {
      // 슬라이드 이동은 히스토리를 쌓지 않습니다.
      // (뒤로가기를 누르면 슬라이드가 아니라 커리큘럼으로 돌아가야 하기 때문)
      var h = '#/' + lesson.slug + (index > 0 ? '/' + (index + 1) : '');
      if (location.hash !== h) {
        history.replaceState(null, '', h);
      }
    }

    function render() {
      img.src = srcFor(index);
      img.alt = lesson.id + '회기 ' + lesson.title + ' — ' +
                (index + 1) + '번째 슬라이드 (전체 ' + lesson.slideCount + '장)';

      var txt = '<b>' + (index + 1) + '</b> <span class="total">/ ' + lesson.slideCount + '</span>';
      counter.innerHTML = txt;
      fsCounter.textContent = (index + 1) + ' / ' + lesson.slideCount;

      var atFirst = index === 0;
      var atLast = index === lesson.slideCount - 1;
      prevBtns.forEach(function (b) { if (b) b.disabled = atFirst; });
      nextBtns.forEach(function (b) { if (b) b.disabled = atLast; });

      thumbEls.forEach(function (t, i) {
        var cur = i === index;
        t.classList.toggle('is-current', cur);
        if (cur) {
          t.setAttribute('aria-current', 'true');
          scrollThumbIntoView(t);
        } else {
          t.removeAttribute('aria-current');
        }
      });

      preload(index + 1);
      preload(index - 1);
    }

    function scrollThumbIntoView(el) {
      var wrapRect = thumbsWrap.getBoundingClientRect();
      var elRect = el.getBoundingClientRect();
      if (elRect.left < wrapRect.left || elRect.right > wrapRect.right) {
        thumbsWrap.scrollTo({
          left: el.offsetLeft - (thumbsWrap.clientWidth - el.offsetWidth) / 2,
          behavior: 'smooth'
        });
      }
    }

    img.addEventListener('error', function () {
      img.alt = (index + 1) + '번째 슬라이드 이미지를 불러오지 못했습니다.';
    });

    /* 슬라이드 크기를 실제 레이아웃에 맞춰 계산합니다.
       회기마다 상단 높이가 다르므로(실습 도구 카드 등) 고정값을 쓰지 않습니다.
       조작 버튼까지 한 화면에 들어오도록 남은 높이를 재어 16:9 로 환산합니다. */
    var shell = stage.parentElement;
    function sizeStage() {
      if (document.fullscreenElement || document.webkitFullscreenElement) {
        stage.style.maxWidth = '';
        return;
      }
      stage.style.maxWidth = '';                 // 먼저 초기화해야 위치를 정확히 잰다
      if (window.innerWidth <= 620) return;      // 모바일은 화면 폭을 그대로 쓴다
      var top = shell.getBoundingClientRect().top + (window.pageYOffset || 0);
      // 조작 버튼 줄까지는 반드시 화면 안에 들어오게 하고,
      // 그 아래 안내 문구는 스크롤로 봐도 되므로 계산에서 뺍니다. (슬라이드를 더 크게)
      var below = 16;
      var c = document.querySelector('.controls');
      if (c) below += c.offsetHeight + 12;
      var avail = window.innerHeight - top - below;
      if (avail < 220) avail = 220;              // 너무 작아지지 않게
      var w = Math.min(1180, avail * 16 / 9);
      stage.style.maxWidth = w + 'px';
    }

    var sizeTimer = null;
    function sizeStageSoon() {
      clearTimeout(sizeTimer);
      sizeTimer = setTimeout(sizeStage, 60);
    }
    on(window, 'resize', sizeStageSoon);
    var gb = document.getElementById('goalsBox');
    if (gb) gb.addEventListener('toggle', sizeStage);
    on(document, 'fullscreenchange', sizeStageSoon);
    on(document, 'webkitfullscreenchange', sizeStageSoon);
    sizeStage();

    prevBtns.forEach(function (b) { if (b) b.addEventListener('click', function () { go(index - 1); }); });
    nextBtns.forEach(function (b) { if (b) b.addEventListener('click', function () { go(index + 1); }); });

    thumbEls.forEach(function (t) {
      t.addEventListener('click', function () { go(parseInt(t.dataset.index, 10)); });
    });

    /* 전체화면 */
    if (!stage.requestFullscreen && !stage.webkitRequestFullscreen) {
      fsBtn.hidden = true;
    } else {
      fsBtn.addEventListener('click', function () {
        var fsEl = document.fullscreenElement || document.webkitFullscreenElement;
        if (fsEl) {
          (document.exitFullscreen || document.webkitExitFullscreen).call(document);
        } else {
          (stage.requestFullscreen || stage.webkitRequestFullscreen).call(stage);
        }
      });
      on(document, 'fullscreenchange', syncFsLabel);
      on(document, 'webkitfullscreenchange', syncFsLabel);
    }
    function syncFsLabel() {
      var on = !!(document.fullscreenElement || document.webkitFullscreenElement);
      fsBtn.textContent = on ? '전체화면 종료' : '전체화면';
    }

    /* 터치 스와이프 */
    var tx = 0, ty = 0, tracking = false;
    stage.addEventListener('touchstart', function (e) {
      if (e.touches.length !== 1) { tracking = false; return; }
      tx = e.touches[0].clientX; ty = e.touches[0].clientY; tracking = true;
    }, { passive: true });
    stage.addEventListener('touchend', function (e) {
      if (!tracking) return;
      tracking = false;
      var t = e.changedTouches[0];
      var dx = t.clientX - tx, dy = t.clientY - ty;
      if (Math.abs(dx) > 45 && Math.abs(dx) > Math.abs(dy) * 1.4) {
        go(dx < 0 ? index + 1 : index - 1);
      }
    }, { passive: true });

    render();
    // 첫 화면은 히스토리를 건드리지 않고, 시작 슬라이드가 1장이 아닐 때만 주소를 맞춥니다.
    if (index > 0) updateHash();
  }

  /* 키보드 (뷰어가 열려 있을 때만) */
  document.addEventListener('keydown', function (e) {
    if (!viewer) return;
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    var tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea' || e.target.isContentEditable) return;

    var n = viewer.lesson.slideCount;
    switch (e.key) {
      case 'ArrowLeft':  case 'PageUp':   viewer.go(viewer.index - 1); break;
      case 'ArrowRight': case 'PageDown': viewer.go(viewer.index + 1); break;
      case 'Home': viewer.go(0); break;
      case 'End':  viewer.go(n - 1); break;
      default: return;
    }
    e.preventDefault();
  });

  /* ================================================================
   *  준비 중 화면
   * ================================================================ */
  function renderComingSoon(lesson) {
    main.innerHTML = '' +
      '<div class="wrap lesson-page">' +
        '<div class="crumb">' +
          '<a class="back-link" href="#/"><span aria-hidden="true">←</span> 전체 커리큘럼</a>' +
          '<span class="pos"><b>' + lesson.id + '회기</b> / ' + COURSE.totalSessions + '회기</span>' +
        '</div>' +
        '<div class="notice">' +
          '<span class="badge badge-soon">준비 중</span>' +
          '<h2 style="margin-top:12px">' + lesson.id + '회기 · ' + esc(lesson.shortTitle) + '</h2>' +
          '<p>이 회기의 강의자료는 아직 등록되지 않았습니다.</p>' +
          '<a class="btn btn-primary" href="#/">전체 커리큘럼으로 돌아가기</a>' +
        '</div>' +
      '</div>';
    document.title = lesson.id + '회기 준비 중 | AI × BUILD';
  }

  function renderNotFound() {
    main.innerHTML = '' +
      '<div class="wrap lesson-page">' +
        '<div class="notice">' +
          '<h2>페이지를 찾을 수 없습니다</h2>' +
          '<p>주소를 다시 확인해 주세요.</p>' +
          '<a class="btn btn-primary" href="#/">전체 커리큘럼으로 돌아가기</a>' +
        '</div>' +
      '</div>';
  }

  /* ================================================================
   *  라우터
   * ================================================================ */
  function route() {
    if (viewer && viewer.destroy) viewer.destroy();
    viewer = null;
    if (navEl) navEl.classList.remove('is-open');
    if (navToggle) navToggle.setAttribute('aria-expanded', 'false');

    var raw = location.hash.replace(/^#\/?/, '');
    var parts = raw.split('/').filter(Boolean);
    var head = parts[0] || '';

    // #/lesson-01  또는  #/lesson-01/5
    if (/^lesson-\d{2}$/.test(head)) {
      var lesson = bySlug(head);
      if (!lesson) { renderNotFound(); window.scrollTo(0, 0); return; }
      if (!isOpen(lesson)) { renderComingSoon(lesson); window.scrollTo(0, 0); return; }
      var n = parseInt(parts[1], 10);
      renderLesson(lesson, isNaN(n) ? 0 : n - 1);
      window.scrollTo(0, 0);
      setActiveNav('lessons');
      return;
    }

    // 홈 + 섹션 이동  (#/tools 는 강의자료 섹션의 실습도구 탭)
    var tools = head === 'tools';
    renderHome(tools ? 'tools' : 'lessons');
    if (tools) head = 'lessons';
    if (SECTIONS.indexOf(head) !== -1) {
      var el = document.getElementById(head);
      if (el) el.scrollIntoView({ behavior: 'auto', block: 'start' });
      setActiveNav(head);
    } else {
      window.scrollTo(0, 0);
      setActiveNav(null);
    }
  }

  function setActiveNav(key) {
    if (!navEl) return;
    Array.prototype.forEach.call(navEl.querySelectorAll('a'), function (a) {
      a.classList.toggle('is-active', a.dataset.nav === key);
    });
  }

  /* 헤더 앵커(#lessons)를 해시 라우트(#/lessons)로 통일 */
  document.addEventListener('click', function (e) {
    var a = e.target.closest ? e.target.closest('a[href^="#"]') : null;
    if (!a) return;
    var href = a.getAttribute('href');
    if (/^#(curriculum|lessons|process)$/.test(href)) {
      e.preventDefault();
      location.hash = '#/' + href.slice(1);
    }
  });

  if (navToggle) {
    navToggle.addEventListener('click', function () {
      var open = navEl.classList.toggle('is-open');
      navToggle.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
  }

  // 강의자료는 보기 전용 — 슬라이드 이미지 우클릭 저장 · 끌어서 저장을 막습니다.
  ['contextmenu', 'dragstart'].forEach(function (type) {
    document.addEventListener(type, function (e) {
      if (e.target.closest && e.target.closest('.stage, .thumbs')) e.preventDefault();
    });
  });

  // 공개용 갤러리(docs/gallery — 이름 가린 사본)가 있을 때만 링크를 보여 줍니다.
  // 실명이 들어 있는 원본 docs/students 는 올리지 않습니다 (.gitignore).
  var studentsLink = document.getElementById('studentsLink');
  if (studentsLink) {
    if (location.protocol === 'file:') studentsLink.hidden = false;
    else fetch('gallery/index.html', { method: 'HEAD' })
      .then(function (r) { if (r.ok) studentsLink.hidden = false; })
      .catch(function () {});
  }

  window.addEventListener('hashchange', route);
  route();
})();
