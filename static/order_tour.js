/**
 * 叫貨工具操作教學（guided tour）
 *
 * 可重用引擎，不改叫貨流程。正式步驟在另一個分支的
 * static/order_flow_tour_steps.js（window.ORDER_FLOW_TOUR_STEPS）。
 * 那個檔案載入後，本引擎會自己 attach。kitchen/base.html 已引入本檔。
 *
 * 步驟（與 ORDER_FLOW_TOUR_STEPS 相同，另可加跟做與暫停）：
 *   {
 *     path: '/admin/order-tool/summary', // 這一步所在網址，不含結尾斜線
 *     exact: true,                       // true 只對這個路徑；否則子路徑也算
 *     selector: 'a.summary-next',        // 聚光燈目標。省略或找不到就只顯示置中說明
 *     title: '去選學校',
 *     text: '菜加好後，按「下一步：選擇學校」。',
 *     before: function () {},            // 顯示前執行，可回傳 Promise
 *     advanceOn: 'click',                // 或 waitFor。click / input / change / navigate
 *                                        // 也可 { event: 'change', selector: '#qty' }
 *   }
 *
 * 跨頁：進度存在 localStorage（order-tool-tour-progress）。
 * 導覽進行中換到下一步的頁面（選菜 → 菜色用量表 → 採購叫貨）會從對的步驟接下去。
 * 暫停或結束導覽之後，每一頁都會出現「繼續導覽」，從停下的地方恢復。
 * 亮起來的目標可以照常點、照常填；其餘畫面是暗的。說明卡會平滑移到下一個目標並指向它。
 *
 *   startTour(steps)
 *   OrderTour.attach(steps)
 *   OrderTour.pauseTour()
 *   OrderTour.resumeTour()
 *   OrderTour.endTour()          // 先停下，之後可按繼續導覽
 *   OrderTour.mountHelpButton(parent, { steps, label: '操作教學' })
 *   OrderTour.mountFirstVisitBanner(parent, { steps, storageKey: 'order-tool-tour-dismissed' })
 *
 * 按鈕：暫停、結束導覽（最後一步為關閉）、上一步、下一步。
 * 鍵盤：← → 、Esc 暫停。正在輸入時不會搶方向鍵。
 * 示範：/static/order_tour_demo.html
 */
(function () {
  'use strict';

  const SVG_NS = 'http://www.w3.org/2000/svg';
  const DEFAULT_SEEN_KEY = 'order-tool-tour-dismissed';
  const DEFAULT_PROGRESS_KEY = 'order-tool-tour-progress';

  let steps = [];
  let catalog = [];
  let index = -1;
  let run = 0;
  let busy = false;
  let onEnd = null;
  let returnFocus = null;
  let root = null;
  let holePath = null;
  let holeSvg = null;
  let resumeBtn = null;
  let placeQueued = false;
  let listening = false;
  let snapNext = false;
  let holeNow = null;
  let holeFrame = 0;
  let inputTimer = 0;
  let progressKey = DEFAULT_PROGRESS_KEY;

  function seen(key) {
    try { return window.localStorage.getItem(key) === '1'; }
    catch (err) { return false; }
  }

  function markSeen(key) {
    try { window.localStorage.setItem(key, '1'); }
    catch (err) { /* 隱私模式寫不進去時，這次仍讓導覽繼續。 */ }
  }

  function loadProgress() {
    try {
      const raw = window.localStorage.getItem(progressKey);
      if (!raw) return null;
      const data = JSON.parse(raw);
      if (!data || typeof data.index !== 'number' || !data.status) return null;
      return data;
    } catch (err) {
      return null;
    }
  }

  function saveProgress(status, i) {
    try {
      window.localStorage.setItem(progressKey, JSON.stringify({ v: 1, status: status, index: i }));
    } catch (err) { /* 換頁後可能無法接續，這一頁仍可看完。 */ }
  }

  function stripPath(value) {
    const raw = String(value || '/');
    const cut = raw.indexOf('?');
    const path = cut === -1 ? raw : raw.slice(0, cut);
    if (path === '/') return '/';
    return path.replace(/\/+$/, '') || '/';
  }

  function currentPath() {
    return stripPath(location.pathname || '/');
  }

  function stepMatches(step) {
    if (!step || !step.path) return true;
    const want = stripPath(step.path);
    const cur = currentPath();
    const pathOk = step.exact ? cur === want : (cur === want || cur.startsWith(want + '/'));
    if (!pathOk) return false;
    const cut = step.path.indexOf('?');
    if (cut === -1) return true;
    const wantQs = new URLSearchParams(step.path.slice(cut + 1));
    const curQs = new URLSearchParams(location.search);
    for (const pair of wantQs) {
      if (curQs.get(pair[0]) !== pair[1]) return false;
    }
    return true;
  }

  function normalize(input) {
    const source = typeof input === 'function' ? input() : input;
    if (!Array.isArray(source)) return [];
    return source.filter(Boolean).map(function (step) {
      const advance = step.advanceOn != null ? step.advanceOn : (step.waitFor != null ? step.waitFor : null);
      return {
        path: step.path ? String(step.path) : null,
        exact: step.exact === true,
        selector: step.selector == null || step.selector === '' ? null : step.selector,
        title: step.title == null ? '' : String(step.title),
        text: step.text == null ? '' : String(step.text),
        before: typeof step.before === 'function' ? step.before : null,
        advanceOn: advance,
        optional: step.optional === true,
      };
    });
  }

  function followMode(step) {
    if (!step || step.advanceOn == null || step.advanceOn === false) return null;
    const raw = step.advanceOn;
    if (typeof raw === 'string') return { event: raw, selector: null };
    if (typeof raw === 'object') return { event: String(raw.event || 'click'), selector: raw.selector || null };
    return null;
  }

  function waitCopy(step) {
    const mode = followMode(step);
    if (!mode) return '';
    if (mode.event === 'input' || mode.event === 'change') {
      return '請在亮起來的地方填寫。填好以後，導覽會自己往下走。';
    }
    if (mode.event === 'navigate') {
      return '請照著亮起來的地方做到下一頁。換頁以後，導覽會接下去。';
    }
    return '請直接按亮起來的地方。做完這一步，導覽會自己往下走。';
  }

  function ensureRoot() {
    if (root && root.isConnected) return root;
    if (!document.body) return null;
    root = document.createElement('div');
    root.id = 'order-tour';
    root.hidden = true;
    root.innerHTML =
      '<div class="tour-safe" aria-hidden="true"></div>' +
      '<div class="tour-spot is-off" aria-hidden="true"></div>' +
      '<div class="tour-tip" role="dialog" aria-modal="false" aria-labelledby="order-tour-title" aria-describedby="order-tour-text">' +
        '<span class="tour-arrow" hidden></span>' +
        '<div class="tour-body">' +
          '<div class="tour-steps" aria-hidden="true"></div>' +
          '<h2 class="tour-title" id="order-tour-title"></h2>' +
          '<p class="tour-text" id="order-tour-text"></p>' +
          '<p class="tour-wait" hidden></p>' +
          '<div class="tour-actions">' +
            '<span class="tour-side">' +
              '<button type="button" class="tour-pause">暫停</button>' +
              '<button type="button" class="tour-end">結束導覽</button>' +
            '</span>' +
            '<span class="tour-nav">' +
              '<button type="button" class="tour-prev">上一步</button>' +
              '<button type="button" class="tour-next">下一步</button>' +
            '</span>' +
          '</div>' +
        '</div>' +
      '</div>';
    holeSvg = document.createElementNS(SVG_NS, 'svg');
    holeSvg.setAttribute('class', 'tour-bg');
    holeSvg.setAttribute('aria-hidden', 'true');
    holePath = document.createElementNS(SVG_NS, 'path');
    holePath.setAttribute('fill-rule', 'evenodd');
    holeSvg.appendChild(holePath);
    root.insertBefore(holeSvg, root.firstChild);
    root.querySelector('.tour-pause').addEventListener('click', pauseTour);
    root.querySelector('.tour-end').addEventListener('click', function () {
      if (index >= steps.length - 1) finishTour();
      else pauseTour();
    });
    root.querySelector('.tour-prev').addEventListener('click', function () {
      if (!busy && index > 0) go(index - 1);
    });
    root.querySelector('.tour-next').addEventListener('click', function () {
      if (busy || index < 0) return;
      if (index >= steps.length - 1) finishTour();
      else go(index + 1);
    });
    document.body.appendChild(root);
    return root;
  }

  function paint() {
    const step = steps[index];
    const last = index === steps.length - 1;
    const dots = root.querySelector('.tour-steps');
    dots.replaceChildren();
    steps.forEach(function (_, i) {
      const dot = document.createElement('i');
      if (i <= index) dot.className = 'on';
      dots.appendChild(dot);
    });
    root.querySelector('.tour-title').textContent = (index + 1) + '／' + steps.length + '　' + step.title;
    root.querySelector('.tour-text').textContent = step.text;
    const wait = root.querySelector('.tour-wait');
    const hint = waitCopy(step);
    wait.hidden = !hint;
    wait.textContent = hint;
    const endBtn = root.querySelector('.tour-end');
    endBtn.textContent = last ? '關閉' : '結束導覽';
    endBtn.classList.toggle('is-primary', last);
    root.querySelector('.tour-next').hidden = last;
    root.querySelector('.tour-prev').disabled = index === 0;
    root.querySelector('.tour-tip').setAttribute('aria-label', '操作教學');
  }

  function querySelectorSafe(selector) {
    if (!selector) return null;
    if (selector.nodeType === 1) return selector.isConnected ? selector : null;
    try { return document.querySelector(selector); }
    catch (err) {
      console.warn('[order-tour] 選擇器無法使用：', selector);
      return null;
    }
  }

  function currentTarget() {
    const step = steps[index];
    if (!step || step.selector == null) return null;
    const el = querySelectorSafe(step.selector);
    if (!el) return null;
    const rect = el.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) return null;
    return el;
  }

  function followElement(step) {
    const mode = followMode(step);
    if (mode && mode.selector) return querySelectorSafe(mode.selector);
    return currentTarget();
  }

  function viewport() {
    const vv = window.visualViewport;
    if (!vv) return { left: 0, top: 0, width: window.innerWidth, height: window.innerHeight };
    return { left: vv.offsetLeft, top: vv.offsetTop, width: vv.width, height: vv.height };
  }

  function safeInsets() {
    const probe = root && root.querySelector('.tour-safe');
    if (!probe) return { top: 0, right: 0, bottom: 0, left: 0 };
    const style = getComputedStyle(probe);
    const n = function (value) { return parseFloat(value) || 0; };
    return {
      top: n(style.paddingTop),
      right: n(style.paddingRight),
      bottom: n(style.paddingBottom),
      left: n(style.paddingLeft),
    };
  }

  function reveal(el) {
    const rect = el.getBoundingClientRect();
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const inView = rect.top >= 12 && rect.left >= 12 && rect.bottom <= vh - 12 && rect.right <= vw - 12;
    if (inView) return true;
    const reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    el.style.scrollMarginTop = '88px';
    el.style.scrollMarginBottom = '120px';
    el.scrollIntoView({
      block: rect.height > vh * 0.7 ? 'nearest' : 'center',
      inline: rect.width < vw * 0.9 ? 'center' : 'nearest',
      behavior: reduce ? 'auto' : 'smooth',
    });
    return false;
  }

  function roundedRect(x, y, w, h, r) {
    const rr = Math.max(0, Math.min(r, w / 2, h / 2));
    return [
      'M', x + rr, y,
      'H', x + w - rr,
      'A', rr, rr, 0, 0, 1, x + w, y + rr,
      'V', y + h - rr,
      'A', rr, rr, 0, 0, 1, x + w - rr, y + h,
      'H', x + rr,
      'A', rr, rr, 0, 0, 1, x, y + h - rr,
      'V', y + rr,
      'A', rr, rr, 0, 0, 1, x + rr, y,
      'Z',
    ].join(' ');
  }

  function drawHole(rect) {
    if (!holeSvg || !holePath) return;
    const width = window.innerWidth;
    const height = window.innerHeight;
    holeSvg.setAttribute('viewBox', '0 0 ' + width + ' ' + height);
    holeSvg.style.width = width + 'px';
    holeSvg.style.height = height + 'px';
    const outer = 'M0 0 H' + width + ' V' + height + ' H0 Z';
    const spot = root.querySelector('.tour-spot');
    if (!rect) {
      holePath.setAttribute('d', outer);
      if (spot) spot.classList.add('is-off');
      return;
    }
    holePath.setAttribute('d', outer + ' ' + roundedRect(rect.x, rect.y, rect.w, rect.h, 12));
    if (!spot) return;
    spot.classList.remove('is-off');
    spot.style.top = rect.y + 'px';
    spot.style.left = rect.x + 'px';
    spot.style.width = rect.w + 'px';
    spot.style.height = rect.h + 'px';
  }

  function moveHole(target, instant) {
    if (holeFrame) cancelAnimationFrame(holeFrame);
    holeFrame = 0;
    const reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (instant || reduce || !holeNow || !target) {
      holeNow = target ? { x: target.x, y: target.y, w: target.w, h: target.h } : null;
      drawHole(holeNow);
      return;
    }
    const from = { x: holeNow.x, y: holeNow.y, w: holeNow.w, h: holeNow.h };
    const start = performance.now();
    const tick = function (now) {
      const t = Math.min(1, (now - start) / 280);
      const e = 1 - Math.pow(1 - t, 3);
      holeNow = {
        x: from.x + (target.x - from.x) * e,
        y: from.y + (target.y - from.y) * e,
        w: from.w + (target.w - from.w) * e,
        h: from.h + (target.h - from.h) * e,
      };
      drawHole(holeNow);
      holeFrame = t < 1 ? requestAnimationFrame(tick) : 0;
    };
    holeFrame = requestAnimationFrame(tick);
  }

  function measureHole(el) {
    if (!el) return null;
    const vv = viewport();
    const rect = el.getBoundingClientRect();
    const pad = 8;
    // getBoundingClientRect is already in the layout viewport's coordinate
    // system. Adding visualViewport.offsetTop/Left again displaces the focus
    // on zoomed mobile browsers (and when the software keyboard opens).
    let top = Math.max(vv.top + 4, rect.top - pad);
    let bot = Math.min(vv.top + vv.height - 4, rect.bottom + pad);
    let left = Math.max(vv.left + 4, rect.left - pad);
    let right = Math.min(vv.left + vv.width - 4, rect.right + pad);
    // A target inside an overflow-x table must not be highlighted outside
    // that table's visible scrollport.
    for (let parent = el.parentElement; parent && parent !== document.body; parent = parent.parentElement) {
      const style = window.getComputedStyle(parent);
      if (!/(auto|scroll|hidden|clip)/.test(style.overflowX + ' ' + style.overflowY)) continue;
      const bounds = parent.getBoundingClientRect();
      if (/(auto|scroll|hidden|clip)/.test(style.overflowY)) {
        top = Math.max(top, bounds.top);
        bot = Math.min(bot, bounds.bottom);
      }
      if (/(auto|scroll|hidden|clip)/.test(style.overflowX)) {
        left = Math.max(left, bounds.left);
        right = Math.min(right, bounds.right);
      }
    }
    const w = right - left;
    const h = bot - top;
    if (w < 8 || h < 8) return null;
    return { x: left, y: top, w: w, h: h };
  }

  function pointArrow(tip, hole) {
    const arrow = root.querySelector('.tour-arrow');
    if (!arrow) return;
    if (!hole) {
      arrow.hidden = true;
      return;
    }
    arrow.hidden = false;
    const tipTop = parseFloat(tip.style.top) || 0;
    const tipLeft = parseFloat(tip.style.left) || 0;
    const below = tipTop >= hole.y + hole.h - 8;
    arrow.classList.toggle('is-top', below);
    arrow.classList.toggle('is-bottom', !below);
    const center = hole.x + hole.w / 2;
    const width = tip.offsetWidth || 0;
    arrow.style.left = Math.max(18, Math.min(width - 36, center - tipLeft - 9)) + 'px';
  }

  function place(instant) {
    if (index < 0 || !root || root.hidden) return;
    const tip = root.querySelector('.tour-tip');
    if (!tip) return;
    const vv = viewport();
    const safe = safeInsets();
    const marginTop = 12 + safe.top;
    const marginBottom = 12 + safe.bottom;
    const marginLeft = 16 + safe.left;
    const marginRight = 16 + safe.right;
    const el = currentTarget();
    const hole = measureHole(el);
    const body = tip.querySelector('.tour-body');
    if (instant) tip.classList.add('is-instant');
    const limitBottom = vv.top + vv.height - marginBottom;
    const limitTop = vv.top + marginTop;
    if (!hole) {
      if (body) body.style.maxHeight = Math.max(180, limitBottom - limitTop) + 'px';
    } else {
      const gap = 16;
      const spaceBelow = limitBottom - (hole.y + hole.h + gap);
      const spaceAbove = hole.y - gap - limitTop;
      const available = Math.max(150, spaceBelow >= spaceAbove ? spaceBelow : spaceAbove);
      if (body) body.style.maxHeight = available + 'px';
    }
    const tw = tip.offsetWidth;
    const th = tip.offsetHeight;
    const smallScreen = vv.width <= 720;
    let tipTop;
    let tipLeft;
    if (smallScreen) {
      // Mobile: anchor the instruction panel to the opposite screen edge,
      // rather than trying to squeeze it beside a wide table cell.
      tipLeft = vv.left + Math.max(marginLeft, (vv.width - tw) / 2);
      tipTop = hole && hole.y > vv.top + vv.height / 2
        ? limitTop
        : Math.max(limitTop, limitBottom - th);
    } else if (!hole) {
      tipLeft = vv.left + Math.max(marginLeft, (vv.width - tw) / 2);
      tipTop = vv.top + Math.max(marginTop, (vv.height - th) / 2);
    } else {
      const gap = 16;
      const spaceBelow = limitBottom - (hole.y + hole.h + gap);
      const spaceAbove = hole.y - gap - limitTop;
      const below = spaceBelow >= spaceAbove;
      tipTop = below ? hole.y + hole.h + gap : Math.max(limitTop, hole.y - gap - th);
      const minLeft = vv.left + marginLeft;
      const maxLeft = vv.left + vv.width - tw - marginRight;
      tipLeft = Math.min(Math.max(minLeft, hole.x), Math.max(minLeft, maxLeft));
    }
    tip.style.top = tipTop + 'px';
    tip.style.left = tipLeft + 'px';
    pointArrow(tip, smallScreen ? null : hole);
    moveHole(hole, instant || !holeNow);
    if (instant) {
      void tip.offsetWidth;
      tip.classList.remove('is-instant');
    }
  }

  function onViewport() {
    if (index < 0 || placeQueued) return;
    placeQueued = true;
    requestAnimationFrame(function () {
      placeQueued = false;
      if (index >= 0) place(true);
    });
  }

  function typingTarget(node) {
    if (!node || !node.tagName) return false;
    if (/^(INPUT|TEXTAREA|SELECT)$/.test(node.tagName)) return true;
    return !!node.isContentEditable;
  }

  function onKey(event) {
    if (index < 0 && !busy) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      pauseTour();
      return;
    }
    if (busy || index < 0) return;
    if (event.altKey || event.ctrlKey || event.metaKey) return;
    if (typingTarget(document.activeElement)) return;
    if (event.key === 'ArrowRight') {
      event.preventDefault();
      if (index < steps.length - 1) go(index + 1);
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault();
      if (index > 0) go(index - 1);
    }
  }

  function eventInside(el, node) {
    return !!(el && node && (el === node || el.contains(node)));
  }

  function onFollowClick(event) {
    if (index < 0 || busy) return;
    if (root && root.contains(event.target)) return;
    const step = steps[index];
    const mode = followMode(step);
    if (!mode || mode.event !== 'click') return;
    const el = followElement(step);
    if (!eventInside(el, event.target)) return;
    const next = steps[index + 1];
    if (next && next.path && !stepMatches(next)) return;
    go(index + 1);
  }

  function onFollowInput(event) {
    if (index < 0 || busy) return;
    const mode = followMode(steps[index]);
    if (!mode || mode.event !== 'input') return;
    const el = followElement(steps[index]);
    if (!eventInside(el, event.target)) return;
    const field = event.target;
    const stepIndex = index;
    window.clearTimeout(inputTimer);
    inputTimer = window.setTimeout(function () {
      if (index !== stepIndex) return;
      if (field.value && String(field.value).trim()) go(stepIndex + 1);
    }, 450);
  }

  function onFollowChange(event) {
    if (index < 0 || busy) return;
    const mode = followMode(steps[index]);
    if (!mode || mode.event !== 'change') return;
    const el = followElement(steps[index]);
    if (!eventInside(el, event.target)) return;
    const field = event.target;
    if (field && (field.type === 'checkbox' || field.type === 'radio' || field.tagName === 'SELECT')) {
      go(index + 1);
      return;
    }
    if (field && field.value && String(field.value).trim()) go(index + 1);
  }

  function bind() {
    if (listening) return;
    listening = true;
    window.addEventListener('resize', onViewport);
    window.addEventListener('scroll', onViewport, true);
    document.addEventListener('keydown', onKey);
    document.addEventListener('click', onFollowClick, false);
    document.addEventListener('input', onFollowInput, true);
    document.addEventListener('change', onFollowChange, true);
    if (window.visualViewport) {
      window.visualViewport.addEventListener('resize', onViewport);
      window.visualViewport.addEventListener('scroll', onViewport);
    }
  }

  function unbind() {
    if (!listening) return;
    listening = false;
    window.removeEventListener('resize', onViewport);
    window.removeEventListener('scroll', onViewport, true);
    document.removeEventListener('keydown', onKey);
    document.removeEventListener('click', onFollowClick, false);
    document.removeEventListener('input', onFollowInput, true);
    document.removeEventListener('change', onFollowChange, true);
    if (window.visualViewport) {
      window.visualViewport.removeEventListener('resize', onViewport);
      window.visualViewport.removeEventListener('scroll', onViewport);
    }
    window.clearTimeout(inputTimer);
  }

  function dismissOpenBanners() {
    document.querySelectorAll('.order-tour-banner').forEach(function (node) {
      const key = node.getAttribute('data-storage-key');
      if (key) markSeen(key);
      node.remove();
    });
  }

  function hideResume() {
    if (resumeBtn) resumeBtn.hidden = true;
  }

  function showResume() {
    if (!document.body) return;
    if (!resumeBtn || !resumeBtn.isConnected) {
      resumeBtn = document.createElement('button');
      resumeBtn.type = 'button';
      resumeBtn.className = 'order-tour-resume';
      resumeBtn.textContent = '繼續導覽';
      resumeBtn.addEventListener('click', resumeTour);
      document.body.appendChild(resumeBtn);
    }
    resumeBtn.hidden = false;
  }

  async function show(to) {
    if (to < 0 || to >= steps.length) return;
    const token = ++run;
    busy = true;
    index = to;
    const step = steps[to];
    try {
      if (step.before) await step.before(step, to);
    } catch (err) {
      console.warn('[order-tour] before() 失敗', err);
    }
    if (token !== run) return;
    index = to;
    busy = false;
    // Optional controls (such as 'copy to other schools' when there is only
    // one school) should not lead to a disconnected, centered pointer.
    if (step.optional && !currentTarget()) {
      go(to + 1);
      return;
    }
    saveProgress('active', to);
    hideResume();
    if (!ensureRoot()) return;
    root.hidden = false;
    paint();
    bind();
    const el = currentTarget();
    if (!el && step.selector) console.warn('[order-tour] 找不到步驟目標：', step.selector);
    const alreadyVisible = el ? reveal(el) : true;
    const snap = !alreadyVisible || snapNext;
    snapNext = false;
    place(snap);
    requestAnimationFrame(function () {
      if (token !== run || index !== to) return;
      place(snap);
      const title = root.querySelector('.tour-title');
      if (title) {
        title.setAttribute('tabindex', '-1');
        title.focus({ preventScroll: true });
      }
    });
    if (document.fonts && document.fonts.ready) {
      document.fonts.ready.then(function () {
        if (token === run && index === to) place(true);
      });
    }
  }

  function teardown() {
    run += 1;
    index = -1;
    busy = false;
    unbind();
    if (holeFrame) cancelAnimationFrame(holeFrame);
    holeFrame = 0;
    if (root) root.hidden = true;
    const focus = returnFocus;
    returnFocus = null;
    if (focus && focus.isConnected && (!root || !root.contains(focus))) {
      try { focus.focus({ preventScroll: true }); } catch (err) { /* 元素可能不接受焦點。 */ }
    }
  }

  function go(to) {
    if (to < 0) return;
    if (!steps.length) steps = catalog.slice();
    if (to >= steps.length) {
      finishTour();
      return;
    }
    const step = steps[to];
    saveProgress('active', to);
    if (step.path && !stepMatches(step)) {
      hideResume();
      teardown();
      location.assign(step.path);
      return;
    }
    show(to);
  }

  function pauseTour() {
    if (index < 0 && !busy) return;
    const i = index >= 0 ? index : 0;
    saveProgress('paused', i);
    teardown();
    showResume();
  }

  function finishTour() {
    const i = index >= 0 ? index : 0;
    saveProgress('done', i);
    const callback = onEnd;
    onEnd = null;
    teardown();
    hideResume();
    if (callback) callback();
  }

  function endTour() {
    pauseTour();
  }

  function resumeTour() {
    if (!steps.length) steps = catalog.slice();
    const saved = loadProgress();
    let i = 0;
    if (saved && saved.index >= 0 && saved.index < steps.length && saved.status !== 'done') i = saved.index;
    hideResume();
    snapNext = true;
    go(i);
  }

  function restore() {
    if (index >= 0) return;
    steps = catalog.slice();
    const saved = loadProgress();
    if (!saved || saved.status === 'done' || saved.index < 0 || saved.index >= steps.length) {
      hideResume();
      return;
    }
    if (saved.status === 'paused') {
      showResume();
      return;
    }
    if (saved.status !== 'active') return;
    if (stepMatches(steps[saved.index])) {
      snapNext = true;
      show(saved.index);
      return;
    }
    for (let j = saved.index + 1; j < steps.length; j += 1) {
      if (stepMatches(steps[j])) {
        saveProgress('active', j);
        snapNext = true;
        show(j);
        return;
      }
    }
    showResume();
  }

  function startTour(input, options) {
    const list = normalize(input);
    if (!list.length) return false;
    const opts = options || {};
    if (opts.progressKey) progressKey = opts.progressKey;
    catalog = list;
    steps = list;
    onEnd = typeof opts.onEnd === 'function' ? opts.onEnd : null;
    if (opts.storageKey) markSeen(opts.storageKey);
    dismissOpenBanners();
    const active = document.activeElement;
    teardown();
    returnFocus = active && active !== document.body ? active : null;
    snapNext = true;
    const startAt = Number.isInteger(opts.index) ? opts.index : 0;
    go(Math.max(0, Math.min(list.length - 1, startAt)));
    return true;
  }

  function attach(input, options) {
    const list = normalize(input);
    if (!list.length) return;
    const opts = options || {};
    if (opts.progressKey) progressKey = opts.progressKey;
    catalog = list;
    if (index >= 0) return;
    restore();
  }

  function mountTarget(target) {
    if (!target) return null;
    if (typeof target === 'string') return document.querySelector(target);
    return target;
  }

  function mountHelpButton(target, options) {
    const opts = options || {};
    const parent = mountTarget(target);
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'order-tour-help' + (opts.className ? ' ' + opts.className : '');
    button.textContent = opts.label || '操作教學';
    button.addEventListener('click', function () {
      if (opts.progressKey) progressKey = opts.progressKey;
      const list = normalize(opts.steps || catalog);
      if (list.length) {
        catalog = list;
        steps = list;
      }
      const saved = loadProgress();
      if (saved && (saved.status === 'paused' || saved.status === 'active')) resumeTour();
      else startTour(list, opts);
    });
    if (parent) {
      const existing = parent.querySelector(':scope > .order-tour-help');
      if (existing) existing.remove();
      parent.appendChild(button);
    }
    return button;
  }

  function mountFirstVisitBanner(target, options) {
    const opts = options || {};
    const key = opts.storageKey || DEFAULT_SEEN_KEY;
    const parent = mountTarget(target);
    if (!parent) return null;
    const existing = parent.querySelector(':scope > .order-tour-banner');
    if (existing) existing.remove();
    if (seen(key)) return null;

    const banner = document.createElement('div');
    banner.className = 'order-tour-banner';
    banner.setAttribute('data-storage-key', key);
    banner.setAttribute('role', 'region');
    banner.setAttribute('aria-label', '操作教學');
    const text = document.createElement('p');
    const title = document.createElement('b');
    title.textContent = opts.title || '第一次使用？';
    text.appendChild(title);
    text.appendChild(document.createTextNode(
      ' ' + (opts.text || '跟著操作教學走一遍。亮起來的地方可以直接操作，做完會自動到下一步。')
    ));
    const actions = document.createElement('div');
    actions.className = 'order-tour-banner-actions';
    const start = document.createElement('button');
    start.type = 'button';
    start.className = 'order-tour-banner-start';
    start.textContent = opts.startLabel || '開始導覽';
    const dismiss = document.createElement('button');
    dismiss.type = 'button';
    dismiss.className = 'order-tour-banner-dismiss';
    dismiss.textContent = opts.dismissLabel || '先不用';
    start.addEventListener('click', function () {
      markSeen(key);
      banner.remove();
      startTour(opts.steps || [], opts);
    });
    dismiss.addEventListener('click', function () {
      markSeen(key);
      banner.remove();
    });
    actions.append(start, dismiss);
    banner.append(text, actions);
    parent.appendChild(banner);
    return banner;
  }

  function boot() {
    if (window.ORDER_FLOW_TOUR_STEPS) attach(window.ORDER_FLOW_TOUR_STEPS);
  }

  window.OrderTour = {
    startTour: startTour,
    endTour: endTour,
    pauseTour: pauseTour,
    resumeTour: resumeTour,
    finishTour: finishTour,
    attach: attach,
    isActive: function () { return index >= 0; },
    mountHelpButton: mountHelpButton,
    mountFirstVisitBanner: mountFirstVisitBanner,
  };
  window.startTour = startTour;

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
