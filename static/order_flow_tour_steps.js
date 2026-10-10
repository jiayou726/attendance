/* 叫貨系統操作教學。
   主流程（總覽接上 總表 → 各校選菜 → 菜色用量表 → 採購叫貨）放在
   window.ORDER_FLOW_TOUR_STEPS，引擎換頁後會從停下的步驟接下去。
   其他頁用自己的進度，避免蓋掉主流程。右上「操作教學」從這一頁的第一步開始。 */
(function () {
  "use strict";

  var ROOT = "/admin/order-tool";
  var FLOW_KEY = "order-tool-tour-progress";

  function step(path, exact, selector, title, text, advanceOn) {
    var row = { selector: selector, title: title, text: text };
    if (path) {
      row.path = path;
      row.exact = exact !== false;
    }
    if (advanceOn) row.advanceOn = advanceOn;
    return row;
  }

  var HOME = [
    step(ROOT, true, "a.home-primary", "從這開始", "每週叫貨都從「開啟總表」開始。"),
    step(ROOT, true, "#home-functions-title", "其他功能", "菜色、食材、學校、廠商，點卡片就進去。"),
    step(ROOT, true, "#purchase-week-detail", "看整週", "按這裡看整週要叫的東西。"),
    step(ROOT, true, "a.btn.secondary[href*='/purchases']", "查舊單", "以前叫過什麼，按「採購歷史」。"),
    step(ROOT, true, "a.home-primary", "開始排菜", "按「開啟總表」，一路帶你做到叫貨。", "click"),
  ];

  var SUMMARY = [
    step(ROOT + "/summary", true, ".week-switcher", "找對週", "按上一週或下一週，找到要叫貨的那週。"),
    step(ROOT + "/summary", true, ".menu-week-grid .quick-dish-form", "加菜", "打菜名、選分類，按「＋ 新增」。", { event: "click", selector: ".menu-week-grid .quick-dish-form button.quick-add" }),
    step(ROOT + "/summary", true, "summary.menu-import-summary", "Excel 更快", "點開選「選擇 Excel 菜單」，按「匯入總表」。"),
    step(ROOT + "/summary", true, "a.summary-next", "分給學校", "菜排好，按「下一步：選擇學校」。", "click"),
  ];

  var SCHOOLS = [
    step(ROOT + "/summary/schools", true, "#school-menu-picker", "選學校", "從清單選要處理的那一間。", "change"),
    step(ROOT + "/summary/schools", true, ".school-week-grid .school-dish-check", "勾菜", "勾這間學校這天吃的菜，自動儲存。", "change"),
    step(ROOT + "/summary/schools", true, ".school-week-grid label.headcount-box", "填人數", "填這天吃幾人；有素食也填「素食人數」。", { event: "input", selector: ".school-week-grid label.headcount-box input" }),
    step(ROOT + "/summary/schools", true, ".school-no-service-toggle", "停餐", "這天沒供餐就勾，不會算進叫貨。"),
    step(ROOT + "/summary/schools", true, ".school-copy-day-btn", "一樣就複製", "勾其他學校，按「確認複製」。"),
    step(ROOT + "/summary/schools", true, "#school-usage-date", "選日子", "選剛勾好菜的那一天。", "change"),
    step(ROOT + "/summary/schools", true, "#school-usage-next", "看用量", "按這裡，看每道菜要用多少。", "click"),
  ];

  var USAGE = [
    step(ROOT + "/summary/production-sheet", true, ".production-variant-switch", "葷素切換", "按「葷食」或「素食」分開看。"),
    step(ROOT + "/summary/production-sheet", true, ".production-actual, .production-empty", "核對用量", "數量不對直接改，自動儲存。"),
    step(ROOT + "/summary/production-sheet", true, "#usage-order-next", "去叫貨", "按這裡，採購單照這些數量算。", "click"),
  ];

  var PROCUREMENT = [
    step(ROOT + "/summary/procurement", true, ".actual-qty-input, .procurement-empty", "叫多少", "填要跟廠商叫的數量。", "input"),
    step(ROOT + "/summary/procurement", true, ".supplier-search-input, .procurement-empty", "向誰叫", "選廠商；沒有就直接打名字。", "input"),
    step(ROOT + "/summary/procurement", true, "td .delivery-fields, .procurement-empty", "幾時送", "選送貨日期和上午／下午。"),
    step(ROOT + "/summary/procurement", true, "details.manual-procurement-card", "臨時加買", "菜單外的東西在這填，按「新增臨時食材」。"),
    step(ROOT + "/summary/procurement", true, ".procurement-table .order-check, .procurement-empty", "叫完打勾", "跟廠商講好就勾「已叫」。", { event: "change", selector: ".procurement-table .order-confirm-toggle" }),
    step(ROOT + "/summary/procurement", true, "a.btn.warn[href*='procurement']", "要列印", "按「匯出 Excel」下載這天的採購單。"),
    step(ROOT + "/summary/procurement", true, "nav[aria-label='切換採購日期'] a:last-child", "一天完成！", "按「後一天」，這週每天都照做一次。"),
    step(ROOT + "/summary/procurement", true, "nav[aria-label='團膳管理主選單'] > a", "最後看整週", "整週叫完，到「總覽」按「查看週採購單」。"),
  ];

  var FLOW = HOME.concat(SUMMARY, SCHOOLS, USAGE, PROCUREMENT);
  window.ORDER_FLOW_TOUR_STEPS = FLOW;

  var PAGES = [
    { id: "home", flow: true, index: 0, match: function (p) { return p === ROOT; } },
    { id: "schools-flow", flow: true, index: HOME.length + SUMMARY.length, match: function (p) { return p === ROOT + "/summary/schools"; } },
    { id: "usage", flow: true, index: HOME.length + SUMMARY.length + SCHOOLS.length, match: function (p) { return p === ROOT + "/summary/production-sheet"; } },
    { id: "procurement", flow: true, index: HOME.length + SUMMARY.length + SCHOOLS.length + USAGE.length, match: function (p) { return p === ROOT + "/summary/procurement"; } },
    { id: "summary", flow: true, index: HOME.length, match: function (p) { return p === ROOT + "/summary"; } },
    page("recipes", ROOT + "/recipes", [
      step(ROOT + "/recipes", true, "section.card.span-5 form.main-form .field:has(input[name='name'])", "新增菜", "打菜名、選分類。"),
      step(ROOT + "/recipes", true, "section.card.span-5 form.main-form .field:has(input[name='serving_output_g'])", "每人份量", "填每人打幾克，不知道先空白。"),
      step(ROOT + "/recipes", true, "section.card.span-5 form.main-form > button.btn", "建立", "按這裡，接著加食材。"),
      step(ROOT + "/recipes", true, "section.card.span-7 .top", "找菜", "打菜名按「搜尋」，或選類別。"),
      step(ROOT + "/recipes", true, "table .actions a[href*='/recipes/']", "改食材", "按「配方」改用哪些食材、各多少。"),
      step(ROOT + "/recipes", true, "table .actions", "改名或刪除", "「編輯」改名稱分類；「刪除」會移除，小心。"),
    ]),
    {
      id: "recipe",
      steps: [
        step(null, false, "section.card.span-8", "食材清單", "這道菜用的食材和「每人 AP」。"),
        step(null, false, "section.card.span-8 tbody tr, section.card.span-8 .empty", "改用量", "改數字按「更新」；不要就按「移除」。"),
        step(null, false, "section.card.span-4 form.main-form, section.card.span-4 .notice", "加食材", "搜尋食材、填用量，按「加入配方」。"),
        step(null, false, "a.btn.ok", "改完回去", "按「回總表」繼續排菜。"),
      ],
      match: function (p) { return /^\/admin\/order-tool\/recipes\/\d+$/.test(p); },
    },
    page("ingredients", ROOT + "/ingredients", [
      step(ROOT + "/ingredients", true, "section.card.span-5 form.main-form .field:has(input[name='name'])", "新增食材", "打名稱，選平常叫的廠商。"),
      step(ROOT + "/ingredients", true, "section.card.span-5 form.main-form .field:has(input[name='grams_per_purchase_unit'])", "換算", "選單位，填 1 單位幾克，例如 1 kg 填 1000。"),
      step(ROOT + "/ingredients", true, "section.card.span-5 form.main-form .field:has(input[name='order_increment'])", "最少叫多少", "填廠商最少量，系統會往上湊。"),
      step(ROOT + "/ingredients", true, "section.card.span-5 form.main-form > button.btn", "存起來", "按「新增」。"),
      step(ROOT + "/ingredients", true, "section.card.span-7 table .actions", "修改", "價格變了按「編輯」；不用了按「停用」。"),
    ]),
    page("schools", ROOT + "/schools", [
      step(ROOT + "/schools", true, "section.card.span-5 form.main-form .field:has(input[name='name'])", "新增學校", "打學校名稱，代碼可空白。"),
      step(ROOT + "/schools", true, "section.card.span-5 form.main-form .field:has(input[name='default_headcount'])", "平常人數", "選菜時會自動帶入。"),
      step(ROOT + "/schools", true, "section.card.span-5 form.main-form > button.btn", "存起來", "按「新增」。"),
      step(ROOT + "/schools", true, "section.card.span-7 table .actions", "修改", "人數變了按「編輯」；不供餐按「停用」。"),
    ]),
    page("suppliers", ROOT + "/suppliers", [
      step(ROOT + "/suppliers", true, "section.card.span-5 form.main-form", "新增廠商", "填好按「新增」。"),
      step(ROOT + "/suppliers", true, "section.card.span-7 table tbody a", "看詳細", "點名字看叫過的食材。"),
      step(ROOT + "/suppliers", true, "section.card.span-7 table .actions", "修改", "電話改了按「編輯」；不合作按「停用」。"),
    ]),
    {
      id: "supplier",
      steps: [
        step(null, false, "section.card.span-6", "聯絡方式", "電話、聯絡人都在這。"),
        step(null, false, "a.btn[href*='edit=']", "改資料", "按「編輯資料」。"),
        step(null, false, "section.card.span-6 .top", "叫過什麼", "打食材名按「搜尋」。"),
        step(null, false, "form.supplier-item-add", "加品項", "填好按「新增品項」。"),
        step(null, false, "a.btn.secondary[href$='/suppliers']", "回清單", "按這裡回去。"),
      ],
      match: function (p) { return /^\/admin\/order-tool\/suppliers\/\d+$/.test(p); },
    },
    page("daily", ROOT + "/summary/daily-kitchen-sheet", [
      step(ROOT + "/summary/daily-kitchen-sheet", true, ".daily-kitchen-controls .production-date-form", "選日期", "選好按「查看」，或按前一天／後一天。"),
      step(ROOT + "/summary/daily-kitchen-sheet", true, "textarea[data-daily-save-field], .daily-kitchen-section .empty", "補備註", "例如「白米（20K）」，自動儲存。"),
      step(ROOT + "/summary/daily-kitchen-sheet", true, "input[data-daily-class-count], .daily-kitchen-section .empty", "核對班級", "確認各校班級數，合菜、便當自動算。"),
      step(ROOT + "/summary/daily-kitchen-sheet", true, "tr[data-daily-count-row], .daily-kitchen-section .empty", "核對份數", "數字不對直接改。"),
      step(ROOT + "/summary/daily-kitchen-sheet", true, "a[href*='daily-kitchen-sheet.xlsx']", "給廚房", "按「匯出 Excel」印出來。"),
    ]),
    page("ai", ROOT + "/ai-menu", [
      step(ROOT + "/ai-menu", true, "form[action*='ai-menu/generate'] .field:has(input[name='name'])", "基本資料", "取名稱，選給哪個年齡層。"),
      step(ROOT + "/ai-menu", true, "form[action*='ai-menu/generate'] .row", "排哪幾天", "選日期；週末也排就勾「週六、週日也排菜」。"),
      step(ROOT + "/ai-menu", true, "form[action*='ai-menu/generate'] .ai-structure-grid", "每天幾道", "填主食、主菜、青菜等各幾道。"),
      step(ROOT + "/ai-menu", true, "form[action*='ai-menu/generate'] .ai-structure-grid:has(input[name='recipe_repeat_days'])", "設規則", "多久不重複、魚至少幾次、炸物最多幾次。"),
      step(ROOT + "/ai-menu", true, "form[action*='ai-menu/generate'] button[type='submit']", "產生", "按這裡，AI 排好草稿。"),
      step(ROOT + "/ai-menu", true, "a[href*='recipe-tags']", "先標好菜", "魚、炸物要先在「菜色標記」勾，規則才準。"),
      step(ROOT + "/ai-menu", true, "section.card.span-12", "找草稿", "之前的草稿在這，點開繼續改。"),
    ]),
    {
      id: "ai-draft",
      steps: [
        step(null, false, "table.ai-public-table", "看草稿", "每天排好的菜都在這。"),
        step(null, false, "table.ai-public-table a[href*='/swap']", "不喜歡就換", "按那道菜旁的「換菜」挑別道。"),
        step(null, false, ".actions:has(form[action*='/regenerate'])", "重排", "只重排沒鎖的，或整份重來。"),
        step(null, false, "a[href*='/ai-menu/drafts/']:not([href*='/swap'])", "排素食", "葷食好了按這裡排素食。"),
        step(null, false, "#public-excel-download", "下載", "按這裡下載公版菜單。"),
        step(null, false, "section.card .top a.btn.secondary", "回去", "要改條件按「回排菜設定」。"),
      ],
      match: function (p) { return /^\/admin\/order-tool\/ai-menu\/drafts\/\d+$/.test(p); },
    },
    page("tags", ROOT + "/recipe-tags", [
      step(ROOT + "/recipe-tags", true, "form.card label.pill", "勾標記", "魚類、炸物、甜湯等，照實際勾。"),
      step(ROOT + "/recipe-tags", true, "h1 + .muted", "參考建議", "「建議」只看菜名猜的，請自己確認。"),
      step(ROOT + "/recipe-tags", true, "form.card button.btn", "存起來", "勾好按「儲存這道菜」。"),
      step(ROOT + "/recipe-tags", true, "a.btn.secondary[href*='ai-menu']", "回去", "按「回 AI 菜單」排菜。"),
    ]),
    page("catalog", ROOT + "/catalog-import", [
      step(ROOT + "/catalog-import", true, "a[href*='catalog-template']", "先拿格式", "按這裡下載範本，照格式填。"),
      step(ROOT + "/catalog-import", true, "input[type='file'][name='file']", "上傳", "選你填好的檔案。", "click"),
      step(ROOT + "/catalog-import", true, "form.main-form button.btn", "比對", "按這裡，先看會新增哪些，還不會存。", "click"),
      step(ROOT + "/catalog-import", true, "section.card.span-7 h2", "確認", "看清楚沒問題，再按下方確認匯入。"),
    ]),
    page("weekly", ROOT + "/weekly-procurement", [
      step(ROOT + "/weekly-procurement", true, ".week-nav", "選週", "或用「選擇週次」按「切換」。"),
      step(ROOT + "/weekly-procurement", true, "section.vendor, section.card .empty", "依廠商分好", "每家廠商每天要送什麼一目了然。"),
      step(ROOT + "/weekly-procurement", true, ".weekly-order-check, section.card .empty", "補勾", "叫好的在這也能勾「已叫」。"),
      step(ROOT + "/weekly-procurement", true, "a.btn.export", "給廠商", "按這裡下載依廠商分好的 Excel。"),
    ]),
    page("purchases", ROOT + "/purchases", [
      step(ROOT + "/purchases", true, "form[method='get']", "找舊單", "選日期範圍按「篩選」。"),
      step(ROOT + "/purchases", true, "label.order-check, .empty", "整張叫完", "整天都叫好，勾「全叫」。"),
      step(ROOT + "/purchases", true, "table a.btn.secondary", "看細項", "按這裡看那天的採購單。"),
      step(ROOT + "/purchases", true, "form.main-form button.btn", "一次多天", "選開始、結束日期，按「產生 / 更新草稿」。"),
    ]),
    {
      id: "purchase",
      steps: [
        step(null, false, "a.btn[href*='/summary/procurement']", "要改數量", "按這裡回採購叫貨頁改。"),
        step(null, false, "table .order-check", "逐項打勾", "叫好一項勾一項。"),
        step(null, false, "form.main-form", "寫備註", "打好按「儲存備註」。"),
        step(null, false, "button.btn.ok", "最後確認", "全部對了按這裡，確認後就不能改。"),
      ],
      match: function (p) { return /^\/admin\/order-tool\/purchases\/\d+$/.test(p); },
    },
  ];

  function page(id, path, steps) {
    return {
      id: id,
      steps: steps,
      match: function (p) { return p === path; },
    };
  }

  function here() {
    var path = location.pathname || "/";
    if (path.length > 1) path = path.replace(/\/+$/, "");
    return path;
  }

  function currentPage() {
    var path = here();
    for (var i = 0; i < PAGES.length; i += 1) {
      if (PAGES[i].match(path)) return PAGES[i];
    }
    return null;
  }

  function readProgress(key) {
    try {
      var raw = localStorage.getItem(key);
      if (!raw) return null;
      var data = JSON.parse(raw);
      if (!data || typeof data.index !== "number") return null;
      if (data.status !== "active" && data.status !== "paused") return null;
      return data;
    } catch (err) {
      return null;
    }
  }

  function startHere() {
    var found = currentPage();
    if (!found || !window.OrderTour) return;
    if (found.flow) {
      window.OrderTour.startTour(FLOW, { index: found.index, progressKey: FLOW_KEY });
      return;
    }
    window.OrderTour.startTour(found.steps, { index: 0, progressKey: "order-tool-tour-" + found.id });
  }

  function mountHelp() {
    var parent = document.querySelector("main.wrap > .top");
    if (!parent) return;
    var existing = parent.querySelector(":scope > .order-tour-help");
    if (existing) existing.remove();
    var button = document.createElement("button");
    button.type = "button";
    button.className = "order-tour-help";
    button.textContent = "操作教學";
    button.addEventListener("click", startHere);
    parent.appendChild(button);
  }

  function mountHomeBanner() {
    if (here() !== ROOT || !window.OrderTour) return;
    var banner = window.OrderTour.mountFirstVisitBanner(document.querySelector("main.wrap"), {
      steps: FLOW,
      progressKey: FLOW_KEY,
      text: "跟著做一遍：選菜、看菜色用量表，再去採購叫貨。亮起來的地方可以直接按。",
    });
    var top = document.querySelector("main.wrap > .top");
    if (banner && top && top.parentElement) top.parentElement.insertBefore(banner, top.nextSibling);
  }

  function restorePageTour() {
    if (!window.OrderTour || window.OrderTour.isActive()) return;
    if (readProgress(FLOW_KEY)) return;
    var found = currentPage();
    if (!found || found.flow) return;
    var key = "order-tool-tour-" + found.id;
    if (!readProgress(key)) return;
    window.OrderTour.attach(found.steps, { progressKey: key });
  }

  function setup() {
    if (!window.OrderTour) return;
    mountHelp();
    mountHomeBanner();
    restorePageTour();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", setup);
  else setup();
})();
