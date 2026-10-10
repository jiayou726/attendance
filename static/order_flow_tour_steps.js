/* 叫貨流程的操作教學步驟。引擎在 DOMContentLoaded 時讀
   window.ORDER_FLOW_TOUR_STEPS 並 attach。
   advanceOn / waitFor：click、input、change。換頁的按鈕用 click，
   下一步若在別的網址，引擎會等真正換頁後再接上。 */
window.ORDER_FLOW_TOUR_STEPS = [
  {
    path: "/admin/order-tool",
    exact: true,
    selector: "a.home-primary",
    title: "先開總表",
    text: "按「開啟總表」。叫貨要先排好這一週的菜。",
    advanceOn: "click",
  },
  {
    path: "/admin/order-tool/summary",
    exact: true,
    selector: ".week-switcher",
    title: "對到這一週",
    text: "按「上一週」或「下一週」，找到要叫貨的那一週。",
    advanceOn: "click",
  },
  {
    path: "/admin/order-tool/summary",
    exact: true,
    selector: ".menu-week-grid .quick-dish-form",
    title: "把菜加進去",
    text: "輸入菜名，選好分類，再按「＋ 新增」。",
    advanceOn: { event: "click", selector: ".menu-week-grid .quick-dish-form button.quick-add" },
  },
  {
    path: "/admin/order-tool/summary",
    exact: true,
    selector: "a.summary-next",
    title: "去選學校",
    text: "菜加好後，按「下一步：選擇學校」。",
    advanceOn: "click",
  },
  {
    path: "/admin/order-tool/summary/schools",
    exact: true,
    selector: "#school-menu-picker",
    title: "先選一間學校",
    text: "點「先選擇學校」，選要處理的那一間。",
    advanceOn: "change",
  },
  {
    path: "/admin/order-tool/summary/schools",
    exact: true,
    selector: ".school-week-grid .school-dish-check",
    title: "勾這天的菜",
    text: "勾這間學校要吃的菜。勾完會自己儲存。",
    advanceOn: "change",
  },
  {
    path: "/admin/order-tool/summary/schools",
    exact: true,
    selector: ".school-week-grid .headcount-box input",
    title: "填人數",
    text: "在「葷食人數」填這天要吃的人數。",
    advanceOn: "input",
  },
  {
    path: "/admin/order-tool/summary/schools",
    exact: true,
    selector: "#school-usage-date",
    title: "選要看的日期",
    text: "選你剛勾完菜的那一天。",
    advanceOn: "change",
  },
  {
    path: "/admin/order-tool/summary/schools",
    exact: true,
    selector: "#school-usage-next",
    title: "去看用量表",
    text: "按「下一步：菜色用量表」，看每道菜要用多少。",
    advanceOn: "click",
  },
  {
    path: "/admin/order-tool/summary/production-sheet",
    exact: true,
    selector: ".production-dish-card, .production-empty",
    title: "核對每道菜的用量",
    text: "看「本菜預估採購量」。看完請點一下這張表。",
    advanceOn: "click",
  },
  {
    path: "/admin/order-tool/summary/production-sheet",
    exact: true,
    selector: "#usage-order-next",
    title: "去採購叫貨",
    text: "按「下一步：採購叫貨」。採購單會用這些數量來算。",
    advanceOn: "click",
  },
  {
    path: "/admin/order-tool/summary/procurement",
    exact: true,
    selector: ".actual-qty-input, .procurement-empty",
    title: "填這次要叫的量",
    text: "在「實際採購量」填要跟廠商叫的數量。",
    advanceOn: "input",
  },
  {
    path: "/admin/order-tool/summary/procurement",
    exact: true,
    selector: ".supplier-search-input, .procurement-empty",
    title: "填供應廠商",
    text: "在「供應廠商」選要向誰叫。沒有的話直接打名字。",
    advanceOn: "input",
  },
  {
    path: "/admin/order-tool/summary/procurement",
    exact: true,
    selector: ".order-confirm-toggle",
    title: "叫完就打勾",
    text: "跟廠商講完後，勾「已叫」。",
    advanceOn: "change",
  },
];

document.addEventListener("DOMContentLoaded", function () {
  if (!window.OrderTour) return;
  var steps = window.ORDER_FLOW_TOUR_STEPS;
  OrderTour.mountHelpButton(".top", { steps: steps, label: "操作教學" });
  var banner = OrderTour.mountFirstVisitBanner(".wrap", {
    steps: steps,
    text: "跟著做一遍：選菜、看菜色用量表，再去採購叫貨。亮起來的地方可以直接按。",
  });
  var top = document.querySelector(".top");
  if (banner && top && top.parentElement) top.parentElement.insertBefore(banner, top.nextSibling);
});
