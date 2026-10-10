/* Order-tool guided-tour steps for the reordered flow.
   Plug into the tour engine as window.ORDER_FLOW_TOUR_STEPS.
   path: pathname to open (no trailing slash, except the home page).
   exact: when true, match only that path. Otherwise the step also
   matches nested paths. selector: the one control this step explains.
   The engine should spotlight that element. */
window.ORDER_FLOW_TOUR_STEPS = [
  {
    path: "/admin/order-tool",
    exact: true,
    selector: "a.home-primary",
    title: "先開總表",
    text: "按「開啟總表」。叫貨要先排好這一週的菜。",
  },
  {
    path: "/admin/order-tool/summary",
    exact: true,
    selector: ".week-switcher",
    title: "對到這一週",
    text: "按「上一週」或「下一週」，找到要叫貨的那一週。",
  },
  {
    path: "/admin/order-tool/summary",
    exact: true,
    selector: ".menu-week-grid .quick-dish-form",
    title: "把菜加進去",
    text: "在每一天輸入菜名，選好分類，再按「＋ 新增」。",
  },
  {
    path: "/admin/order-tool/summary",
    exact: true,
    selector: "a.summary-next",
    title: "去選學校",
    text: "菜加好後，按「下一步：選擇學校」。",
  },
  {
    path: "/admin/order-tool/summary/schools",
    exact: true,
    selector: "#school-menu-picker",
    title: "先選一間學校",
    text: "點「先選擇學校」，選要處理的那一間。下面會列出這間學校的菜。",
  },
  {
    path: "/admin/order-tool/summary/schools",
    exact: true,
    selector: ".school-week-grid .school-dish-check",
    title: "勾菜、填人數",
    text: "勾這間學校要吃的菜，並填人數。改完會自己儲存。",
  },
  {
    path: "/admin/order-tool/summary/schools",
    exact: true,
    selector: "#school-usage-next",
    title: "去看用量表",
    text: "菜勾好後，選日期，再按「下一步：菜色用量表」。",
  },
  {
    path: "/admin/order-tool/summary/production-sheet",
    exact: true,
    selector: ".production-dish-card, .production-empty",
    title: "核對每道菜的用量",
    text: "看「本菜預估採購量」。這就是下一步要叫的數量。有草稿時，數字不對就直接改。",
  },
  {
    path: "/admin/order-tool/summary/production-sheet",
    exact: true,
    selector: "#usage-order-next",
    title: "去採購叫貨",
    text: "用量看完後，按「下一步：採購叫貨」。採購單會用這些數量來算。",
  },
  {
    path: "/admin/order-tool/summary/procurement",
    exact: true,
    selector: ".actual-qty-input, .procurement-empty",
    title: "填這次要叫的量",
    text: "在「實際採購量」填要跟廠商叫的數量。",
  },
  {
    path: "/admin/order-tool/summary/procurement",
    exact: true,
    selector: ".supplier-search-input, .procurement-empty",
    title: "填供應廠商",
    text: "在「供應廠商」選這項食材要向誰叫。沒有的話可以直接打新名字。",
  },
  {
    path: "/admin/order-tool/summary/procurement",
    exact: true,
    selector: ".order-confirm-toggle, .procurement-empty",
    title: "叫完就打勾",
    text: "跟廠商講完後，勾「已叫」，才知道這項已經叫過。",
  },
];
