(() => {
  const path = window.location.pathname.replace(/\/$/, "");

  // Highlight the current kitchen section in both desktop and mobile nav.
  const navLinks = document.querySelectorAll('.wrap > .top.no-print > .nav a');
  let best = null;
  let bestLength = -1;

  navLinks.forEach((link) => {
    try {
      const linkPath = new URL(link.href, window.location.origin).pathname.replace(/\/$/, "");
      const matches = path === linkPath || (linkPath !== '/admin/order-tool' && path.startsWith(linkPath + '/'));
      if (matches && linkPath.length > bestLength) {
        best = link;
        bestLength = linkPath.length;
      }
    } catch (_) {
      // Ignore malformed links; all generated url_for links should be valid.
    }
  });

  if (!best && path === '/admin/order-tool') {
    best = navLinks[0] || null;
  }
  if (best) {
    best.classList.add('active');
    best.setAttribute('aria-current', 'page');
  }

  // Wrap data tables so wide desktop tables remain usable on phones.
  document.querySelectorAll('.card > table').forEach((table) => {
    if (table.parentElement?.classList.contains('table-scroll')) return;
    const wrapper = document.createElement('div');
    wrapper.className = 'table-scroll';
    table.parentNode.insertBefore(wrapper, table);
    wrapper.appendChild(table);
  });
})();
