(() => {
  const path = window.location.pathname.replace(/\/$/, "");

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
      // Generated url_for links should be valid; ignore anything malformed.
    }
  });

  if (!best && path === '/admin/order-tool') {
    best = navLinks[0] || null;
  }

  if (best) {
    best.classList.add('active');
    best.setAttribute('aria-current', 'page');

    // On phones the bottom navigation is horizontally scrollable. Keep the
    // current section visible after each page load instead of leaving the
    // user looking at the left-most tabs.
    if (window.matchMedia('(max-width: 800px)').matches) {
      requestAnimationFrame(() => {
        best.scrollIntoView({ behavior: 'auto', block: 'nearest', inline: 'center' });
      });
    }
  }

  document.querySelectorAll('.card > table').forEach((table) => {
    if (table.parentElement?.classList.contains('table-scroll')) return;
    const wrapper = document.createElement('div');
    wrapper.className = 'table-scroll';
    table.parentNode.insertBefore(wrapper, table);
    wrapper.appendChild(table);
  });
})();
