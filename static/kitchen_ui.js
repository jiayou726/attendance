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

  document.querySelectorAll('[data-school-assignment]').forEach((form) => {
    const school = form.querySelector('select[name="school_id"]');
    const headcount = form.querySelector('input[name="headcount"]');
    if (!school || !headcount) return;
    school.addEventListener('change', () => {
      headcount.value = school.selectedOptions[0]?.dataset.headcount || '0';
    });
  });

  const recipeDataNode = document.getElementById('recipe-search-data');
  let recipeOptions = [];
  if (recipeDataNode) {
    try {
      recipeOptions = JSON.parse(recipeDataNode.textContent || '[]');
    } catch (_) {
      recipeOptions = [];
    }
  }

  document.querySelectorAll('[data-dish-search]').forEach((search) => {
    const input = search.querySelector('.dish-search-input');
    const recipeId = search.querySelector('input[name="recipe_id"]');
    const results = search.querySelector('.dish-search-results');
    const category = search.closest('form')?.querySelector('select[name="category"]');
    if (!input || !recipeId || !results) return;

    const closeResults = () => {
      results.hidden = true;
      results.replaceChildren();
    };

    const showMatches = () => {
      recipeId.value = '';
      const query = input.value.trim().toLocaleLowerCase('zh-Hant');
      results.replaceChildren();
      if (!query) {
        results.hidden = true;
        return;
      }

      const matches = recipeOptions.filter((recipe) =>
        String(recipe.name).toLocaleLowerCase('zh-Hant').includes(query)
      ).slice(0, 10);

      matches.forEach((recipe) => {
        const option = document.createElement('button');
        option.type = 'button';
        option.className = 'dish-search-option';
        option.setAttribute('role', 'option');
        const tag = document.createElement('span');
        tag.textContent = recipe.category || '其他';
        const name = document.createElement('b');
        name.textContent = recipe.name;
        option.append(tag, name);
        option.addEventListener('click', () => {
          input.value = recipe.name;
          recipeId.value = String(recipe.id);
          if (category && [...category.options].some((item) => item.value === recipe.category)) {
            category.value = recipe.category;
          }
          closeResults();
        });
        results.appendChild(option);
      });

      const exactMatch = recipeOptions.some((recipe) =>
        String(recipe.name).toLocaleLowerCase('zh-Hant') === query
      );
      if (!exactMatch) {
        const hint = document.createElement('div');
        hint.className = 'dish-create-hint';
        hint.textContent = matches.length
          ? `找不到完全相同的菜色；按新增會建立「${input.value.trim()}」`
          : `沒有相關菜色；按新增會建立「${input.value.trim()}」`;
        results.appendChild(hint);
      }
      results.hidden = false;
    };

    input.addEventListener('input', showMatches);
    input.addEventListener('focus', showMatches);
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') closeResults();
    });
    document.addEventListener('click', (event) => {
      if (!search.contains(event.target)) closeResults();
    });
  });

  document.querySelectorAll('.file-picker input[type="file"]').forEach((input) => {
    const label = input.closest('.file-picker');
    const text = label?.querySelector('span');
    if (!text) return;
    input.addEventListener('change', () => {
      text.textContent = input.files?.[0]?.name || '選擇 Excel 菜單';
    });
  });
})();
