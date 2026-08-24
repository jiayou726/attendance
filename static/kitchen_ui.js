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

  document.querySelectorAll('[data-school-picker] select[name="school_id"]').forEach((select) => {
    select.addEventListener('change', () => select.form?.requestSubmit());
  });

  document.querySelectorAll('.school-dish-check input[type="checkbox"]').forEach((checkbox) => {
    const refresh = () => checkbox.closest('.school-dish-check')?.classList.toggle('selected', checkbox.checked);
    checkbox.addEventListener('change', refresh);
    refresh();
  });

  const missingSchoolsNode = document.getElementById('missing-schools-data');
  let missingSchools = {};
  if (missingSchoolsNode) {
    try {
      missingSchools = JSON.parse(missingSchoolsNode.textContent || '{}');
    } catch (_) {
      missingSchools = {};
    }
  }

  const pendingSchoolMenuSaves = new Set();
  const schoolMenuForm = document.querySelector('[data-school-menu-autosave]');
  if (schoolMenuForm) {
    schoolMenuForm.addEventListener('submit', (event) => event.preventDefault());
    const csrf = schoolMenuForm.querySelector('input[name="_csrf_token"]')?.value || '';
    const schoolId = schoolMenuForm.dataset.schoolId || '';
    const schoolName = schoolMenuForm.dataset.schoolName || '';
    const saveUrl = schoolMenuForm.dataset.saveUrl || '';

    schoolMenuForm.querySelectorAll('[data-school-menu-day]').forEach((day) => {
      const headcount = day.querySelector('.headcount-box input');
      const checkboxes = [...day.querySelectorAll('.school-dish-check input[type="checkbox"]')];
      const count = day.querySelector('[data-selected-count]');
      const state = day.querySelector('[data-auto-save-state]');
      const serviceDate = day.dataset.serviceDate || '';
      let saveChain = Promise.resolve();

      const updateCount = () => {
        if (count) count.textContent = String(checkboxes.filter((item) => item.checked).length);
      };
      const saveDay = () => {
        if (!saveUrl || !serviceDate || !headcount || day.classList.contains('locked')) return;
        updateCount();
        const body = new URLSearchParams({
          _csrf_token: csrf,
          school_id: schoolId,
          service_date: serviceDate,
          headcount: headcount.value,
        });
        checkboxes.filter((item) => item.checked).forEach((item) => body.append('recipe_ids', item.value));
        if (state) state.textContent = '儲存中…';

        const operation = saveChain.then(async () => {
          const response = await fetch(saveUrl, {
            method: 'POST',
            headers: { 'X-Requested-With': 'school-menu-autosave' },
            body,
          });
          if (!response.ok) throw new Error('save failed');
          const selectedCount = checkboxes.filter((item) => item.checked).length;
          const names = new Set(missingSchools[serviceDate] || []);
          if (selectedCount) names.delete(schoolName);
          else names.add(schoolName);
          missingSchools[serviceDate] = [...names];
          if (state) state.textContent = '已儲存';
        });
        pendingSchoolMenuSaves.add(operation);
        saveChain = operation.catch(() => {
          if (state) state.textContent = '儲存失敗';
          window.alert('菜單自動儲存失敗，請重新整理後再試。');
        }).finally(() => pendingSchoolMenuSaves.delete(operation));
      };

      checkboxes.forEach((checkbox) => checkbox.addEventListener('change', saveDay));
      headcount?.addEventListener('change', saveDay);
      updateCount();
    });
  }

  document.querySelectorAll('[data-procurement-generate]').forEach((form) => {
    form.addEventListener('submit', async (event) => {
      if (form.dataset.ready === 'true') return;
      event.preventDefault();
      await Promise.allSettled([...pendingSchoolMenuSaves]);
      const serviceDate = form.querySelector('input[name="date"]')?.value || '';
      const names = missingSchools[serviceDate] || [];
      if (names.length) {
        window.alert(`尚有學校未完成菜單勾選：${names.join('、')}。請先完成後再產生採購單。`);
        return;
      }
      form.dataset.ready = 'true';
      form.requestSubmit();
    });
  });

  const supplierConversionNode = document.getElementById('supplier-conversion-data');
  let supplierConversions = {};
  if (supplierConversionNode) {
    try {
      supplierConversions = JSON.parse(supplierConversionNode.textContent || '{}');
    } catch (_) {
      supplierConversions = {};
    }
  }

  const tidyQuantity = (value) => {
    if (!Number.isFinite(value)) return '';
    return value.toFixed(4).replace(/\.?0+$/, '');
  };

  document.querySelectorAll('[data-procurement-item]').forEach((row) => {
    const itemId = row.dataset.procurementItem;
    const actualInput = row.querySelector('.actual-qty-input');
    const packageInput = row.querySelector('.package-qty-input');
    const packageUnitInput = row.querySelector('.package-unit-input');
    const supplierInput = row.querySelector('.supplier-search-input');
    const hint = row.querySelector('.supplier-conversion-hint');
    if (!actualInput || !packageInput || !packageUnitInput || !supplierInput || !hint) return;

    let activeRule = null;
    const recalculatePackage = () => {
      const actual = Number(actualInput.value);
      const factor = Number(activeRule?.purchasePerPackage);
      if (Number.isFinite(actual) && Number.isFinite(factor) && factor > 0) {
        packageInput.value = tidyQuantity(actual / factor);
      }
    };
    const selectSupplierRule = ({ overwrite = false } = {}) => {
      activeRule = supplierConversions[itemId]?.[supplierInput.value.trim()] || null;
      if (!activeRule) {
        hint.textContent = '此廠商未提供這項食材的換算，可自行填寫';
        return;
      }
      hint.textContent = `廠商換算：${activeRule.label}`;
      if (activeRule.packageUnit) packageUnitInput.value = activeRule.packageUnit;
      if (overwrite || !packageInput.value) recalculatePackage();
    };

    supplierInput.addEventListener('change', () => selectSupplierRule({ overwrite: true }));
    supplierInput.addEventListener('input', () => selectSupplierRule({ overwrite: true }));
    actualInput.addEventListener('input', recalculatePackage);
    packageInput.addEventListener('input', () => {
      const packages = Number(packageInput.value);
      const factor = Number(activeRule?.purchasePerPackage);
      if (Number.isFinite(packages) && Number.isFinite(factor) && factor > 0) {
        actualInput.value = tidyQuantity(packages * factor);
      }
    });
    selectSupplierRule();
  });

  document.querySelectorAll('.order-confirm-toggle').forEach((checkbox) => {
    const refresh = () => checkbox.closest('tr')?.classList.toggle('is-ordered', checkbox.checked);
    checkbox.addEventListener('change', async () => {
      refresh();
      if (checkbox.dataset.autoSubmit === 'true') checkbox.form?.requestSubmit();
      if (checkbox.dataset.autoSaveUrl) {
        checkbox.disabled = true;
        const body = new URLSearchParams({ _csrf_token: checkbox.dataset.csrf || '' });
        if (checkbox.checked) body.set('ordered', '1');
        try {
          const response = await fetch(checkbox.dataset.autoSaveUrl, {
            method: 'POST',
            headers: { 'X-Requested-With': 'procurement-tracking' },
            body,
          });
          if (!response.ok) throw new Error('save failed');
        } catch (_) {
          checkbox.checked = !checkbox.checked;
          refresh();
          window.alert('叫貨狀態儲存失敗，請重新整理後再試。');
        } finally {
          checkbox.disabled = false;
        }
      }
    });
    refresh();
  });
})();
