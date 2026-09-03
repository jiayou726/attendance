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
    const selectOnly = search.dataset.searchMode === 'select';
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
      if (!exactMatch || (selectOnly && matches.length === 0)) {
        const hint = document.createElement('div');
        hint.className = 'dish-create-hint';
        hint.textContent = selectOnly
          ? (matches.length ? '請點選上方的菜色。' : '找不到相關菜色。')
          : (matches.length
            ? `找不到完全相同的菜色；按新增會建立「${input.value.trim()}」`
            : `沒有相關菜色；按新增會建立「${input.value.trim()}」`);
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

    if (selectOnly) {
      search.closest('form')?.addEventListener('submit', (event) => {
        if (recipeId.value) return;
        const query = input.value.trim().toLocaleLowerCase('zh-Hant');
        const exact = recipeOptions.find((recipe) =>
          String(recipe.name).toLocaleLowerCase('zh-Hant') === query
        );
        if (exact) {
          recipeId.value = String(exact.id);
          return;
        }
        event.preventDefault();
        showMatches();
        input.focus();
      });
    }
  });

  const ingredientDataNode = document.getElementById('ingredient-search-data');
  let ingredientOptions = [];
  if (ingredientDataNode) {
    try {
      ingredientOptions = JSON.parse(ingredientDataNode.textContent || '[]');
    } catch (_) {
      ingredientOptions = [];
    }
  }

  document.querySelectorAll('[data-ingredient-search]').forEach((search) => {
    const input = search.querySelector('.ingredient-search-input');
    const ingredientId = search.querySelector('input[name="ingredient_id"]');
    const results = search.querySelector('.dish-search-results');
    if (!input || !ingredientId || !results) return;

    const closeResults = () => {
      results.hidden = true;
      results.replaceChildren();
    };

    const showMatches = () => {
      ingredientId.value = '';
      const query = input.value.trim().toLocaleLowerCase('zh-Hant');
      results.replaceChildren();
      if (!query) {
        results.hidden = true;
        return;
      }

      const matches = ingredientOptions.filter((ingredient) =>
        String(ingredient.name).toLocaleLowerCase('zh-Hant').includes(query)
      ).slice(0, 10);

      matches.forEach((ingredient) => {
        const option = document.createElement('button');
        option.type = 'button';
        option.className = 'dish-search-option';
        option.setAttribute('role', 'option');
        const tag = document.createElement('span');
        tag.textContent = `${ingredient.base_unit || 'g'}/人`;
        const name = document.createElement('b');
        name.textContent = `${ingredient.name}｜採購 ${ingredient.purchase_unit || '-'}`;
        option.append(tag, name);
        option.addEventListener('click', () => {
          input.value = ingredient.name;
          ingredientId.value = String(ingredient.id);
          closeResults();
        });
        results.appendChild(option);
      });

      const hint = document.createElement('div');
      hint.className = 'dish-create-hint';
      hint.textContent = matches.length ? '請點選上方的食材。' : '找不到相關食材。';
      results.appendChild(hint);
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
    search.closest('form')?.addEventListener('submit', (event) => {
      if (ingredientId.value) return;
      const query = input.value.trim().toLocaleLowerCase('zh-Hant');
      const exact = ingredientOptions.find((ingredient) =>
        String(ingredient.name).toLocaleLowerCase('zh-Hant') === query
      );
      if (exact) {
        ingredientId.value = String(exact.id);
        return;
      }
      event.preventDefault();
      showMatches();
      input.focus();
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
    const variantSwitch = schoolMenuForm.querySelector('[data-meal-variant-switch]');
    const currentMealLabel = schoolMenuForm.querySelector('[data-current-meal-label]');
    const showVariant = (variant) => {
      schoolMenuForm.querySelectorAll('[data-variant-panel]').forEach((panel) => {
        panel.hidden = panel.dataset.variantPanel !== variant;
      });
      variantSwitch?.querySelectorAll('[data-meal-variant]').forEach((button) => {
        button.classList.toggle('active', button.dataset.mealVariant === variant);
      });
      if (currentMealLabel) currentMealLabel.textContent = variant === 'vegetarian' ? '素食菜單' : '葷食菜單';
    };
    variantSwitch?.querySelectorAll('[data-meal-variant]').forEach((button) => {
      button.addEventListener('click', () => showVariant(button.dataset.mealVariant || 'regular'));
    });
    showVariant('regular');
    schoolMenuForm.addEventListener('submit', (event) => event.preventDefault());
    const csrf = schoolMenuForm.querySelector('input[name="_csrf_token"]')?.value || '';
    const schoolId = schoolMenuForm.dataset.schoolId || '';
    const schoolName = schoolMenuForm.dataset.schoolName || '';
    const saveUrl = schoolMenuForm.dataset.saveUrl || '';

    schoolMenuForm.querySelectorAll('[data-school-menu-day]').forEach((day) => {
      const headcount = day.querySelector('[data-variant-panel="regular"] input[type="number"]');
      const vegetarianHeadcount = day.querySelector('[data-variant-panel="vegetarian"] input[type="number"]');
      const checkboxes = [...day.querySelectorAll('.school-dish-check input[type="checkbox"]')];
      const noServiceToggle = day.querySelector('[data-no-service-toggle]');
      const serviceStatusBadge = day.querySelector('[data-service-status-badge]');
      const counts = {
        regular: day.querySelector('[data-selected-count="regular"]'),
        vegetarian: day.querySelector('[data-selected-count="vegetarian"]'),
      };
      const state = day.querySelector('[data-auto-save-state]');
      const serviceDate = day.dataset.serviceDate || '';
      let saveChain = Promise.resolve();
      let headcountSaveTimer = null;

      const isNoService = () => Boolean(noServiceToggle?.checked);
      const updateServiceStatus = () => {
        const noService = isNoService();
        day.classList.toggle('no-service', noService);
        if (headcount) headcount.disabled = noService || day.classList.contains('locked');
        if (vegetarianHeadcount) vegetarianHeadcount.disabled = noService || day.classList.contains('locked');
        checkboxes.forEach((item) => {
          item.disabled = noService || day.classList.contains('locked');
        });
        if (serviceStatusBadge) serviceStatusBadge.hidden = !noService;
      };

      const currentState = () => JSON.stringify({
        headcount: headcount?.value || '',
        vegetarianHeadcount: vegetarianHeadcount?.value || '',
        regularRecipeIds: checkboxes.filter((item) => item.checked && item.dataset.menuVariant === 'regular').map((item) => item.value),
        vegetarianRecipeIds: checkboxes.filter((item) => item.checked && item.dataset.menuVariant === 'vegetarian').map((item) => item.value),
        serviceStatus: isNoService() ? 'no_service' : 'serving',
      });
      let lastQueuedState = currentState();

      const updateCount = () => {
        Object.entries(counts).forEach(([variant, node]) => {
          if (node) node.textContent = String(checkboxes.filter((item) => item.checked && item.dataset.menuVariant === variant).length);
        });
      };
      const saveDay = () => {
        if (!saveUrl || !serviceDate || !headcount || !vegetarianHeadcount || day.classList.contains('locked')) return;
        if (!isNoService() && (headcount.value.trim() === '' || vegetarianHeadcount.value.trim() === '')) {
          if (state) state.textContent = '請輸入人數';
          return;
        }
        const nextState = currentState();
        if (nextState === lastQueuedState) return;
        lastQueuedState = nextState;
        updateCount();
        const body = new URLSearchParams({
          _csrf_token: csrf,
          school_id: schoolId,
          service_date: serviceDate,
          headcount: headcount.value,
          vegetarian_headcount: vegetarianHeadcount.value,
          service_status: isNoService() ? 'no_service' : 'serving',
        });
        checkboxes.filter((item) => item.checked && item.dataset.menuVariant === 'regular').forEach((item) => body.append('regular_recipe_ids', item.value));
        checkboxes.filter((item) => item.checked && item.dataset.menuVariant === 'vegetarian').forEach((item) => body.append('vegetarian_recipe_ids', item.value));
        if (state) state.textContent = '儲存中…';

        const operation = saveChain.then(async () => {
          const response = await fetch(saveUrl, {
            method: 'POST',
            headers: { 'X-Requested-With': 'school-menu-autosave' },
            body,
          });
          if (!response.ok) throw new Error('save failed');
          const regularCount = checkboxes.filter((item) => item.checked && item.dataset.menuVariant === 'regular').length;
          const vegetarianCount = checkboxes.filter((item) => item.checked && item.dataset.menuVariant === 'vegetarian').length;
          const names = new Set(missingSchools[serviceDate] || []);
          const complete = isNoService() || (
            regularCount > 0 && Number(headcount.value) > 0
            && (Number(vegetarianHeadcount.value) <= 0 || vegetarianCount > 0)
          );
          if (complete) names.delete(schoolName);
          else names.add(schoolName);
          missingSchools[serviceDate] = [...names];
          if (state) state.textContent = isNoService() ? '已停餐' : (complete ? '已儲存' : '尚未完成');
        });
        pendingSchoolMenuSaves.add(operation);
        saveChain = operation.catch(() => {
          lastQueuedState = '';
          if (state) state.textContent = '儲存失敗';
          window.alert('菜單自動儲存失敗，請重新整理後再試。');
        }).finally(() => pendingSchoolMenuSaves.delete(operation));
      };

      checkboxes.forEach((checkbox) => checkbox.addEventListener('change', saveDay));
      noServiceToggle?.addEventListener('change', () => {
        window.clearTimeout(headcountSaveTimer);
        updateServiceStatus();
        if (state) state.textContent = '儲存中…';
        saveDay();
      });
      headcount?.addEventListener('input', () => {
        window.clearTimeout(headcountSaveTimer);
        if (headcount.value.trim() === '') {
          if (state) state.textContent = '請輸入人數';
          return;
        }
        if (state) state.textContent = '等待儲存…';
        headcountSaveTimer = window.setTimeout(saveDay, 1000);
      });
      headcount?.addEventListener('change', () => {
        window.clearTimeout(headcountSaveTimer);
        saveDay();
      });
      vegetarianHeadcount?.addEventListener('input', () => {
        window.clearTimeout(headcountSaveTimer);
        if (vegetarianHeadcount.value.trim() === '') {
          if (state) state.textContent = '請輸入人數';
          return;
        }
        if (state) state.textContent = '等待儲存…';
        headcountSaveTimer = window.setTimeout(saveDay, 1000);
      });
      vegetarianHeadcount?.addEventListener('change', () => {
        window.clearTimeout(headcountSaveTimer);
        saveDay();
      });
      updateCount();
      updateServiceStatus();
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

  const procurementForm = document.querySelector('[data-procurement-autosave]');
  procurementForm?.addEventListener('submit', (event) => event.preventDefault());
  const procurementCsrf = procurementForm?.querySelector('input[name="_csrf_token"]')?.value || '';
  const supplierDatalist = document.getElementById('supplier-search-options');
  const procurementBody = procurementForm?.querySelector('.procurement-table tbody');
  const regroupProcurementRows = () => {
    if (!procurementBody) return;
    procurementBody.querySelectorAll('.supplier-group-row').forEach((row) => row.remove());
    const rows = [...procurementBody.querySelectorAll('[data-procurement-item]')];
    const compareText = (left, right) => {
      const normalizedLeft = String(left || '').toLocaleLowerCase('zh-Hant');
      const normalizedRight = String(right || '').toLocaleLowerCase('zh-Hant');
      if (normalizedLeft === normalizedRight) return 0;
      return normalizedLeft < normalizedRight ? -1 : 1;
    };
    rows.sort((left, right) => {
      const leftSupplier = left.dataset.supplierName || '⚠ 未指定供應商';
      const rightSupplier = right.dataset.supplierName || '⚠ 未指定供應商';
      const leftMissing = leftSupplier.startsWith('⚠');
      const rightMissing = rightSupplier.startsWith('⚠');
      if (leftMissing !== rightMissing) return leftMissing ? 1 : -1;
      const supplierOrder = compareText(leftSupplier, rightSupplier);
      if (supplierOrder) return supplierOrder;
      return compareText(left.dataset.ingredientName, right.dataset.ingredientName);
    });
    let currentSupplier = null;
    rows.forEach((row) => {
      const supplierName = row.dataset.supplierName || '⚠ 未指定供應商';
      if (supplierName !== currentSupplier) {
        const group = document.createElement('tr');
        group.className = 'supplier-group-row';
        const cell = document.createElement('td');
        cell.colSpan = 8;
        cell.textContent = supplierName;
        group.appendChild(cell);
        procurementBody.appendChild(group);
        currentSupplier = supplierName;
      }
      procurementBody.appendChild(row);
    });
  };

  document.querySelectorAll('[data-procurement-item]').forEach((row) => {
    const itemId = row.dataset.procurementItem;
    const actualInput = row.querySelector('.actual-qty-input');
    const packageInput = row.querySelector('.package-qty-input');
    const packageUnitInput = row.querySelector('.package-unit-input');
    const supplierInput = row.querySelector('.supplier-search-input');
    const hint = row.querySelector('.supplier-conversion-hint');
    const deliveryDateInput = row.querySelector('input[type="date"]');
    const deliverySlotInput = row.querySelector('.delivery-fields select');
    const saveState = row.querySelector('[data-procurement-save-state]');
    if (!actualInput || !packageInput || !packageUnitInput || !supplierInput || !hint) return;

    let activeRule = null;
    let saveTimer = null;
    let saveChain = Promise.resolve();
    const currentValues = () => JSON.stringify({
      actual: actualInput.value,
      packageQty: packageInput.value,
      packageUnit: packageUnitInput.value,
      deliveryDate: deliveryDateInput?.value || '',
      deliverySlot: deliverySlotInput?.value || '',
      supplierName: supplierInput.value.trim(),
    });
    let lastQueuedValues = currentValues();
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
      if (activeRule.packageUnit) {
        if (![...packageUnitInput.options].some((option) => option.value === activeRule.packageUnit)) {
          const option = document.createElement('option');
          option.value = activeRule.packageUnit;
          option.textContent = activeRule.packageUnit;
          packageUnitInput.appendChild(option);
        }
        packageUnitInput.value = activeRule.packageUnit;
      }
      if (overwrite || !packageInput.value) recalculatePackage();
    };

    const saveRow = () => {
      window.clearTimeout(saveTimer);
      if (!row.dataset.autoSaveUrl || !deliveryDateInput || !deliverySlotInput) return;
      if (actualInput.value.trim() === '' || deliveryDateInput.value === '') {
        if (saveState) saveState.textContent = '請完成欄位';
        return;
      }
      const nextValues = currentValues();
      if (nextValues === lastQueuedValues) return;
      lastQueuedValues = nextValues;
      if (saveState) saveState.textContent = '儲存中…';
      const body = new URLSearchParams({
        _csrf_token: procurementCsrf,
        actual: actualInput.value,
        package_qty: packageInput.value,
        package_unit: packageUnitInput.value,
        delivery_date: deliveryDateInput.value,
        delivery_slot: deliverySlotInput.value,
        supplier_name: supplierInput.value.trim(),
      });
      const operation = saveChain.then(async () => {
        const response = await fetch(row.dataset.autoSaveUrl, { method: 'POST', body });
        const data = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(data.message || '儲存失敗');
        packageInput.value = data.packageQty ?? packageInput.value;
        packageUnitInput.value = data.packageUnit ?? packageUnitInput.value;
        if (data.conversionLabel) hint.textContent = `廠商換算：${data.conversionLabel}`;
        row.dataset.supplierName = data.supplierName || '⚠ 未指定供應商';
        if (data.supplierCreated && supplierDatalist && data.supplierName) {
          const exists = [...supplierDatalist.options].some((option) => option.value === data.supplierName);
          if (!exists) {
            const option = document.createElement('option');
            option.value = data.supplierName;
            supplierDatalist.appendChild(option);
          }
        }
        lastQueuedValues = currentValues();
        regroupProcurementRows();
        if (saveState) saveState.textContent = data.supplierCreated ? '已儲存・已新增廠商' : '已儲存';
      });
      saveChain = operation.catch((error) => {
        lastQueuedValues = '';
        if (saveState) saveState.textContent = error.message || '儲存失敗';
      });
    };
    const scheduleSave = () => {
      if (saveState) saveState.textContent = '等待儲存…';
      window.clearTimeout(saveTimer);
      saveTimer = window.setTimeout(saveRow, 1000);
    };

    supplierInput.addEventListener('change', () => {
      selectSupplierRule({ overwrite: true });
      saveRow();
    });
    supplierInput.addEventListener('blur', saveRow);
    supplierInput.addEventListener('input', () => selectSupplierRule({ overwrite: true }));
    actualInput.addEventListener('input', () => {
      recalculatePackage();
      scheduleSave();
    });
    actualInput.addEventListener('change', saveRow);
    packageInput.addEventListener('input', () => {
      const packages = Number(packageInput.value);
      const factor = Number(activeRule?.purchasePerPackage);
      if (Number.isFinite(packages) && Number.isFinite(factor) && factor > 0) {
        actualInput.value = tidyQuantity(packages * factor);
      }
      scheduleSave();
    });
    packageInput.addEventListener('change', saveRow);
    packageUnitInput.addEventListener('input', scheduleSave);
    packageUnitInput.addEventListener('change', saveRow);
    deliveryDateInput?.addEventListener('change', saveRow);
    deliverySlotInput?.addEventListener('change', saveRow);
    selectSupplierRule();
  });
  regroupProcurementRows();

  document.querySelectorAll('[data-daily-count-row]').forEach((row) => {
    const inputs = [...row.querySelectorAll('[data-daily-count]')];
    const total = row.querySelector('[data-daily-total]');
    const refreshTotal = () => {
      if (!total) return;
      total.textContent = String(inputs.reduce((sum, input) => {
        const value = Number(input.value);
        return sum + (Number.isFinite(value) && value >= 0 ? value : 0);
      }, 0));
    };
    inputs.forEach((input) => input.addEventListener('input', refreshTotal));
    refreshTotal();
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
