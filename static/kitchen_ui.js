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
    const copyDialog = document.querySelector('[data-menu-copy-dialog]');
    const copyForm = copyDialog?.querySelector('[data-menu-copy-form]');
    const copyFeedback = document.querySelector('[data-copy-feedback]');
    const copySelectAll = copyForm?.querySelector('[data-copy-select-all]');
    const copyTargets = [...(copyForm?.querySelectorAll('[data-copy-target]') || [])];
    const copyStatus = copyForm?.querySelector('[data-copy-modal-status]');
    const copySubmit = copyForm?.querySelector('[data-copy-submit]');
    const closeCopyDialog = () => {
      if (copyDialog) copyDialog.hidden = true;
    };
    copyDialog?.querySelectorAll('[data-copy-close]').forEach((button) => {
      button.addEventListener('click', closeCopyDialog);
    });
    copyDialog?.addEventListener('click', (event) => {
      if (event.target === copyDialog) closeCopyDialog();
    });
    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape' && copyDialog && !copyDialog.hidden) closeCopyDialog();
    });
    copySelectAll?.addEventListener('change', () => {
      copyTargets.forEach((input) => { input.checked = copySelectAll.checked; });
    });
    copyTargets.forEach((input) => input.addEventListener('change', () => {
      if (copySelectAll) copySelectAll.checked = copyTargets.every((target) => target.checked);
    }));
    copyForm?.addEventListener('submit', async (event) => {
      event.preventDefault();
      const selected = copyTargets.filter((input) => input.checked);
      if (!selected.length) {
        if (copyStatus) copyStatus.textContent = '請至少選擇一所目標學校。';
        return;
      }
      if (copySubmit) copySubmit.disabled = true;
      if (copyStatus) copyStatus.textContent = '複製中…';
      try {
        const body = new URLSearchParams(new FormData(copyForm));
        const response = await fetch(copyForm.dataset.copyUrl, {
          method: 'POST',
          headers: { 'X-Requested-With': 'school-menu-copy' },
          body,
        });
        const data = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(data.message || '複製失敗');
        closeCopyDialog();
        if (copyFeedback) {
          copyFeedback.hidden = false;
          copyFeedback.textContent = `${data.message} 若已產生採購草稿，請重新產生採購單更新食材需求。`;
        }
      } catch (error) {
        if (copyStatus) copyStatus.textContent = error.message || '複製失敗';
      } finally {
        if (copySubmit) copySubmit.disabled = false;
      }
    });

    schoolMenuForm.querySelectorAll('[data-school-menu-day]').forEach((day) => {
      const headcount = day.querySelector('[data-variant-panel="regular"] input[type="number"]');
      const vegetarianHeadcount = day.querySelector('[data-variant-panel="vegetarian"] input[type="number"]');
      const checkboxes = [...day.querySelectorAll('.school-dish-check input[type="checkbox"]')];
      const noServiceToggle = day.querySelector('[data-no-service-toggle]');
      const serviceStatusBadge = day.querySelector('[data-service-status-badge]');
      const copyButton = day.querySelector('[data-copy-day-menu]');
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
        if (copyButton) {
          copyButton.disabled = noService || day.classList.contains('locked')
            || !checkboxes.some((item) => item.checked);
        }
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
        if (copyButton) {
          copyButton.disabled = isNoService() || day.classList.contains('locked')
            || !checkboxes.some((item) => item.checked);
        }
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

      copyButton?.addEventListener('click', async () => {
        // A copy is always based on the SAVED source menu, not unsent UI edits.
        window.clearTimeout(headcountSaveTimer);
        saveDay();
        copyButton.disabled = true;
        try {
          await Promise.all([...pendingSchoolMenuSaves]);
          if (!copyDialog || !copyForm) return;
          copyForm.reset();
          copyForm.querySelector('input[name="service_date"]').value = serviceDate;
          if (copySelectAll) copySelectAll.checked = false;
          if (copyStatus) copyStatus.textContent = '';
          const label = copyForm.querySelector('[data-copy-description]');
          if (label) label.textContent = `來源：${schoolName}　日期：${serviceDate}（同一天複製）`;
          copyDialog.hidden = false;
          copyTargets[0]?.focus();
        } catch (_) {
          window.alert('來源菜單尚未成功儲存，請先確認儲存狀態。');
        } finally {
          updateCount();
        }
      });
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


  // Independent per-dish estimates, serialized by purchase-item ID.
  const productionGroups = new Map();
  document.querySelectorAll('[data-production-key]').forEach((editor) => {
    const id = editor.dataset.productionItem;
    if (!productionGroups.has(id)) {
      productionGroups.set(id, { editors: [], queue: Promise.resolve(), total: editor.dataset.productionTotal });
    }
    productionGroups.get(id).editors.push(editor);
  });
  productionGroups.forEach((group, itemId) => {
    const updateTotal = (actual) => {
      group.total = String(actual);
      document.querySelectorAll('[data-production-day-total]').forEach((node) => {
        if (node.dataset.productionDayTotal !== itemId) return;
        const unit = group.editors[0]?.querySelector('.input-unit span')?.textContent || '';
        node.textContent = `當日採購總量：${group.total} ${unit}`;
      });
    };
    group.editors.forEach((editor) => {
      const input = editor.querySelector('.production-estimate-input');
      const status = editor.querySelector('[data-production-save-state]');
      if (!input) return;
      let committed = input.value;
      let queued = committed;
      let timer = null;
      const save = () => {
        window.clearTimeout(timer);
        const next = input.value.trim();
        if (!next || !Number.isFinite(Number(next)) || Number(next) < 0) {
          if (status) status.textContent = '請輸入有效的非負數量';
          return;
        }
        if (next === queued) return;
        queued = next;
        if (next === committed) {
          if (status) status.textContent = '已儲存';
          return;
        }
        if (status) status.textContent = '儲存中…';
        group.queue = group.queue.then(async () => {
          const body = new URLSearchParams({
            _csrf_token: editor.dataset.productionCsrf,
            date: editor.dataset.productionDate,
            estimate_key: editor.dataset.productionKey,
            estimate: next,
            expected_estimate: committed,
            expected_actual: group.total,
          });
          const response = await fetch(editor.dataset.productionSaveUrl, { method: 'POST', body });
          const data = await response.json().catch(() => ({}));
          if (!response.ok) throw new Error(data.message || '儲存失敗，請重新整理');
          committed = String(data.estimate);
          updateTotal(data.actual);
          if (document.activeElement !== input || input.value === next) input.value = committed;
          if (status) status.textContent = '已儲存';
        }).catch((error) => {
          queued = null; // retries are permitted after a stale-page conflict
          if (status) status.textContent = error.message || '儲存失敗';
        });
      };
      input.addEventListener('input', () => {
        if (status) status.textContent = '等待儲存…';
        window.clearTimeout(timer);
        timer = window.setTimeout(save, 850);
      });
      input.addEventListener('change', save);
      input.addEventListener('blur', save);
    });
  });

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
    let lastSavedActual = actualInput.value;
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
      const values = {
        actual: actualInput.value,
        package_qty: packageInput.value,
        package_unit: packageUnitInput.value,
        delivery_date: deliveryDateInput.value,
        delivery_slot: deliverySlotInput.value,
        supplier_name: supplierInput.value.trim(),
      };
      const operation = saveChain.then(async () => {
        const body = new URLSearchParams({
          _csrf_token: procurementCsrf,
          ...values,
          expected_actual: lastSavedActual,
        });
        const response = await fetch(row.dataset.autoSaveUrl, { method: 'POST', body });
        const data = await response.json().catch(() => ({}));
        if (!response.ok) {
          if (response.status === 409 && data.actual !== undefined) {
            actualInput.value = String(data.actual);
            packageInput.value = data.packageQty ?? packageInput.value;
            packageUnitInput.value = data.packageUnit ?? packageUnitInput.value;
            lastSavedActual = String(data.actual);
          }
          throw new Error(data.message || '儲存失敗');
        }
        lastSavedActual = String(data.actual ?? values.actual);
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
    const classInputs = [...row.querySelectorAll('[data-daily-class-count]')];
    const classTotal = row.querySelector('[data-daily-class-total]');
    const refreshTotal = () => {
      if (!total) return;
      total.textContent = String(inputs.reduce((sum, input) => {
        const value = Number(input.value);
        return sum + (Number.isFinite(value) && value >= 0 ? value : 0);
      }, 0));
    };
    inputs.forEach((input) => input.addEventListener('input', refreshTotal));
    const refreshClassTotal = () => {
      if (!classTotal) return;
      classTotal.textContent = String(classInputs.reduce((sum, input) => {
        const value = Number(input.value);
        return sum + (input.value.trim() !== '' && Number.isFinite(value) && value >= 0 ? value : 0);
      }, 0));
    };
    classInputs.forEach((input) => input.addEventListener('input', refreshClassTotal));
    refreshTotal();
    refreshClassTotal();
  });

  const dailyKitchenForm = document.querySelector('[data-daily-kitchen-autosave]');
  if (dailyKitchenForm) {
    const saveState = dailyKitchenForm.querySelector('[data-daily-kitchen-save-state]');
    const fields = [...dailyKitchenForm.querySelectorAll('[data-daily-save-field]')];
    const numericFields = fields.filter((field) => field.matches('input[type="number"]'));
    let saveTimer = null;
    let saveChain = Promise.resolve();
    let lastSaved = new URLSearchParams(new FormData(dailyKitchenForm)).toString();

    dailyKitchenForm.addEventListener('submit', (event) => event.preventDefault());
    const save = () => {
      window.clearTimeout(saveTimer);
      saveTimer = null;
      if (!numericFields.every((field) => field.checkValidity())) {
        if (saveState) saveState.textContent = '請修正數字';
        return;
      }
      const body = new URLSearchParams(new FormData(dailyKitchenForm));
      const serialized = body.toString();
      if (serialized === lastSaved) {
        if (saveState) saveState.textContent = '已儲存';
        return;
      }
      if (saveState) saveState.textContent = '儲存中…';
      const operation = saveChain.then(async () => {
        const response = await fetch(dailyKitchenForm.action, {
          method: 'POST',
          headers: { 'X-Requested-With': 'daily-kitchen-autosave' },
          body,
          keepalive: true,
        });
        const data = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(data.message || '儲存失敗');
        lastSaved = serialized;
        if (saveState) saveState.textContent = data.message || '已儲存';
      });
      saveChain = operation.catch((error) => {
        const message = error.message || '請重新整理後再試';
        if (saveState) saveState.textContent = `儲存失敗：${message}`;
      });
    };
    const scheduleSave = () => {
      if (saveState) saveState.textContent = '等待儲存…';
      window.clearTimeout(saveTimer);
      saveTimer = window.setTimeout(save, 400);
    };
    fields.forEach((field) => {
      field.addEventListener('input', scheduleSave);
      field.addEventListener('change', save);
    });
  }

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
