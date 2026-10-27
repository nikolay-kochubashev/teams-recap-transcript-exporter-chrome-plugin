(() => {
  let pollTimer = null;

  const rangeMode = document.getElementById('currentChatRangeMode');
  const periodWrap = document.getElementById('currentChatPeriodWrap');
  const periodFromInput = document.getElementById('currentChatPeriodFrom');
  const periodToInput = document.getElementById('currentChatPeriodTo');
  const sinceWrap = document.getElementById('currentChatSinceWrap');
  const sinceInput = document.getElementById('currentChatSince');
  const monthsWrap = document.getElementById('currentChatMonthsWrap');
  const monthsInput = document.getElementById('currentChatMonths');
  const startBtn = document.getElementById('startCurrentChatExport');
  const stopBtn = document.getElementById('stopCurrentChatExport');
  const resetBtn = document.getElementById('resetCurrentChatExport');
  const openFolderBtn = document.getElementById('openCurrentChatFolder');
  const diagnosticBtn = document.getElementById('currentChatDiagnostic');
  const statusEl = document.getElementById('currentChatStatus');
  const metricsEl = document.getElementById('currentChatMetrics');
  const progressBar = document.getElementById('currentChatProgressBar');

  function runtimeMessage(message) {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendMessage(message, response => {
        const error = chrome.runtime.lastError;
        if (error) return reject(new Error(error.message));
        resolve(response);
      });
    });
  }

  async function boundTab() {
    try {
      const response = await runtimeMessage({ type: 'SIDE_PANEL_GET_BOUND_TAB' });
      if (response?.ok && response?.tab?.id) return response.tab;
    } catch (_) {}

    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tabs.length) throw new Error('Не удалось определить вкладку Teams.');
    return tabs[0];
  }

  function dateStamp(date) {
    return String(date.getFullYear()) +
      String(date.getMonth() + 1).padStart(2, '0') +
      String(date.getDate()).padStart(2, '0');
  }

  function isoDate(stamp) {
    if (!/^\d{8}$/.test(stamp || '')) return '';
    return stamp.slice(0, 4) + '-' + stamp.slice(4, 6) + '-' + stamp.slice(6, 8);
  }

  function stampFromIso(value) {
    return String(value || '').replace(/-/g, '');
  }

  function monthsAgo(months) {
    const now = new Date();
    const originalDay = now.getDate();
    const start = new Date(now.getFullYear(), now.getMonth(), 1);
    start.setMonth(start.getMonth() - months);
    const daysInTargetMonth = new Date(start.getFullYear(), start.getMonth() + 1, 0).getDate();
    start.setDate(Math.min(originalDay, daysInTargetMonth));
    return start;
  }

  function syncRangeUi() {
    const mode = rangeMode.value;
    periodWrap.hidden = mode !== 'period';
    sinceWrap.hidden = mode !== 'since';
    monthsWrap.hidden = mode !== 'months';
  }

  async function savePreferences() {
    await chrome.storage.local.set({
      currentChatRangeMode: rangeMode.value,
      currentChatPeriodFrom: periodFromInput.value,
      currentChatPeriodTo: periodToInput.value,
      currentChatSince: sinceInput.value,
      currentChatMonths: monthsInput.value
    });
  }

  function buildRange() {
    const mode = rangeMode.value;
    const today = new Date();
    const todayStamp = dateStamp(today);

    if (mode === 'all') {
      return { mode, startDate: '', endDate: '', months: 0 };
    }

    if (mode === 'period') {
      const startDate = stampFromIso(periodFromInput.value);
      const endDate = stampFromIso(periodToInput.value);

      if (!/^\d{8}$/.test(startDate) || !/^\d{8}$/.test(endDate)) {
        throw new Error('Укажи начало и конец периода.');
      }
      if (startDate > endDate) {
        throw new Error('Дата начала периода не может быть позже даты окончания.');
      }
      if (endDate > todayStamp) {
        throw new Error('Дата окончания периода не может быть в будущем.');
      }

      return { mode, startDate, endDate, months: 0 };
    }

    if (mode === 'since') {
      const startDate = stampFromIso(sinceInput.value);
      if (!/^\d{8}$/.test(startDate)) {
        throw new Error('Укажи дату, с которой читать сообщения.');
      }
      if (startDate > todayStamp) {
        throw new Error('Начальная дата не может быть в будущем.');
      }
      return { mode, startDate, endDate: todayStamp, months: 0 };
    }

    const months = Math.max(1, Math.min(60, Number(monthsInput.value || 0)));
    if (!Number.isFinite(months)) throw new Error('Укажи количество месяцев.');
    return {
      mode,
      startDate: dateStamp(monthsAgo(months)),
      endDate: todayStamp,
      months
    };
  }

  function render(state) {
    const running = state?.status === 'running';

    startBtn.disabled = running;
    stopBtn.disabled = !running;
    resetBtn.disabled = running;
    openFolderBtn.disabled = !state?.filePath;
    diagnosticBtn.disabled = running;
    rangeMode.disabled = running;
    periodFromInput.disabled = running;
    periodToInput.disabled = running;
    sinceInput.disabled = running;
    monthsInput.disabled = running;

    statusEl.textContent = state?.message || 'Открой нужный чат Teams и выбери диапазон сообщений.';

    const progress = Math.max(0, Math.min(100, Number(state?.progress || 0)));
    progressBar.style.width = progress + '%';

    const parts = [];
    if (state?.conversation) parts.push('Чат: ' + state.conversation);
    if (state?.messageCount) parts.push('сообщений: ' + state.messageCount);
    if (state?.loadedCount && state.loadedCount !== state.messageCount) {
      parts.push('загружено: ' + state.loadedCount);
    }
    metricsEl.textContent = parts.join(' | ');
  }

  async function refresh() {
    try {
      const response = await runtimeMessage({ type: 'CURRENT_CHAT_GET_STATE' });
      if (response?.ok) render(response.state);
    } catch (e) {
      statusEl.textContent = 'Ошибка: ' + (e.message || e);
    }
  }

  function ensurePolling() {
    if (pollTimer) return;
    pollTimer = setInterval(refresh, 800);
  }

  rangeMode.addEventListener('change', async () => {
    syncRangeUi();
    await savePreferences();
  });
  periodFromInput.addEventListener('change', savePreferences);
  periodToInput.addEventListener('change', savePreferences);
  sinceInput.addEventListener('change', savePreferences);
  monthsInput.addEventListener('change', savePreferences);

  startBtn.addEventListener('click', async () => {
    try {
      const tab = await boundTab();
      if (!/^https:\/\/(?:[^/]+\.)?teams\.microsoft\.com\//i.test(tab.url || '') &&
          !/^https:\/\/(?:[^/]+\.)?teams\.cloud\.microsoft\//i.test(tab.url || '')) {
        throw new Error('Открой нужный чат в Teams Web.');
      }

      const range = buildRange();
      await savePreferences();

      const response = await runtimeMessage({
        type: 'CURRENT_CHAT_START',
        tabId: tab.id,
        ...range
      });

      if (!response?.ok) throw new Error(response?.error || 'Не удалось запустить сбор.');

      statusEl.textContent = 'Запускаю чтение сообщений текущего чата...';
      ensurePolling();
      setTimeout(refresh, 250);
    } catch (e) {
      statusEl.textContent = 'Ошибка: ' + (e.message || e);
    }
  });

  stopBtn.addEventListener('click', async () => {
    try {
      await runtimeMessage({ type: 'CURRENT_CHAT_STOP' });
      await refresh();
    } catch (e) {
      statusEl.textContent = 'Ошибка: ' + (e.message || e);
    }
  });

  resetBtn.addEventListener('click', async () => {
    try {
      const response = await runtimeMessage({ type: 'CURRENT_CHAT_RESET' });
      if (response?.ok) render(response.state);
    } catch (e) {
      statusEl.textContent = 'Ошибка: ' + (e.message || e);
    }
  });

  openFolderBtn.addEventListener('click', async () => {
    try {
      const response = await runtimeMessage({ type: 'CURRENT_CHAT_OPEN_FOLDER' });
      if (!response?.ok) throw new Error(response?.error || 'Не удалось открыть папку.');
    } catch (e) {
      statusEl.textContent = 'Ошибка: ' + (e.message || e);
    }
  });

  diagnosticBtn.addEventListener('click', async () => {
    try {
      const tab = await boundTab();
      const response = await chrome.tabs.sendMessage(tab.id, { type: 'CURRENT_CHAT_DIAGNOSTIC' });
      if (!response?.ok) throw new Error(response?.error || 'Не удалось получить диагностику.');
      await navigator.clipboard.writeText(response.text || '');
      statusEl.textContent = 'Диагностика сообщений текущего чата скопирована в буфер обмена.';
    } catch (e) {
      statusEl.textContent = 'Ошибка диагностики: ' + (e.message || e);
    }
  });

  window.addEventListener('ttre:current-chat-mode', () => {
    syncRangeUi();
    refresh();
    ensurePolling();
  });

  (async () => {
    const saved = await chrome.storage.local.get([
      'currentChatRangeMode',
      'currentChatPeriodFrom',
      'currentChatPeriodTo',
      'currentChatSince',
      'currentChatMonths'
    ]);

    rangeMode.value = saved.currentChatRangeMode || 'all';
    monthsInput.value = saved.currentChatMonths || '3';

    const defaultFrom = isoDate(dateStamp(monthsAgo(1)));
    const defaultTo = isoDate(dateStamp(new Date()));

    periodFromInput.value = saved.currentChatPeriodFrom || defaultFrom;
    periodToInput.value = saved.currentChatPeriodTo || defaultTo;

    if (saved.currentChatSince) {
      sinceInput.value = saved.currentChatSince;
    } else {
      sinceInput.value = defaultFrom;
    }

    syncRangeUi();
    await refresh();
  })();
})();
