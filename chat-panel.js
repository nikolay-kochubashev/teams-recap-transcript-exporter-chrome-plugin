(() => {
  let pollTimer = null;
  let localStatusUntil = 0;

  const section = document.getElementById('chatSection');
  const authorInput = document.getElementById('chatAuthor');
  const periodEl = document.getElementById('chatPeriod');
  const statusEl = document.getElementById('chatStatus');
  const metricsEl = document.getElementById('chatMetrics');
  const progressBar = document.getElementById('chatProgressBar');
  const startBtn = document.getElementById('startChatExport');
  const stopBtn = document.getElementById('stopChatExport');
  const openFolderBtn = document.getElementById('openChatFolder');
  const diagnosticBtn = document.getElementById('chatDiagnostic');

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
    // SIDE_PANEL_GET_BOUND_TAB пока не реализован background.js.
    // Поэтому запрос может завершиться chrome.runtime.lastError. Это штатный
    // случай: используем активную вкладку того же окна, как уже делает panel.js.
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

  function displayDate(stamp) {
    return stamp.slice(6, 8) + '.' + stamp.slice(4, 6) + '.' + stamp.slice(0, 4);
  }

  function previousCalendarWeek() {
    const now = new Date();
    const day = now.getDay() || 7;
    const mondayThisWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - day + 1);
    const start = new Date(mondayThisWeek.getFullYear(), mondayThisWeek.getMonth(), mondayThisWeek.getDate() - 7);
    const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6);
    return { start: dateStamp(start), end: dateStamp(end) };
  }

  function currentPeriod() {
    const range = previousCalendarWeek();
    periodEl.textContent = 'Прошлая неделя: ' + displayDate(range.start) + ' - ' + displayDate(range.end);
    return range;
  }

  async function detectAuthor() {
    if (authorInput.value.trim()) return authorInput.value.trim();

    try {
      const tab = await boundTab();
      const response = await chrome.tabs.sendMessage(tab.id, { type: 'CHAT_DETECT_SELF' });
      if (response?.ok && response.author) {
        authorInput.value = response.author;
        await chrome.storage.local.set({ weeklyChatAuthorQuery: response.author });
        return response.author;
      }
    } catch (_) {}

    return '';
  }

  function showLocalStatus(message, ttlMs = 5000) {
    localStatusUntil = Date.now() + ttlMs;
    statusEl.textContent = message;
  }

  function render(state) {
    const running = state?.status === 'running';
    startBtn.disabled = running;
    stopBtn.disabled = !running;
    openFolderBtn.disabled = !state?.folderPath;
    diagnosticBtn.disabled = running;

    if (Date.now() >= localStatusUntil || running || state?.status === 'done' || state?.status === 'error') {
      statusEl.textContent = state?.message || 'Готово к сбору переписки.';
    }
    const progress = Math.max(0, Math.min(100, Number(state?.progress || 0)));
    progressBar.style.width = progress + '%';

    if (state?.pages || state?.messageCount) {
      metricsEl.textContent =
        'Страниц: ' + (state.pages || 0) +
        ' | сообщений: ' + (state.messageCount || 0) +
        ' | чатов: ' + (state.conversationCount || 0);
    } else {
      metricsEl.textContent = '';
    }

    if (!authorInput.value && state?.authorQuery) authorInput.value = state.authorQuery;
  }

  async function refresh() {
    try {
      const response = await runtimeMessage({ type: 'CHAT_GET_STATE' });
      if (response?.ok) render(response.state);
    } catch (e) {
      showLocalStatus('Ошибка: ' + (e.message || e));
    }
  }

  function ensurePolling() {
    if (pollTimer) return;
    pollTimer = setInterval(refresh, 800);
  }

  startBtn.addEventListener('click', async () => {
    try {
      const tab = await boundTab();
      if (!/^https:\/\/(?:[^/]+\.)?teams\.microsoft\.com\//i.test(tab.url || '') &&
          !/^https:\/\/(?:[^/]+\.)?teams\.cloud\.microsoft\//i.test(tab.url || '')) {
        throw new Error('Открой обычную вкладку Teams Web.');
      }

      const author = await detectAuthor();
      if (!author) {
        authorInput.focus();
        throw new Error('Укажи имя или фамилию, по которой Teams находит тебя в People.');
      }

      await chrome.storage.local.set({ weeklyChatAuthorQuery: author });
      const range = currentPeriod();

      const response = await runtimeMessage({
        type: 'CHAT_START',
        tabId: tab.id,
        authorQuery: author,
        startDate: range.start,
        endDate: range.end
      });

      if (!response?.ok) throw new Error(response?.error || 'Не удалось запустить сбор.');
      localStatusUntil = 0;
      statusEl.textContent = 'Запускаю сбор переписки...';
      ensurePolling();
      setTimeout(refresh, 250);
    } catch (e) {
      statusEl.textContent = 'Ошибка: ' + (e.message || e);
    }
  });

  stopBtn.addEventListener('click', async () => {
    try {
      await runtimeMessage({ type: 'CHAT_STOP' });
      await refresh();
    } catch (e) {
      statusEl.textContent = 'Ошибка: ' + (e.message || e);
    }
  });

  openFolderBtn.addEventListener('click', async () => {
    try {
      const response = await runtimeMessage({ type: 'CHAT_OPEN_FOLDER' });
      if (!response?.ok) throw new Error(response?.error || 'Не удалось открыть папку.');
    } catch (e) {
      statusEl.textContent = 'Ошибка: ' + (e.message || e);
    }
  });

  diagnosticBtn.addEventListener('click', async () => {
    try {
      const tab = await boundTab();
      const response = await chrome.tabs.sendMessage(tab.id, { type: 'CHAT_SEARCH_DIAGNOSTIC' });
      if (!response?.ok) throw new Error(response?.error || 'Не удалось получить диагностику.');
      await navigator.clipboard.writeText(response.text || '');
      statusEl.textContent = 'Диагностика Teams Search скопирована в буфер обмена.';
    } catch (e) {
      showLocalStatus('Ошибка диагностики: ' + (e.message || e));
    }
  });

  window.addEventListener('ttre:chat-mode', () => {
    currentPeriod();
    refresh();
    detectAuthor();
    ensurePolling();
  });

  (async () => {
    currentPeriod();
    const saved = await chrome.storage.local.get('weeklyChatAuthorQuery');
    if (saved?.weeklyChatAuthorQuery) authorInput.value = saved.weeklyChatAuthorQuery;
    await refresh();
  })();
})();