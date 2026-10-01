(() => {
  const NATIVE_HOST = 'com.openai.teams_recap_transcript_exporter';
  const STATE_KEY = 'currentChatExportStateV1';

  let runningPromise = null;
  let stopRequested = false;

  function blankState() {
    return {
      status: 'idle',
      message: 'Открой нужный чат Teams.',
      tabId: null,
      mode: 'all',
      startDate: '',
      endDate: '',
      months: 0,
      conversation: '',
      messageCount: 0,
      loadedCount: 0,
      progress: 0,
      folderPath: '',
      filePath: '',
      error: '',
      startedAt: '',
      finishedAt: ''
    };
  }

  async function getState() {
    const data = await chrome.storage.local.get(STATE_KEY);
    return data?.[STATE_KEY] || blankState();
  }

  async function patchState(patch) {
    const state = Object.assign({}, await getState(), patch || {});
    await chrome.storage.local.set({ [STATE_KEY]: state });
    return state;
  }

  async function recoverInterruptedState() {
    try {
      const state = await getState();
      if (state.status === 'running' || /Останавливаю/i.test(state.message || '')) {
        await patchState({
          status: 'stopped',
          message: 'Предыдущий сбор был прерван перезапуском браузера.',
          error: '',
          finishedAt: new Date().toISOString()
        });
      }
    } catch (_) {}
  }

  function nativeMessage(message) {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendNativeMessage(NATIVE_HOST, message, response => {
        const error = chrome.runtime.lastError;
        if (error) return reject(new Error(error.message));
        if (!response) return reject(new Error('Windows helper не вернул ответ.'));
        if (!response.ok) return reject(new Error(response.error || 'Windows helper вернул ошибку.'));
        resolve(response);
      });
    });
  }

  async function sendToTab(tabId, message) {
    try {
      return await chrome.tabs.sendMessage(tabId, message);
    } catch (e) {
      throw new Error(
        'Не удалось связаться со вкладкой Teams. Обнови вкладку после обновления расширения. ' +
        (e.message || e)
      );
    }
  }

  function safeFileName(value) {
    return String(value || 'Teams chat')
      .replace(/[\\/:*?"<>|]+/g, '-')
      .replace(/\s+/g, ' ')
      .replace(/[. ]+$/g, '')
      .trim()
      .slice(0, 150) || 'Teams chat';
  }

  function displayStamp(stamp) {
    if (!/^\d{8}$/.test(stamp || '')) return stamp || '';
    return stamp.slice(6, 8) + '.' + stamp.slice(4, 6) + '.' + stamp.slice(0, 4);
  }

  function periodText(state) {
    if (state.mode === 'all') return 'весь доступный чат';
    if (state.mode === 'months') {
      return 'последние ' + state.months + ' мес. (' +
        displayStamp(state.startDate) + ' - ' + displayStamp(state.endDate) + ')';
    }
    return 'с ' + displayStamp(state.startDate) + ' по ' + displayStamp(state.endDate);
  }

  function addLinks(out, links) {
    const unique = Array.from(new Set(links || [])).filter(Boolean);
    if (!unique.length) return;
    out.push('Links:');
    unique.forEach(link => out.push('- ' + link));
  }

  function buildText(state, result) {
    const messages = result?.messages || [];
    const out = [
      'Teams current chat export',
      'Conversation: ' + (result?.conversation || state.conversation || ''),
      'Period: ' + periodText(state),
      'Messages: ' + messages.length,
      ''
    ];

    for (const message of messages) {
      out.push(
        '[' + (message.timestamp || message.dateTime || '') + '] ' +
        (message.author || 'Unknown')
      );

      if (message.quote?.preview) {
        const quoteHead = [
          message.quote.author || '',
          message.quote.timestamp ? '[' + message.quote.timestamp + ']' : ''
        ].filter(Boolean).join(' ');
        out.push('Reply to ' + (quoteHead || 'message') + ': ' + message.quote.preview);
      }

      if (message.message) out.push(message.message);
      addLinks(out, message.links);
      out.push('');
    }

    return out.join('\n').trim() + '\n';
  }

  async function runExport(request) {
    stopRequested = false;
    const startedAt = new Date().toISOString();

    let state = await patchState({
      status: 'running',
      message: 'Читаю текущий чат...',
      tabId: request.tabId,
      mode: request.mode || 'all',
      startDate: request.startDate || '',
      endDate: request.endDate || '',
      months: Number(request.months || 0),
      conversation: '',
      messageCount: 0,
      loadedCount: 0,
      progress: 10,
      folderPath: '',
      filePath: '',
      error: '',
      startedAt,
      finishedAt: ''
    });

    try {
      await nativeMessage({ action: 'ping' });
      const folder = await nativeMessage({ action: 'createBatchFolder' });
      state = await patchState({
        folderPath: folder.path,
        progress: 18,
        message: state.mode === 'all'
          ? 'Прокручиваю чат до самого начала...'
          : 'Прокручиваю чат до начала выбранного периода...'
      });

      const result = await sendToTab(request.tabId, {
        type: 'CURRENT_CHAT_COLLECT',
        mode: request.mode || 'all',
        startDate: request.startDate || '',
        endDate: request.endDate || ''
      });

      if (!result?.ok) {
        throw new Error(result?.error || 'Не удалось прочитать текущий чат.');
      }

      if (stopRequested) throw new Error('Остановлено пользователем.');

      state = await patchState({
        conversation: result.conversation || '',
        messageCount: result.count || 0,
        loadedCount: result.loaded || 0,
        progress: 90,
        message: 'Сохраняю TXT...'
      });

      const text = buildText(state, result);
      const periodSuffix = state.mode === 'all'
        ? 'all'
        : (state.startDate + '-' + state.endDate);
      const fileName =
        'Teams chat - ' + safeFileName(result.conversation || 'current') +
        ' - ' + periodSuffix + '.txt';

      const saved = await nativeMessage({
        action: 'saveTextInFolder',
        folderPath: state.folderPath,
        fileName,
        text
      });

      await patchState({
        status: 'done',
        message: 'Готово. Сообщений: ' + (result.count || 0) + '.',
        conversation: result.conversation || '',
        messageCount: result.count || 0,
        loadedCount: result.loaded || 0,
        progress: 100,
        filePath: saved.path,
        finishedAt: new Date().toISOString()
      });
    } catch (error) {
      const stopped = stopRequested || /Остановлено пользователем/i.test(error?.message || '');
      await patchState({
        status: stopped ? 'stopped' : 'error',
        message: stopped
          ? 'Сбор текущего чата остановлен.'
          : 'Ошибка: ' + (error?.message || error),
        error: stopped ? '' : (error?.message || String(error)),
        finishedAt: new Date().toISOString()
      });
    } finally {
      runningPromise = null;
      stopRequested = false;
    }
  }

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    const type = message?.type;

    if (type === 'CURRENT_CHAT_GET_STATE') {
      getState().then(state => sendResponse({ ok: true, state }));
      return true;
    }

    if (type === 'CURRENT_CHAT_START') {
      if (runningPromise) {
        sendResponse({ ok: false, error: 'Сбор текущего чата уже выполняется.' });
        return;
      }
      if (!message.tabId) {
        sendResponse({ ok: false, error: 'Не передана вкладка Teams.' });
        return;
      }

      runningPromise = runExport(message);
      sendResponse({ ok: true });
      return;
    }

    if (type === 'CURRENT_CHAT_STOP') {
      (async () => {
        if (!runningPromise) {
          const state = await patchState({
            status: 'stopped',
            message: 'Сбор текущего чата остановлен.',
            error: '',
            finishedAt: new Date().toISOString()
          });
          sendResponse({ ok: true, state });
          return;
        }

        stopRequested = true;
        const state = await getState();
        try {
          if (state.tabId) {
            await sendToTab(state.tabId, { type: 'CURRENT_CHAT_CANCEL' });
          }
        } catch (_) {}

        const next = await patchState({ message: 'Останавливаю...' });
        sendResponse({ ok: true, state: next });
      })();
      return true;
    }

    if (type === 'CURRENT_CHAT_OPEN_FOLDER') {
      getState().then(async state => {
        // Use the exact saved file as the source of truth. folderPath can be
        // stale for a short time between runs, while filePath is written only
        // after the current export has successfully completed.
        if (state.filePath) {
          const response = await nativeMessage({ action: 'showInFolder', path: state.filePath });
          sendResponse(response);
          return;
        }

        if (!state.folderPath) throw new Error('Результат еще не сохранен.');
        const response = await nativeMessage({ action: 'openDirectory', path: state.folderPath });
        sendResponse(response);
      }).catch(error => sendResponse({ ok: false, error: error.message || String(error) }));
      return true;
    }

    if (type === 'CURRENT_CHAT_RESET') {
      chrome.storage.local.set({ [STATE_KEY]: blankState() })
        .then(() => sendResponse({ ok: true, state: blankState() }));
      return true;
    }
  });

  recoverInterruptedState();
})();
