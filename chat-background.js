(() => {
  const CHAT_NATIVE_HOST = 'com.openai.teams_recap_transcript_exporter';
  const CHAT_STATE_KEY = 'weeklyChatExportStateV1';
  const MAX_PAGES = 100;
  let stopRequested = false;
  let runningPromise = null;

  function blankState() {
    return {
      status: 'idle',
      message: 'Готово к сбору переписки.',
      tabId: null,
      authorQuery: '',
      author: '',
      startDate: '',
      endDate: '',
      pages: 0,
      messageCount: 0,
      conversationCount: 0,
      progress: 0,
      folderPath: '',
      filePath: '',
      logPath: '',
      error: '',
      startedAt: '',
      finishedAt: ''
    };
  }

  async function getState() {
    const data = await chrome.storage.local.get(CHAT_STATE_KEY);
    return data?.[CHAT_STATE_KEY] || blankState();
  }

  async function patchState(patch) {
    const state = Object.assign({}, await getState(), patch || {});
    await chrome.storage.local.set({ [CHAT_STATE_KEY]: state });
    return state;
  }

  function nativeMessage(message) {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendNativeMessage(CHAT_NATIVE_HOST, message, response => {
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
      throw new Error('Не удалось связаться со вкладкой Teams. Обнови вкладку после обновления расширения. ' + (e.message || e));
    }
  }

  function stampDisplay(stamp) {
    if (!/^\d{8}$/.test(stamp || '')) return stamp || '';
    return stamp.slice(6, 8) + '.' + stamp.slice(4, 6) + '.' + stamp.slice(0, 4);
  }

  function safeFileName(value) {
    return String(value || 'Teams messages')
      .replace(/[\\/:*?"<>|]+/g, '-')
      .replace(/\s+/g, ' ')
      .replace(/[. ]+$/g, '')
      .trim()
      .slice(0, 170) || 'Teams messages';
  }

  function uniqueKey(item) {
    return item?.id || item?.key || [
      item?.conversation || '',
      item?.sender || '',
      item?.timestamp || '',
      item?.message || ''
    ].join('|');
  }

  function buildText(state, messages) {
    const conversations = Array.from(new Set(
      messages.map(x => x.conversation).filter(Boolean)
    )).sort((a, b) => a.localeCompare(b, 'ru'));

    const ordered = messages.slice().reverse();
    const out = [
      'Teams messages for work report',
      'Period: ' + stampDisplay(state.startDate) + ' - ' + stampDisplay(state.endDate),
      'Author: ' + (state.author || state.authorQuery || ''),
      'Messages: ' + ordered.length,
      'Conversations: ' + conversations.length,
      '',
      'Conversations:',
      ...conversations.map(x => '- ' + x),
      '',
      'Messages',
      '========',
      ''
    ];

    ordered.forEach((item, index) => {
      const head = [
        item.timestamp ? '[' + item.timestamp + ']' : '',
        item.conversation || '',
        item.sender && item.sender !== state.author ? '- ' + item.sender : ''
      ].filter(Boolean).join(' ');

      out.push(head || ('Message ' + (index + 1)));
      out.push(item.message || '');

      const links = Array.from(new Set(item.links || []));
      if (links.length) {
        out.push('Links:');
        links.forEach(link => out.push('- ' + link));
      }
      out.push('');
    });

    return out.join('\n').trim() + '\n';
  }

  function buildLog(entries, state) {
    const out = [
      'Teams Recap Transcript Exporter - weekly chat operation log',
      'Started: ' + (state.startedAt || ''),
      'Finished: ' + (state.finishedAt || ''),
      'Period: ' + (state.startDate || '') + ' - ' + (state.endDate || ''),
      'Author: ' + (state.author || state.authorQuery || ''),
      'Folder: ' + (state.folderPath || ''),
      ''
    ];
    for (const entry of entries) {
      out.push('[' + entry.ts + '] ' + entry.step + ' ' + JSON.stringify(entry.data || {}));
    }
    return out.join('\n') + '\n';
  }

  async function runExport(request) {
    stopRequested = false;
    const startedAt = new Date().toISOString();
    const log = [];
    const seen = new Map();

    const addLog = (step, data) => {
      log.push({ ts: new Date().toISOString(), step, data: data || {} });
    };

    let state = await patchState({
      status: 'running',
      message: 'Запускаю поиск сообщений в Teams...',
      tabId: request.tabId,
      authorQuery: request.authorQuery || '',
      author: '',
      startDate: request.startDate,
      endDate: request.endDate,
      pages: 0,
      messageCount: 0,
      conversationCount: 0,
      progress: 2,
      folderPath: '',
      filePath: '',
      logPath: '',
      error: '',
      startedAt,
      finishedAt: ''
    });

    try {
      await nativeMessage({ action: 'ping' });
      const folder = await nativeMessage({ action: 'createBatchFolder' });
      state = await patchState({ folderPath: folder.path });
      addLog('CREATE_FOLDER', { path: folder.path });

      if (stopRequested) throw new Error('Остановлено пользователем.');

      state = await patchState({
        message: 'Открываю Teams Search и применяю фильтры...',
        progress: 8
      });

      const prepared = await sendToTab(request.tabId, {
        type: 'CHAT_PREPARE_SEARCH',
        authorQuery: request.authorQuery || '',
        startDate: request.startDate,
        endDate: request.endDate
      });

      if (!prepared?.ok) {
        addLog('PREPARE_FAILED', prepared || {});
        throw new Error(prepared?.error || 'Не удалось подготовить Teams Search.');
      }

      state = await patchState({
        author: prepared.author || request.authorQuery || '',
        authorQuery: prepared.query || request.authorQuery || '',
        message: 'Фильтры применены. Собираю страницы результатов...',
        progress: 12
      });
      addLog('SEARCH_READY', prepared);

      for (let page = 1; page <= MAX_PAGES; page++) {
        if (stopRequested) throw new Error('Остановлено пользователем.');

        const result = await sendToTab(request.tabId, { type: 'CHAT_COLLECT_PAGE' });
        if (!result?.ok) {
          addLog('PAGE_READ_FAILED', { page, result });
          throw new Error(result?.error || 'Не удалось прочитать страницу результатов.');
        }

        let added = 0;
        for (const item of result.items || []) {
          const key = uniqueKey(item);
          if (!key || seen.has(key)) continue;
          seen.set(key, item);
          added++;
        }

        const conversations = new Set(Array.from(seen.values()).map(x => x.conversation).filter(Boolean));
        const progress = Math.min(92, 12 + page * 5);

        state = await patchState({
          pages: page,
          messageCount: seen.size,
          conversationCount: conversations.size,
          progress,
          message: 'Страница ' + page + ': сообщений ' + seen.size + ', чатов ' + conversations.size + '.'
        });

        addLog('PAGE_COLLECTED', {
          page,
          pageItems: result.count || 0,
          added,
          total: seen.size,
          hasNext: !!result.hasNext,
          signature: result.signature || ''
        });

        if (!result.hasNext) break;

        const moved = await sendToTab(request.tabId, { type: 'CHAT_GO_NEXT' });
        if (!moved?.ok) {
          addLog('NEXT_FAILED', { page, moved });
          throw new Error(moved?.error || 'Не удалось перейти на следующую страницу Teams Search.');
        }
        if (!moved.moved) break;
      }

      if (stopRequested) throw new Error('Остановлено пользователем.');

      const finishedAt = new Date().toISOString();
      state = await patchState({ finishedAt });
      const messages = Array.from(seen.values());
      const text = buildText(state, messages);
      const fileName = safeFileName('Teams messages - ' + state.startDate + '-' + state.endDate) + '.txt';

      const saved = await nativeMessage({
        action: 'saveTextInFolder',
        folderPath: state.folderPath,
        fileName,
        text
      });

      addLog('MESSAGES_SAVED', { path: saved.path, messages: messages.length });

      const finalState = await patchState({
        status: 'done',
        message: 'Готово. Собрано сообщений: ' + messages.length + '.',
        messageCount: messages.length,
        conversationCount: new Set(messages.map(x => x.conversation).filter(Boolean)).size,
        progress: 100,
        filePath: saved.path,
        finishedAt
      });

      const logText = buildLog(log, finalState);
      const logSaved = await nativeMessage({
        action: 'saveTextInFolder',
        folderPath: finalState.folderPath,
        fileName: 'chat-operation-log.txt',
        text: logText
      });

      await patchState({ logPath: logSaved.path });
    } catch (error) {
      const stopped = stopRequested || /Остановлено пользователем/i.test(error?.message || '');
      const finishedAt = new Date().toISOString();
      addLog(stopped ? 'STOPPED' : 'ERROR', { error: error?.message || String(error) });

      state = await patchState({
        status: stopped ? 'stopped' : 'error',
        message: stopped ? 'Сбор переписки остановлен.' : 'Ошибка: ' + (error?.message || error),
        error: stopped ? '' : (error?.message || String(error)),
        finishedAt,
        progress: stopped ? state.progress || 0 : state.progress || 0
      });

      if (state.folderPath) {
        try {
          const logSaved = await nativeMessage({
            action: 'saveTextInFolder',
            folderPath: state.folderPath,
            fileName: 'chat-operation-log.txt',
            text: buildLog(log, state)
          });
          await patchState({ logPath: logSaved.path });
        } catch (_) {}
      }
    } finally {
      runningPromise = null;
      stopRequested = false;
    }
  }

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    const type = message?.type;

    if (type === 'CHAT_GET_STATE') {
      getState().then(state => sendResponse({ ok: true, state }));
      return true;
    }

    if (type === 'CHAT_START') {
      if (runningPromise) {
        sendResponse({ ok: false, error: 'Сбор переписки уже выполняется.' });
        return;
      }

      if (!message.tabId || !message.startDate || !message.endDate) {
        sendResponse({ ok: false, error: 'Не переданы вкладка Teams или период.' });
        return;
      }

      runningPromise = runExport(message);
      sendResponse({ ok: true });
      return;
    }

    if (type === 'CHAT_STOP') {
      stopRequested = true;
      patchState({ message: 'Останавливаю после текущего шага...' })
        .then(state => sendResponse({ ok: true, state }));
      return true;
    }

    if (type === 'CHAT_OPEN_FOLDER') {
      getState().then(async state => {
        if (!state.folderPath) throw new Error('Папка результата еще не создана.');
        const response = await nativeMessage({ action: 'openDirectory', path: state.folderPath });
        sendResponse(response);
      }).catch(error => sendResponse({ ok: false, error: error.message || String(error) }));
      return true;
    }

    if (type === 'CHAT_RESET') {
      chrome.storage.local.set({ [CHAT_STATE_KEY]: blankState() })
        .then(() => sendResponse({ ok: true, state: blankState() }));
      return true;
    }
  });
})();