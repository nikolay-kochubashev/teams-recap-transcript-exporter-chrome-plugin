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
      searchHitCount: 0,
      messageCount: 0,
      conversationCount: 0,
      contextConversationCount: 0,
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
      throw new Error(
        'Не удалось связаться со вкладкой Teams. Обнови вкладку после обновления расширения. ' +
        (e.message || e)
      );
    }
  }

  function stampDisplay(stamp) {
    if (!/^\d{8}$/.test(stamp || '')) return stamp || '';
    return stamp.slice(6, 8) + '.' + stamp.slice(4, 6) + '.' + stamp.slice(0, 4);
  }

  function safeFileName(value) {
    return String(value || 'Teams chats')
      .replace(/[\\/:*?"<>|]+/g, '-')
      .replace(/\s+/g, ' ')
      .replace(/[. ]+$/g, '')
      .trim()
      .slice(0, 170) || 'Teams chats';
  }

  function uniqueKey(item) {
    return item?.id || item?.key || [
      item?.conversation || '',
      item?.sender || '',
      item?.timestamp || '',
      item?.message || ''
    ].join('|');
  }

  function conversationKey(value) {
    return String(value || '').replace(/\s+/g, ' ').trim().toLocaleLowerCase();
  }

  function addLinks(out, links) {
    const unique = Array.from(new Set(links || [])).filter(Boolean);
    if (!unique.length) return;
    out.push('Links:');
    unique.forEach(link => out.push('- ' + link));
  }

  function buildText(state, searchHits, contexts, contextFailures) {
    const hitsByConversation = new Map();
    for (const hit of searchHits) {
      const name = String(hit.conversation || '').trim() || '(чат не определен)';
      const key = conversationKey(name);
      if (!hitsByConversation.has(key)) hitsByConversation.set(key, { name, items: [] });
      hitsByConversation.get(key).items.push(hit);
    }

    // Search is discovery-only. It tells us which conversations are relevant,
    // but Search rows themselves are never exported as conversation content.
    const allKeys = new Set([
      ...hitsByConversation.keys(),
      ...contexts.keys(),
      ...contextFailures.keys()
    ]);

    const groups = Array.from(allKeys).map(key => {
      const fallback = hitsByConversation.get(key);
      const context = contexts.get(key);
      return {
        key,
        name: context?.conversation || fallback?.name || key || '(чат не определен)',
        context,
        fallback: fallback?.items || [],
        failure: contextFailures.get(key) || ''
      };
    }).sort((a, b) => a.name.localeCompare(b.name, 'ru'));

    let savedMessages = 0;
    for (const group of groups) {
      savedMessages += group.context?.messages?.length || 0;
    }

    const out = [
      'Teams chats for work report',
      'Period: ' + stampDisplay(state.startDate) + ' - ' + stampDisplay(state.endDate),
      'Author used for discovery: ' + (state.author || state.authorQuery || ''),
      'Search hits from author: ' + searchHits.length,
      'Conversations: ' + groups.length,
      'Full-context conversations: ' + contexts.size,
      'Conversations not opened: ' + contextFailures.size,
      'Messages saved: ' + savedMessages,
      '',
      'Conversations:',
      ...groups.map(x => '- ' + x.name),
      ''
    ];

    for (const group of groups) {
      out.push('============================================================');
      out.push(group.name);
      out.push('============================================================');

      if (group.context?.messages?.length) {
        out.push('Source: full Teams conversation for selected period');
        out.push('');

        for (const message of group.context.messages) {
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
      } else {
        out.push('Source: conversation context was not collected.');
        if (group.failure) out.push('Reason: ' + group.failure);
        out.push('Search result content is intentionally not exported because it does not contain the full conversation context.');
        out.push('');
      }
    }

    return {
      text: out.join('\n').trim() + '\n',
      savedMessages,
      conversationCount: groups.length
    };
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
    const seenSearch = new Map();
    const contexts = new Map();
    const contextFailures = new Map();
    const attemptedContexts = new Set();

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
      searchHitCount: 0,
      messageCount: 0,
      conversationCount: 0,
      contextConversationCount: 0,
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
        message: 'Фильтры применены. Search определяет релевантные чаты, затем читаю сообщения из самих чатов...',
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
          if (!key || seenSearch.has(key)) continue;
          seenSearch.set(key, item);
          added++;
        }

        const knownConversations = new Set(
          Array.from(seenSearch.values())
            .map(x => String(x.conversation || '').trim())
            .filter(Boolean)
            .map(conversationKey)
        );

        state = await patchState({
          pages: page,
          searchHitCount: seenSearch.size,
          conversationCount: knownConversations.size,
          contextConversationCount: contexts.size,
          messageCount: Array.from(contexts.values())
            .reduce((sum, x) => sum + (x.messages?.length || 0), 0),
          progress: Math.min(92, 12 + page * 5),
          message:
            'Страница поиска ' + page +
            ': найдено моих сообщений ' + seenSearch.size +
            ', релевантных чатов ' + knownConversations.size +
            ', чатов прочитано ' + contexts.size + '.'
        });

        addLog('PAGE_COLLECTED', {
          page,
          pageItems: result.count || 0,
          added,
          total: seenSearch.size,
          hasNext: !!result.hasNext,
          signature: result.signature || ''
        });

        const pageConversations = new Map();
        for (const item of result.items || []) {
          const name = String(item.conversation || '').trim();
          if (!name) continue;
          const key = conversationKey(name);
          if (!pageConversations.has(key)) pageConversations.set(key, item);
        }

        for (const [key, item] of pageConversations.entries()) {
          if (stopRequested) throw new Error('Остановлено пользователем.');
          if (attemptedContexts.has(key)) continue;
          attemptedContexts.add(key);

          await patchState({
            message: 'Открываю переписку: ' + item.conversation,
            contextConversationCount: contexts.size
          });

          const context = await sendToTab(request.tabId, {
            type: 'CHAT_COLLECT_CONVERSATION_CONTEXT',
            item,
            startDate: request.startDate,
            endDate: request.endDate
          });

          if (context?.ok) {
            const normalized = Object.assign({}, context, {
              conversation: context.conversation || item.conversation
            });
            contexts.set(key, normalized);

            addLog('CONVERSATION_CONTEXT_COLLECTED', {
              conversation: normalized.conversation,
              messages: normalized.count || normalized.messages?.length || 0,
              loaded: normalized.loaded || 0,
              reachedStart: !!normalized.reachedStart,
              reachedEnd: !!normalized.reachedEnd
            });

            await patchState({
              contextConversationCount: contexts.size,
              messageCount: Array.from(contexts.values())
                .reduce((sum, x) => sum + (x.messages?.length || 0), 0),
              message:
                'Контекст: ' + normalized.conversation +
                ' - сообщений за период ' + (normalized.count || normalized.messages?.length || 0) + '.'
            });
          } else {
            const reason = context?.error || 'Контекст переписки не получен.';
            contextFailures.set(key, reason);
            addLog('CONVERSATION_CONTEXT_FAILED', {
              conversation: item.conversation,
              error: reason,
              diagnostic: context?.diagnostic || ''
            });
          }
        }

        if (!result.hasNext) break;

        const moved = await sendToTab(request.tabId, { type: 'CHAT_GO_NEXT' });
        if (!moved?.ok) {
          addLog('NEXT_FAILED', { page, moved });
          throw new Error(moved?.error || 'Не удалось перейти на следующую страницу Teams Search.');
        }
        if (!moved.moved) break;
      }

      if (stopRequested) throw new Error('Остановлено пользователем.');

      const searchHits = Array.from(seenSearch.values());

      // Ensure every discovered named conversation has a result in either
      // full-context or fallback maps.
      for (const hit of searchHits) {
        const name = String(hit.conversation || '').trim();
        if (!name) continue;
        const key = conversationKey(name);
        if (!contexts.has(key) && !contextFailures.has(key)) {
          contextFailures.set(key, 'Чат не был открыт для чтения полного контекста.');
        }
      }

      const finishedAt = new Date().toISOString();
      state = await patchState({ finishedAt });

      const built = buildText(state, searchHits, contexts, contextFailures);
      const fileName = safeFileName(
        'Teams chats - ' + state.startDate + '-' + state.endDate
      ) + '.txt';

      const saved = await nativeMessage({
        action: 'saveTextInFolder',
        folderPath: state.folderPath,
        fileName,
        text: built.text
      });

      addLog('CHATS_SAVED', {
        path: saved.path,
        searchHits: searchHits.length,
        conversations: built.conversationCount,
        fullContextConversations: contexts.size,
        fallbackConversations: contextFailures.size,
        messages: built.savedMessages
      });

      const finalState = await patchState({
        status: 'done',
        message:
          'Готово. Переписок: ' + built.conversationCount +
          ', сообщений с контекстом: ' + built.savedMessages + '.',
        searchHitCount: searchHits.length,
        messageCount: built.savedMessages,
        conversationCount: built.conversationCount,
        contextConversationCount: contexts.size,
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
        progress: state.progress || 0
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