(() => {
  const NATIVE_HOST = 'com.openai.teams_recap_transcript_exporter';
  const STATE_KEY = 'currentTranscriptBackgroundSaveStateV1';
  const SAVED_RUNS_KEY = 'currentTranscriptBackgroundSavedRunsV1';
  const inFlight = new Map();

  function sanitizeFileName(value) {
    return String(value || 'teams-transcript')
      .replace(/[\\/:*?"<>|]+/g, '-')
      .replace(/\s+/g, ' ')
      .replace(/[. ]+$/g, '')
      .trim()
      .slice(0, 150) || 'teams-transcript';
  }

  function stamp(value) {
    const date = new Date(Number(value) || Date.now());
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    const hh = String(date.getHours()).padStart(2, '0');
    const mi = String(date.getMinutes()).padStart(2, '0');
    const ss = String(date.getSeconds()).padStart(2, '0');
    return {
      date: `${yyyy}${mm}${dd}`,
      time: `${hh}${mi}${ss}`
    };
  }

  function runKey(tabId, frameId, result) {
    return [
      tabId,
      frameId,
      result?.startedAt || 0,
      result?.finishedAt || 0,
      result?.chars || String(result?.text || '').length,
      result?.title || ''
    ].join('|');
  }

  function fileNameFor(result) {
    const title = sanitizeFileName(result?.title || 'teams-transcript');
    const s = stamp(result?.startedAt || result?.finishedAt || Date.now());
    return `${title} - ${s.date} - ${s.time}.txt`;
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

  async function getStateMap() {
    const data = await chrome.storage.session.get(STATE_KEY);
    return data?.[STATE_KEY] || {};
  }

  async function setTabState(tabId, state) {
    const map = await getStateMap();
    map[String(tabId)] = state;
    await chrome.storage.session.set({ [STATE_KEY]: map });
    return state;
  }

  async function getTabState(tabId) {
    const map = await getStateMap();
    return map[String(tabId)] || {
      status: 'idle',
      tabId,
      frameId: null,
      runKey: '',
      mode: '',
      path: '',
      downloadId: null,
      fileName: '',
      error: '',
      savedAt: 0
    };
  }

  async function resetTabState(tabId) {
    return await setTabState(tabId, {
      status: 'idle',
      tabId,
      frameId: null,
      runKey: '',
      mode: '',
      path: '',
      downloadId: null,
      fileName: '',
      error: '',
      savedAt: 0
    });
  }

  async function savedRunKeys() {
    const data = await chrome.storage.session.get(SAVED_RUNS_KEY);
    return new Set(Array.isArray(data?.[SAVED_RUNS_KEY]) ? data[SAVED_RUNS_KEY] : []);
  }

  async function rememberRun(key) {
    const keys = await savedRunKeys();
    keys.add(key);
    await chrome.storage.session.set({
      [SAVED_RUNS_KEY]: Array.from(keys).slice(-100)
    });
  }

  async function downloadText(text, fileName) {
    const url = 'data:text/plain;charset=utf-8,' + encodeURIComponent(text);
    const downloadId = await chrome.downloads.download({
      url,
      filename: fileName,
      saveAs: false,
      conflictAction: 'uniquify'
    });

    return {
      mode: 'downloads',
      downloadId,
      fileName
    };
  }

  async function persistResult(tabId, frameId, result) {
    const key = runKey(tabId, frameId, result);
    const existing = await savedRunKeys();

    if (existing.has(key)) {
      return await getTabState(tabId);
    }

    if (inFlight.has(key)) return await inFlight.get(key);

    const task = (async () => {
      const fileName = fileNameFor(result);

      await setTabState(tabId, {
        status: 'saving',
        tabId,
        frameId,
        runKey: key,
        mode: '',
        path: '',
        downloadId: null,
        fileName,
        error: '',
        savedAt: 0
      });

      try {
        let saved;
        try {
          const response = await nativeMessage({
            action: 'saveText',
            fileName,
            text: String(result?.text || '')
          });

          saved = {
            mode: 'documents',
            path: response.path || '',
            downloadId: null,
            fileName
          };
        } catch (_) {
          saved = await downloadText(String(result?.text || ''), fileName);
        }

        await rememberRun(key);

        return await setTabState(tabId, {
          status: 'saved',
          tabId,
          frameId,
          runKey: key,
          mode: saved.mode,
          path: saved.path || '',
          downloadId: saved.downloadId ?? null,
          fileName,
          error: '',
          savedAt: Date.now()
        });
      } catch (error) {
        return await setTabState(tabId, {
          status: 'error',
          tabId,
          frameId,
          runKey: key,
          mode: '',
          path: '',
          downloadId: null,
          fileName,
          error: error?.message || String(error),
          savedAt: Date.now()
        });
      } finally {
        inFlight.delete(key);
      }
    })();

    inFlight.set(key, task);
    return await task;
  }

  async function openSaved(tabId) {
    const state = await getTabState(tabId);
    if (state.status !== 'saved') {
      throw new Error('Сохраненный транскрипт для этой вкладки не найден.');
    }

    if (state.mode === 'documents' && state.path) {
      return await nativeMessage({ action: 'showInFolder', path: state.path });
    }

    if (state.mode === 'downloads' && Number.isInteger(state.downloadId)) {
      chrome.downloads.show(state.downloadId);
      return { ok: true };
    }

    throw new Error('Не удалось определить сохраненный файл.');
  }

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    const type = message?.type;

    if (type === 'CURRENT_TRANSCRIPT_DONE') {
      const tabId = sender.tab?.id;
      const frameId = Number(sender.frameId || 0);
      const result = message.result || {};

      if (!tabId) {
        sendResponse({ ok: false, error: 'Не удалось определить вкладку завершенного Transcript.' });
        return;
      }

      if (!result.text || result.status !== 'done') {
        sendResponse({ ok: false, error: 'Завершенный Transcript не содержит текста.' });
        return;
      }

      persistResult(tabId, frameId, result)
        .then(state => sendResponse({ ok: state.status === 'saved', state, error: state.error || '' }))
        .catch(error => sendResponse({ ok: false, error: error.message || String(error) }));
      return true;
    }

    if (type === 'CURRENT_TRANSCRIPT_GET_SAVE_STATE') {
      const tabId = Number(message.tabId || sender.tab?.id || 0);
      if (!tabId) {
        sendResponse({ ok: false, error: 'Не передана вкладка.' });
        return;
      }

      getTabState(tabId)
        .then(state => sendResponse({ ok: true, state }))
        .catch(error => sendResponse({ ok: false, error: error.message || String(error) }));
      return true;
    }

    if (type === 'CURRENT_TRANSCRIPT_RESET_SAVE_STATE') {
      const tabId = Number(message.tabId || sender.tab?.id || 0);
      if (!tabId) {
        sendResponse({ ok: false, error: 'Не передана вкладка.' });
        return;
      }

      resetTabState(tabId)
        .then(state => sendResponse({ ok: true, state }))
        .catch(error => sendResponse({ ok: false, error: error.message || String(error) }));
      return true;
    }

    if (type === 'CURRENT_TRANSCRIPT_OPEN_SAVED') {
      const tabId = Number(message.tabId || sender.tab?.id || 0);
      if (!tabId) {
        sendResponse({ ok: false, error: 'Не передана вкладка.' });
        return;
      }

      openSaved(tabId)
        .then(response => sendResponse({ ok: true, response }))
        .catch(error => sendResponse({ ok: false, error: error.message || String(error) }));
      return true;
    }
  });

  chrome.tabs.onRemoved.addListener(async tabId => {
    try {
      const map = await getStateMap();
      if (!(String(tabId) in map)) return;
      delete map[String(tabId)];
      await chrome.storage.session.set({ [STATE_KEY]: map });
    } catch (_) {}
  });
})();
