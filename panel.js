let currentState = null;
let currentTabId = null;
let currentFrameId = null;
let currentSaveState = null;
let pollTimer = null;
let nativeAvailable = false;

const NATIVE_HOST = 'com.openai.teams_recap_transcript_exporter';

const extractBtn = document.getElementById('extract');
const stopBtn = document.getElementById('stop');
const debugBtn = document.getElementById('debug');
const clearBtn = document.getElementById('clear');
const copyBtn = document.getElementById('copy');
const showFolderBtn = document.getElementById('showFolder');
const output = document.getElementById('output');
const statusEl = document.getElementById('status');
const saveModeEl = document.getElementById('saveMode');
const progressBar = document.getElementById('progressBar');
const progressText = document.getElementById('progressText');

function sanitizeFileName(value) {
  return (value || 'teams-transcript')
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/[. ]+$/g, '')
    .trim()
    .slice(0, 150) || 'teams-transcript';
}

function localDateStamp() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}${mm}${dd}`;
}

function diagnosticFileName() {
  const title = sanitizeFileName(currentState?.title || 'teams-transcript');
  return `${title} - ${localDateStamp()} - diagnostic.txt`;
}

async function runtimeMessage(message) {
  return await chrome.runtime.sendMessage(message);
}

async function getActiveTab() {
  try {
    const response = await runtimeMessage({ type: 'SIDE_PANEL_GET_BOUND_TAB' });
    if (response?.ok && response?.tab?.id) return response.tab;
  } catch (_) {}

  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tabs.length) throw new Error('Не удалось определить вкладку Side Panel.');
  return tabs[0];
}

async function findTranscriptFrame(tabId) {
  let frames = [];
  try {
    frames = await chrome.webNavigation.getAllFrames({ tabId }) || [];
  } catch (_) {
    frames = [{ frameId: 0 }];
  }

  const probes = [];
  for (const frame of frames) {
    try {
      const probe = await chrome.tabs.sendMessage(
        tabId,
        { type: 'PROBE_TRANSCRIPT' },
        { frameId: frame.frameId }
      );
      if (probe?.ok) probes.push({ frameId: frame.frameId, ...probe });
    } catch (_) {}
  }

  probes.sort((a, b) => {
    const as = a.best ? (a.best.strong ? 10000 : 0) + (a.best.score || 0) : 0;
    const bs = b.best ? (b.best.strong ? 10000 : 0) + (b.best.score || 0) : 0;
    return bs - as;
  });

  const selected = probes.find(x =>
    x.best && (
      x.best.strong ||
      (x.best.score || 0) >= 45 ||
      ((x.best.votes || 0) > 0 && (x.best.times || 0) > 0)
    )
  );

  return selected?.frameId ?? null;
}

async function sendToActiveTab(message) {
  const tab = await getActiveTab();
  if (currentTabId !== tab.id) {
    currentTabId = tab.id;
    currentFrameId = null;
    currentSaveState = null;
  }

  if (!/^https:\/\//i.test(tab.url || '')) {
    throw new Error('Открой Recap в обычной HTTPS-вкладке Chrome.');
  }

  try {
    if (
      message?.type === 'START_EXTRACT' ||
      message?.type === 'GET_DEBUG' ||
      (message?.type === 'GET_STATE' && currentFrameId === null)
    ) {
      currentFrameId = await findTranscriptFrame(tab.id);
    }

    if (currentFrameId !== null) {
      return await chrome.tabs.sendMessage(tab.id, message, { frameId: currentFrameId });
    }

    return await chrome.tabs.sendMessage(tab.id, message, { frameId: 0 });
  } catch (_) {
    throw new Error(
      'Эта страница не поддерживается расширением. Открой Teams/SharePoint Recap и обнови вкладку после установки новой версии.'
    );
  }
}

function sendNative(message) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendNativeMessage(NATIVE_HOST, message, response => {
      const err = chrome.runtime.lastError;
      if (err) return reject(new Error(err.message));
      if (!response) return reject(new Error('Windows helper не вернул ответ.'));
      resolve(response);
    });
  });
}

async function detectNativeHost() {
  try {
    const response = await sendNative({ action: 'ping' });
    nativeAvailable = !!response?.ok;
  } catch (_) {
    nativeAvailable = false;
  }
  renderSaveMode();
}

function renderSaveMode() {
  if (nativeAvailable) {
    saveModeEl.textContent =
      'После завершения каждая вкладка сохраняет TXT независимо через background в Windows Documents. Side Panel для сохранения не требуется.';
  } else {
    saveModeEl.textContent =
      'После завершения каждая вкладка сохраняет TXT независимо через background. Если Windows helper недоступен, используется стандартная папка загрузок Chrome.';
  }
}

function setLocalError(message) {
  statusEl.textContent = `Ошибка: ${message}`;
  progressText.textContent = '';
}

function saveStatusText(state, saveState) {
  if (state?.status !== 'done') return state?.message || 'Готово к работе.';

  if (saveState?.status === 'saving') {
    return 'Транскрипция собрана. Background сохраняет TXT...';
  }

  if (saveState?.status === 'saved') {
    if (saveState.mode === 'documents') {
      return `Готово. TXT автоматически сохранен: ${saveState.path}`;
    }
    return `Готово. TXT автоматически сохранен в загрузки Chrome: ${saveState.fileName}`;
  }

  if (saveState?.status === 'error') {
    return `Транскрипция собрана, но TXT сохранить не удалось: ${saveState.error}`;
  }

  return state?.message || 'Готово к работе.';
}

function render(state, saveState = currentSaveState) {
  currentState = state;
  currentSaveState = saveState || null;

  const running = state?.status === 'running';
  const done = state?.status === 'done';
  const hasText = !!state?.text;

  extractBtn.disabled = running;
  stopBtn.disabled = !running;
  debugBtn.disabled = running;
  clearBtn.disabled = running;
  copyBtn.disabled = !hasText;
  showFolderBtn.disabled = currentSaveState?.status !== 'saved';

  const progress = Number.isFinite(state?.progress) ? state.progress : 0;
  progressBar.style.width = `${Math.max(0, Math.min(100, progress))}%`;
  progressText.textContent = running
    ? `${progress}% - блоков: ${state.items || 0}`
    : done
      ? `100% - блоков: ${state.items || 0}, символов: ${state.chars || 0}`
      : '';

  statusEl.textContent = saveStatusText(state, currentSaveState);

  if (hasText) {
    if (output.value !== state.text) output.value = state.text;
  } else if (state?.status === 'idle') {
    output.value = '';
  }
}

async function getBackgroundSaveState(tabId) {
  if (!tabId) return null;
  try {
    const response = await runtimeMessage({
      type: 'CURRENT_TRANSCRIPT_GET_SAVE_STATE',
      tabId
    });
    return response?.ok ? response.state : null;
  } catch (_) {
    return null;
  }
}

async function refreshState() {
  try {
    const response = await sendToActiveTab({ type: 'GET_STATE' });
    if (!response?.ok) throw new Error('Не удалось получить состояние сборщика.');

    const saveState = await getBackgroundSaveState(currentTabId);
    render(response.state, saveState);
  } catch (e) {
    setLocalError(e.message || String(e));
  }
}

function startPolling() {
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(refreshState, 700);
}

async function downloadViaChrome(text, fileName) {
  const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  try {
    const id = await new Promise((resolve, reject) => {
      chrome.downloads.download({
        url,
        filename: fileName,
        saveAs: false,
        conflictAction: 'uniquify'
      }, downloadId => {
        const err = chrome.runtime.lastError;
        if (err) return reject(new Error(err.message));
        resolve(downloadId);
      });
    });
    return { mode: 'downloads', id, fileName };
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }
}

async function saveDiagnosticText(text, fileName) {
  if (nativeAvailable) {
    try {
      const response = await sendNative({ action: 'saveText', fileName, text });
      if (!response?.ok) throw new Error(response?.error || 'Не удалось сохранить файл.');
      return { mode: 'documents', path: response.path };
    } catch (_) {
      nativeAvailable = false;
      renderSaveMode();
    }
  }

  return await downloadViaChrome(text, fileName);
}

extractBtn.addEventListener('click', async () => {
  try {
    const tab = await getActiveTab();
    currentTabId = tab.id;
    currentFrameId = null;
    currentSaveState = null;

    await runtimeMessage({
      type: 'CURRENT_TRANSCRIPT_RESET_SAVE_STATE',
      tabId: tab.id
    });

    const response = await sendToActiveTab({ type: 'START_EXTRACT' });
    if (response?.state) render(response.state, null);
    startPolling();
  } catch (e) {
    setLocalError(e.message || String(e));
  }
});

stopBtn.addEventListener('click', async () => {
  try {
    await sendToActiveTab({ type: 'STOP_EXTRACT' });
    await refreshState();
  } catch (e) {
    setLocalError(e.message || String(e));
  }
});

clearBtn.addEventListener('click', async () => {
  try {
    if (currentTabId) {
      await runtimeMessage({
        type: 'CURRENT_TRANSCRIPT_RESET_SAVE_STATE',
        tabId: currentTabId
      });
    }
    currentSaveState = null;

    const response = await sendToActiveTab({ type: 'CLEAR_RESULT' });
    if (response?.state) render(response.state, null);
  } catch (e) {
    setLocalError(e.message || String(e));
  }
});

debugBtn.addEventListener('click', async () => {
  try {
    statusEl.textContent = 'Собираю диагностику DOM...';
    const response = await sendToActiveTab({ type: 'GET_DEBUG' });
    if (!response?.ok) throw new Error('Не удалось собрать диагностику.');

    const debugText = response.text || '';
    if (!debugText) throw new Error('Диагностика получилась пустой.');

    const fileName = diagnosticFileName();
    const saved = await saveDiagnosticText(debugText, fileName);
    statusEl.textContent = saved.mode === 'documents'
      ? `Диагностика сохранена: ${saved.path}`
      : `Диагностика сохранена в папку загрузок Chrome: ${fileName}`;
  } catch (e) {
    setLocalError(e.message || String(e));
  }
});

copyBtn.addEventListener('click', async () => {
  const text = currentState?.text || output.value || '';
  if (!text) return;

  try {
    await navigator.clipboard.writeText(text);
    statusEl.textContent = 'Текст скопирован в буфер обмена.';
  } catch (e) {
    setLocalError(`Не удалось скопировать: ${e.message || e}`);
  }
});

showFolderBtn.addEventListener('click', async () => {
  try {
    const tab = await getActiveTab();
    const response = await runtimeMessage({
      type: 'CURRENT_TRANSCRIPT_OPEN_SAVED',
      tabId: tab.id
    });

    if (!response?.ok) {
      throw new Error(response?.error || 'Не удалось открыть сохраненный TXT.');
    }
  } catch (e) {
    setLocalError(e.message || String(e));
  }
});

chrome.tabs.onActivated.addListener(() => {
  currentFrameId = null;
  currentSaveState = null;
  showFolderBtn.disabled = true;
  refreshState();
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (tabId === currentTabId && changeInfo.status === 'complete') {
    currentFrameId = null;
    currentSaveState = null;
    showFolderBtn.disabled = true;
    setTimeout(refreshState, 400);
  }
});

(async () => {
  await detectNativeHost();
  await refreshState();
  startPolling();
})();
