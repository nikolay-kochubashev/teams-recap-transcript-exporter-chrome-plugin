let currentState = null;
let currentTabId = null;
let currentFrameId = null;
let pollTimer = null;
let nativeAvailable = false;
let lastSaved = null;
let autoSaveInFlightKey = '';
let autoSaveAttemptedRunKey = '';
let autoSaveCompletedRunKey = '';
let autoSaveMessage = '';

const AUTO_SAVE_STORAGE_KEY = 'currentTranscriptAutoSavedRunKeysV1';

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

function transcriptFileName(state = currentState) {
  const title = sanitizeFileName(state?.title || 'teams-transcript');
  return `${title} - ${localDateStamp()}.txt`;
}

function autoSaveRunKey(state) {
  if (!state || state.status !== 'done' || !state.text || !state.finishedAt) return '';
  return [
    currentTabId || 0,
    currentFrameId ?? 'top',
    state.startedAt || 0,
    state.finishedAt || 0,
    state.chars || state.text.length || 0,
    state.title || ''
  ].join('|');
}

async function loadAutoSavedRunKeys() {
  try {
    const data = await chrome.storage.session.get(AUTO_SAVE_STORAGE_KEY);
    const values = Array.isArray(data?.[AUTO_SAVE_STORAGE_KEY])
      ? data[AUTO_SAVE_STORAGE_KEY]
      : [];
    return new Set(values);
  } catch (_) {
    return new Set();
  }
}

async function rememberAutoSavedRunKey(runKey) {
  if (!runKey) return;
  try {
    const keys = await loadAutoSavedRunKeys();
    keys.add(runKey);
    const values = Array.from(keys).slice(-50);
    await chrome.storage.session.set({ [AUTO_SAVE_STORAGE_KEY]: values });
  } catch (_) {}
}

function diagnosticFileName() {
  const title = sanitizeFileName(currentState?.title || 'teams-transcript');
  return `${title} - ${localDateStamp()} - diagnostic.txt`;
}

async function getActiveTab() {
  try {
    const response = await chrome.runtime.sendMessage({ type: 'SIDE_PANEL_GET_BOUND_TAB' });
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
      const probe = await chrome.tabs.sendMessage(tabId, { type: 'PROBE_TRANSCRIPT' }, { frameId: frame.frameId });
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
  if (currentTabId !== tab.id) currentFrameId = null;
  currentTabId = tab.id;

  if (!/^https:\/\//i.test(tab.url || '')) {
    throw new Error('Открой Recap в обычной HTTPS-вкладке Chrome.');
  }

  try {
    if (message?.type === 'START_EXTRACT' || message?.type === 'GET_DEBUG') {
      currentFrameId = await findTranscriptFrame(tab.id);
    }

    if (currentFrameId !== null) {
      return await chrome.tabs.sendMessage(tab.id, message, { frameId: currentFrameId });
    }

    return await chrome.tabs.sendMessage(tab.id, message, { frameId: 0 });
  } catch (e) {
    throw new Error('Эта страница не поддерживается расширением. Открой Teams/SharePoint Recap и обнови вкладку после установки новой версии.');
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
    saveModeEl.textContent = 'После завершения транскрипт автоматически сохраняется в Windows Documents. Кнопка "Открыть каталог" покажет сохраненный файл в Проводнике.';
  } else {
    saveModeEl.textContent = 'После завершения транскрипт автоматически сохраняется в стандартную папку загрузок Chrome. Чтобы сохранять напрямую в Windows Documents, один раз запусти Install-Windows-Integration.cmd из папки расширения.';
  }
}

function setLocalError(message) {
  statusEl.textContent = `Ошибка: ${message}`;
  progressText.textContent = '';
}

function render(state) {
  currentState = state;
  const running = state?.status === 'running';
  const done = state?.status === 'done';
  const hasText = !!state?.text;

  extractBtn.disabled = running;
  stopBtn.disabled = !running;
  debugBtn.disabled = running;
  clearBtn.disabled = running;
  copyBtn.disabled = !hasText;
  showFolderBtn.disabled = !lastSaved;

  const progress = Number.isFinite(state?.progress) ? state.progress : 0;
  progressBar.style.width = `${Math.max(0, Math.min(100, progress))}%`;
  progressText.textContent = running
    ? `${progress}% - блоков: ${state.items || 0}`
    : done
      ? `100% - блоков: ${state.items || 0}, символов: ${state.chars || 0}`
      : '';

  const runKey = autoSaveRunKey(state);
  if (runKey && runKey === autoSaveInFlightKey) {
    statusEl.textContent = 'Транскрипция собрана. Автоматически сохраняю TXT...';
  } else if (runKey && runKey === autoSaveCompletedRunKey && autoSaveMessage) {
    statusEl.textContent = autoSaveMessage;
  } else {
    statusEl.textContent = state?.message || 'Готово к работе.';
  }

  if (hasText) {
    if (output.value !== state.text) output.value = state.text;
  } else if (state?.status === 'idle') {
    output.value = '';
  }
}

async function maybeAutoSaveTranscript(state) {
  const runKey = autoSaveRunKey(state);
  if (!runKey) return;
  if (runKey === autoSaveInFlightKey || runKey === autoSaveAttemptedRunKey) return;

  const alreadySaved = await loadAutoSavedRunKeys();
  if (alreadySaved.has(runKey)) {
    autoSaveAttemptedRunKey = runKey;
    autoSaveCompletedRunKey = runKey;
    autoSaveMessage = 'Готово. TXT уже был автоматически сохранен для этого запуска.';
    render(state);
    return;
  }

  autoSaveAttemptedRunKey = runKey;
  autoSaveInFlightKey = runKey;
  autoSaveMessage = '';
  render(state);

  try {
    const fileName = transcriptFileName(state);
    const saved = await saveText(state.text, fileName);
    await rememberAutoSavedRunKey(runKey);

    autoSaveCompletedRunKey = runKey;
    autoSaveMessage = saved.mode === 'documents'
      ? `Готово. TXT автоматически сохранен: ${saved.path}`
      : `Готово. TXT автоматически сохранен в папку загрузок Chrome: ${fileName}`;
  } catch (e) {
    autoSaveCompletedRunKey = runKey;
    autoSaveMessage = `Ошибка автоматического сохранения TXT: ${e.message || e}`;
  } finally {
    autoSaveInFlightKey = '';
    render(state);
  }
}

async function refreshState() {
  try {
    const response = await sendToActiveTab({ type: 'GET_STATE' });
    if (!response?.ok) throw new Error('Не удалось получить состояние сборщика.');
    render(response.state);
    await maybeAutoSaveTranscript(response.state);
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
      chrome.downloads.download({ url, filename: fileName, saveAs: false, conflictAction: 'overwrite' }, downloadId => {
        const err = chrome.runtime.lastError;
        if (err) return reject(new Error(err.message));
        resolve(downloadId);
      });
    });
    lastSaved = { kind: 'download', id, fileName };
    showFolderBtn.disabled = false;
    return { mode: 'downloads', fileName };
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }
}

async function saveText(text, fileName) {
  if (nativeAvailable) {
    try {
      const response = await sendNative({ action: 'saveText', fileName, text });
      if (!response?.ok) throw new Error(response?.error || 'Не удалось сохранить файл.');
      lastSaved = { kind: 'native', path: response.path, fileName };
      showFolderBtn.disabled = false;
      return { mode: 'documents', path: response.path };
    } catch (e) {
      nativeAvailable = false;
      renderSaveMode();
      statusEl.textContent = `Windows-интеграция недоступна (${e.message || e}). Сохраняю в стандартную папку загрузок Chrome.`;
    }
  }
  return await downloadViaChrome(text, fileName);
}

extractBtn.addEventListener('click', async () => {
  try {
    lastSaved = null;
    autoSaveInFlightKey = '';
    autoSaveAttemptedRunKey = '';
    autoSaveCompletedRunKey = '';
    autoSaveMessage = '';
    showFolderBtn.disabled = true;

    const response = await sendToActiveTab({ type: 'START_EXTRACT' });
    if (response?.state) render(response.state);
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
    lastSaved = null;
    autoSaveInFlightKey = '';
    autoSaveAttemptedRunKey = '';
    autoSaveCompletedRunKey = '';
    autoSaveMessage = '';
    showFolderBtn.disabled = true;

    const response = await sendToActiveTab({ type: 'CLEAR_RESULT' });
    if (response?.state) render(response.state);
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
    const saved = await saveText(debugText, fileName);
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
  if (!lastSaved) return;
  try {
    if (lastSaved.kind === 'native') {
      const response = await sendNative({ action: 'showInFolder', path: lastSaved.path });
      if (!response?.ok) throw new Error(response?.error || 'Не удалось открыть Проводник.');
    } else {
      chrome.downloads.show(lastSaved.id);
    }
  } catch (e) {
    setLocalError(e.message || String(e));
  }
});

chrome.tabs.onActivated.addListener(() => {
  currentFrameId = null;
  lastSaved = null;
  autoSaveInFlightKey = '';
  autoSaveAttemptedRunKey = '';
  autoSaveCompletedRunKey = '';
  autoSaveMessage = '';
  showFolderBtn.disabled = true;
  refreshState();
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (tabId === currentTabId && changeInfo.status === 'complete') {
    currentFrameId = null;
    lastSaved = null;
    autoSaveInFlightKey = '';
    autoSaveAttemptedRunKey = '';
    autoSaveCompletedRunKey = '';
    autoSaveMessage = '';
    showFolderBtn.disabled = true;
    setTimeout(refreshState, 400);
  }
});

(async () => {
  await detectNativeHost();
  await refreshState();
  startPolling();
})();
