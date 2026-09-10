let currentState = null;
let currentTabId = null;
let pollTimer = null;
let nativeAvailable = false;
let lastSaved = null;

const NATIVE_HOST = 'com.openai.teams_recap_transcript_exporter';

const extractBtn = document.getElementById('extract');
const stopBtn = document.getElementById('stop');
const debugBtn = document.getElementById('debug');
const clearBtn = document.getElementById('clear');
const copyBtn = document.getElementById('copy');
const txtBtn = document.getElementById('txt');
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

function transcriptFileName() {
  const title = sanitizeFileName(currentState?.title || 'teams-transcript');
  return `${title} - ${localDateStamp()}.txt`;
}

function diagnosticFileName() {
  const title = sanitizeFileName(currentState?.title || 'teams-transcript');
  return `${title} - ${localDateStamp()} - diagnostic.txt`;
}

async function getActiveTab() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tabs.length) throw new Error('Не удалось определить активную вкладку.');
  return tabs[0];
}

async function sendToActiveTab(message) {
  const tab = await getActiveTab();
  currentTabId = tab.id;
  if (!/^https:\/\//i.test(tab.url || '')) {
    throw new Error('Открой Recap в обычной HTTPS-вкладке Chrome.');
  }
  try {
    return await chrome.tabs.sendMessage(tab.id, message);
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
    saveModeEl.textContent = 'Сохранение: Windows Documents. Кнопка "Открыть каталог" покажет сохраненный файл в Проводнике.';
  } else {
    saveModeEl.textContent = 'Сохранение: стандартная папка загрузок Chrome. Чтобы сохранять напрямую в Windows Documents, один раз запусти Install-Windows-Integration.cmd из папки расширения.';
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
  txtBtn.disabled = !hasText;
  showFolderBtn.disabled = !lastSaved;

  const progress = Number.isFinite(state?.progress) ? state.progress : 0;
  progressBar.style.width = `${Math.max(0, Math.min(100, progress))}%`;
  progressText.textContent = running
    ? `${progress}% - блоков: ${state.items || 0}`
    : done
      ? `100% - блоков: ${state.items || 0}, символов: ${state.chars || 0}`
      : '';

  statusEl.textContent = state?.message || 'Готово к работе.';

  if (hasText) {
    if (output.value !== state.text) output.value = state.text;
  } else if (state?.status === 'idle') {
    output.value = '';
  }
}

async function refreshState() {
  try {
    const response = await sendToActiveTab({ type: 'GET_STATE' });
    if (!response?.ok) throw new Error('Не удалось получить состояние сборщика.');
    render(response.state);
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

txtBtn.addEventListener('click', async () => {
  const text = currentState?.text || output.value || '';
  if (!text) return;
  try {
    const fileName = transcriptFileName();
    const saved = await saveText(text, fileName);
    statusEl.textContent = saved.mode === 'documents'
      ? `TXT сохранен: ${saved.path}`
      : `TXT сохранен в папку загрузок Chrome: ${fileName}`;
  } catch (e) {
    setLocalError(e.message || String(e));
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
  lastSaved = null;
  showFolderBtn.disabled = true;
  refreshState();
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo) => {
  if (tabId === currentTabId && changeInfo.status === 'complete') {
    lastSaved = null;
    showFolderBtn.disabled = true;
    setTimeout(refreshState, 400);
  }
});

(async () => {
  await detectNativeHost();
  await refreshState();
  startPolling();
})();
