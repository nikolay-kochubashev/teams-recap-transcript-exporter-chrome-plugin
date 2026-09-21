(() => {
  const BATCH_NATIVE_HOST = 'com.openai.teams_recap_transcript_exporter';
  let batchNativeAvailable = false;
  let batchState = null;
  let batchPollTimer = null;

  const currentSection = document.getElementById('currentSection');
  const batchSection = document.getElementById('batchSection');
  const modeCurrentBtn = document.getElementById('modeCurrent');
  const modeBatchBtn = document.getElementById('modeBatch');

  const scanCalendarBtn = document.getElementById('scanCalendar');
  const calendarDiagnosticBtn = document.getElementById('calendarDiagnostic');
  const batchStatusEl = document.getElementById('batchStatus');
  const batchSaveModeEl = document.getElementById('batchSaveMode');
  const selectAllEl = document.getElementById('selectAll');
  const meetingCountEl = document.getElementById('meetingCount');
  const meetingListEl = document.getElementById('meetingList');
  const startBatchBtn = document.getElementById('startBatch');
  const stopBatchBtn = document.getElementById('stopBatch');
  const openBatchFolderBtn = document.getElementById('openBatchFolder');
  const batchProgressBar = document.getElementById('batchProgressBar');
  const batchProgressText = document.getElementById('batchProgressText');
  const batchLogEl = document.getElementById('batchLog');

  function runtimeMessage(message) {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendMessage(message, response => {
        const err = chrome.runtime.lastError;
        if (err) return reject(new Error(err.message));
        resolve(response);
      });
    });
  }

  function nativeMessage(message) {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendNativeMessage(BATCH_NATIVE_HOST, message, response => {
        const err = chrome.runtime.lastError;
        if (err) return reject(new Error(err.message));
        if (!response) return reject(new Error('Windows helper не вернул ответ.'));
        resolve(response);
      });
    });
  }

  async function activeTab() {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tabs.length) throw new Error('Не удалось определить активную вкладку.');
    return tabs[0];
  }

  function sanitizeFileName(value) {
    return (value || 'Teams Calendar')
      .replace(/[\\/:*?"<>|]+/g, '-')
      .replace(/\s+/g, ' ')
      .replace(/[. ]+$/g, '')
      .trim()
      .slice(0, 150) || 'Teams Calendar';
  }

  function dateStamp() {
    const d = new Date();
    return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
  }

  async function saveDiagnostic(text) {
    const fileName = `${sanitizeFileName('Teams Calendar')} - ${dateStamp()} - diagnostic.txt`;
    if (batchNativeAvailable) {
      const response = await nativeMessage({ action: 'saveText', fileName, text });
      if (!response?.ok) throw new Error(response?.error || 'Не удалось сохранить диагностику.');
      return response.path;
    }

    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    try {
      await new Promise((resolve, reject) => {
        chrome.downloads.download({ url, filename: fileName, saveAs: false }, id => {
          const err = chrome.runtime.lastError;
          if (err) reject(new Error(err.message));
          else resolve(id);
        });
      });
    } finally {
      setTimeout(() => URL.revokeObjectURL(url), 30000);
    }
    return fileName;
  }

  function setMode(mode) {
    const batch = mode === 'batch';
    currentSection.hidden = batch;
    batchSection.hidden = !batch;
    modeCurrentBtn.classList.toggle('active', !batch);
    modeBatchBtn.classList.toggle('active', batch);
    if (batch) refreshBatchState();
  }

  modeCurrentBtn.addEventListener('click', () => setMode('current'));
  modeBatchBtn.addEventListener('click', () => setMode('batch'));

  function selectedMeetingIds() {
    return Array.from(meetingListEl.querySelectorAll('input[data-meeting-id]:checked')).map(x => x.dataset.meetingId);
  }

  function syncSelectAll() {
    const all = Array.from(meetingListEl.querySelectorAll('input[data-meeting-id]'));
    const checked = all.filter(x => x.checked);
    selectAllEl.checked = all.length > 0 && checked.length === all.length;
    selectAllEl.indeterminate = checked.length > 0 && checked.length < all.length;
  }

  function statusLabel(status) {
    const map = {
      pending: 'Ожидает',
      running: 'В работе',
      done: 'Готово',
      error: 'Ошибка',
      skip: 'Нет записи'
    };
    return map[status] || status || 'Ожидает';
  }

  function renderMeetingList(state) {
    const meetings = state?.meetings || [];
    const logs = new Map((state?.logs || []).map(x => [x.meetingId, x]));
    const oldSelected = new Set(selectedMeetingIds());
    const selected = oldSelected.size ? oldSelected : new Set(state?.selectedIds || meetings.map(m => m.id));

    meetingCountEl.textContent = meetings.length ? `${meetings.length} встреч` : '';
    if (!meetings.length) {
      meetingListEl.innerHTML = '<div class="empty">Встречи пока не считаны.</div>';
      startBatchBtn.disabled = true;
      return;
    }

    meetingListEl.innerHTML = '';
    for (const meeting of meetings) {
      const row = document.createElement('div');
      row.className = 'meeting-item';
      const log = logs.get(meeting.id) || {};

      const check = document.createElement('input');
      check.type = 'checkbox';
      check.dataset.meetingId = meeting.id;
      check.checked = selected.has(meeting.id);
      check.disabled = state?.status === 'running';
      check.addEventListener('change', () => {
        startBatchBtn.disabled = !batchNativeAvailable || !selectedMeetingIds().length || batchState?.status === 'running';
        syncSelectAll();
      });

      const content = document.createElement('div');
      const title = document.createElement('div');
      title.className = 'meeting-title';
      title.textContent = meeting.title || meeting.label;
      const meta = document.createElement('div');
      meta.className = 'meeting-meta';
      meta.textContent = meeting.label;
      content.append(title, meta);

      const badge = document.createElement('span');
      badge.className = `meeting-status ${log.status || 'pending'}`;
      badge.textContent = statusLabel(log.status);
      badge.title = log.message || '';

      row.append(check, content, badge);
      meetingListEl.append(row);
    }
    syncSelectAll();
  }

  function renderBatchLog(state) {
    const lines = [];
    for (const item of state?.logs || []) {
      if (!item || item.status === 'pending') continue;
      lines.push(`${String(item.status || '').toUpperCase()}  ${item.title || item.meetingId}${item.message ? ` - ${item.message}` : ''}`);
      for (const file of item.files || []) lines.push(`      ${file.fileName || file.path}`);
    }
    const ops = (state?.operationLog || []).slice(-80);
    if (ops.length) {
      lines.push('', '--- Operations ---');
      for (const op of ops) {
        const time = op.ts ? op.ts.substring(11, 19) : '';
        const details = { ...op };
        delete details.ts;
        delete details.level;
        delete details.step;
        lines.push(`${time} ${op.level || 'INFO'} ${op.step || ''} ${JSON.stringify(details)}`);
      }
    }
    batchLogEl.textContent = lines.join('\n');
    batchLogEl.scrollTop = batchLogEl.scrollHeight;
  }

  function renderBatchState(state) {
    batchState = state || batchState;
    if (!batchState) return;

    const running = batchState.status === 'running';
    const meetings = batchState.meetings || [];
    const selected = batchState.selectedIds || [];
    const terminal = (batchState.logs || []).filter(x => ['done', 'skip', 'error'].includes(x.status)).length;
    const total = selected.length || meetings.length;
    const progress = total ? Math.round((terminal / total) * 100) : 0;

    batchStatusEl.textContent = batchState.message || 'Пакетный режим готов.';
    batchProgressBar.style.width = `${Math.max(0, Math.min(100, progress))}%`;
    batchProgressText.textContent = total ? `${terminal}/${total} встреч` : '';

    scanCalendarBtn.disabled = running;
    calendarDiagnosticBtn.disabled = running;
    stopBatchBtn.disabled = !running;
    openBatchFolderBtn.disabled = !batchState.folderPath;
    selectAllEl.disabled = running;

    renderMeetingList(batchState);
    renderBatchLog(batchState);

    startBatchBtn.disabled = running || !batchNativeAvailable || !meetings.length || !selectedMeetingIds().length;
  }

  async function refreshBatchState() {
    try {
      const response = await runtimeMessage({ type: 'BATCH_GET_STATE' });
      if (response?.ok) renderBatchState(response.state);
    } catch (e) {
      batchStatusEl.textContent = `Ошибка: ${e.message || e}`;
    }
  }

  async function detectNative() {
    try {
      const response = await nativeMessage({ action: 'ping' });
      batchNativeAvailable = !!response?.ok && response?.version === '2.0.0';
    } catch (_) {
      batchNativeAvailable = false;
    }
    batchSaveModeEl.textContent = batchNativeAvailable
      ? 'Batch: файлы будут сохранены в отдельную папку Windows Documents\\Teams Transcripts.'
      : 'Для пакетного режима нужен Windows helper v2.0. Повторно запусти Install-Windows-Integration.cmd после обновления.';
    renderBatchState(batchState);
  }

  scanCalendarBtn.addEventListener('click', async () => {
    try {
      const tab = await activeTab();
      batchStatusEl.textContent = 'Считываю встречи из текущего календаря Teams...';
      const response = await runtimeMessage({ type: 'BATCH_SCAN_CALENDAR', tabId: tab.id });
      if (!response?.ok) throw new Error(response?.error || 'Не удалось прочитать календарь.');
      renderBatchState(response.state);
    } catch (e) {
      batchStatusEl.textContent = `Ошибка: ${e.message || e}`;
    }
  });

  selectAllEl.addEventListener('change', () => {
    for (const el of meetingListEl.querySelectorAll('input[data-meeting-id]')) el.checked = selectAllEl.checked;
    startBatchBtn.disabled = !batchNativeAvailable || !selectedMeetingIds().length || batchState?.status === 'running';
  });

  startBatchBtn.addEventListener('click', async () => {
    try {
      const ids = selectedMeetingIds();
      const response = await runtimeMessage({ type: 'BATCH_START', selectedIds: ids });
      if (!response?.ok) throw new Error(response?.error || 'Не удалось запустить пакетный сбор.');
      batchStatusEl.textContent = 'Пакетный сбор запущен.';
      await refreshBatchState();
    } catch (e) {
      batchStatusEl.textContent = `Ошибка: ${e.message || e}`;
    }
  });

  stopBatchBtn.addEventListener('click', async () => {
    try {
      await runtimeMessage({ type: 'BATCH_STOP' });
      batchStatusEl.textContent = 'Останавливаю после текущей операции...';
    } catch (e) {
      batchStatusEl.textContent = `Ошибка: ${e.message || e}`;
    }
  });

  openBatchFolderBtn.addEventListener('click', async () => {
    try {
      const response = await runtimeMessage({ type: 'BATCH_OPEN_FOLDER' });
      if (!response?.ok) throw new Error(response?.error || 'Не удалось открыть папку.');
    } catch (e) {
      batchStatusEl.textContent = `Ошибка: ${e.message || e}`;
    }
  });

  calendarDiagnosticBtn.addEventListener('click', async () => {
    try {
      const tab = await activeTab();
      batchStatusEl.textContent = 'Собираю диагностику календаря...';
      const response = await runtimeMessage({ type: 'BATCH_CALENDAR_DIAGNOSTIC', tabId: tab.id });
      if (!response?.ok || !response.text) throw new Error(response?.error || 'Диагностика получилась пустой.');
      const path = await saveDiagnostic(response.text);
      batchStatusEl.textContent = `Диагностика календаря сохранена: ${path}`;
    } catch (e) {
      batchStatusEl.textContent = `Ошибка: ${e.message || e}`;
    }
  });

  (async () => {
    await refreshBatchState();
    await detectNative();
    if (batchPollTimer) clearInterval(batchPollTimer);
    batchPollTimer = setInterval(refreshBatchState, 900);
  })();
})();
