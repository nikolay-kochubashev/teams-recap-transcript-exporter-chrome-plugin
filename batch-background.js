const NATIVE_HOST = 'com.openai.teams_recap_transcript_exporter';
const BATCH_KEY = 'teamsTranscriptBatchStateV3';

let stopRequested = false;
let runPromise = null;
let batchState = {
  status: 'idle',
  calendarTabId: null,
  meetings: [],
  selectedIds: [],
  currentIndex: -1,
  folderPath: '',
  startedAt: 0,
  finishedAt: 0,
  message: 'Пакетный режим готов.',
  logs: [],
  operationLog: []
};

async function restoreBatchState() {
  try {
    const data = await chrome.storage.local.get(BATCH_KEY);
    if (data?.[BATCH_KEY]) {
      batchState = {
        ...batchState,
        ...data[BATCH_KEY],
        status: data[BATCH_KEY].status === 'running' ? 'interrupted' : data[BATCH_KEY].status
      };
    }
  } catch (_) {}
}

async function persistBatchState() {
  try { await chrome.storage.local.set({ [BATCH_KEY]: batchState }); } catch (_) {}
}

async function patchBatch(patch) {
  batchState = { ...batchState, ...patch };
  await persistBatchState();
}

async function appendOperation(step, details = {}, level = 'INFO') {
  const operationLog = Array.isArray(batchState.operationLog) ? [...batchState.operationLog] : [];
  operationLog.push({
    ts: new Date().toISOString(),
    level,
    step,
    ...details
  });
  if (operationLog.length > 500) operationLog.splice(0, operationLog.length - 500);
  await patchBatch({ operationLog });
}

async function logMeeting(meetingId, patch) {
  const logs = Array.isArray(batchState.logs) ? [...batchState.logs] : [];
  const index = logs.findIndex(x => x.meetingId === meetingId);
  const previous = index >= 0
    ? logs[index]
    : { meetingId, title: '', status: 'pending', message: '', files: [] };
  const next = { ...previous, ...patch };
  if (index >= 0) logs[index] = next;
  else logs.push(next);
  await patchBatch({ logs });
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function nativeMessage(message) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendNativeMessage(NATIVE_HOST, message, response => {
      const err = chrome.runtime.lastError;
      if (err) return reject(new Error(err.message));
      if (!response) return reject(new Error('Windows helper не вернул ответ.'));
      if (response.ok === false) return reject(new Error(response.error || 'Windows helper error.'));
      resolve(response);
    });
  });
}

async function sendTab(tabId, message, timeoutMs = 12000) {
  const started = Date.now();
  let lastError = null;
  while (Date.now() - started < timeoutMs) {
    try {
      return await chrome.tabs.sendMessage(tabId, message);
    } catch (e) {
      lastError = e;
      await delay(300);
    }
  }
  throw new Error(lastError?.message || 'Страница не ответила расширению.');
}

async function waitTabComplete(tabId, timeoutMs = 30000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const tab = await chrome.tabs.get(tabId);
    if (tab.status === 'complete') return tab;
    await delay(250);
  }
  return await chrome.tabs.get(tabId);
}

function isRecordingUrl(url) {
  const v = String(url || '').toLowerCase();
  return /sharepoint\.com/.test(v) &&
    (/stream\.aspx/.test(v) || /recording/.test(v) || /\.mp4(?:\?|$)/.test(v));
}

function sanitizeFileName(value) {
  return (value || 'teams-transcript')
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/[. ]+$/g, '')
    .trim()
    .slice(0, 150) || 'teams-transcript';
}

function dateStampFromUrl(url) {
  let decoded = String(url || '');
  try { decoded = decodeURIComponent(decoded); } catch (_) {}
  const m = decoded.match(/\b(20\d{6})[_-]/);
  return m ? m[1] : '';
}

function todayStamp() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

async function scanCalendar(tabId) {
  await appendOperation('SCAN_CALENDAR_START', { tabId });
  const response = await sendTab(tabId, { type: 'CALENDAR_SCAN' }, 10000);
  if (!response?.ok) throw new Error(response?.error || 'Не удалось прочитать календарь Teams.');
  const meetings = (response.meetings || []).map(m => ({ ...m, status: 'pending' }));
  await appendOperation('SCAN_CALENDAR_RESULT', {
    tabId,
    count: meetings.length,
    meetings: meetings.slice(0, 30).map(m => ({
      id: m.id,
      title: m.title,
      label: m.label,
      score: m.score,
      dom: m.dom || {},
      rect: m.rect || {}
    }))
  });
  await patchBatch({
    calendarTabId: tabId,
    meetings,
    selectedIds: [],
    currentIndex: -1,
    status: 'ready',
    message: `Найдено встреч: ${meetings.length}.`,
    logs: meetings.map(m => ({
      meetingId: m.id,
      title: m.title || m.label,
      status: 'pending',
      message: '',
      files: []
    }))
  });
  return meetings;
}

async function findActionsWithWait(tabId, timeoutMs = 8000) {
  const started = Date.now();
  let last = null;
  while (Date.now() - started < timeoutMs) {
    try {
      last = await sendTab(tabId, { type: 'PAGE_FIND_RECORDINGS' }, 2500);
      const has = (last?.recordingLinks?.length || 0) + (last?.actions?.length || 0);
      if (has > 0) return last;
    } catch (_) {}
    await delay(450);
  }
  return last || { ok: true, recordingLinks: [], actions: [] };
}

async function triggerPreferredAction(tabId, actions, kinds) {
  const action = (actions || []).find(a => kinds.includes(a.kind));
  if (!action) return null;

  if (action.href && /^https?:/i.test(action.href)) {
    await chrome.tabs.update(tabId, { url: action.href, active: false });
    await waitTabComplete(tabId, 30000);
    await delay(700);
    return action;
  }

  await sendTab(tabId, { type: 'PAGE_TRIGGER_ACTION', id: action.id }, 5000);
  await delay(1200);
  return action;
}

async function extractRecordingFromTab(tabId, meeting, recordingUrl, recordingIndex) {
  await patchBatch({
    message: `Собираю транскрипцию: ${meeting.title || meeting.label}`
  });

  await sendTab(tabId, { type: 'START_EXTRACT' }, 20000);

  const started = Date.now();
  let state = null;

  while (Date.now() - started < 12 * 60 * 1000) {
    if (stopRequested) throw new Error('Остановлено пользователем.');

    const response = await sendTab(tabId, { type: 'GET_STATE' }, 5000);
    state = response?.state;

    if (state?.status === 'done') break;
    if (state?.status === 'error') {
      throw new Error(state.error || state.message || 'Ошибка сборщика транскрипции.');
    }
    if (state?.status === 'cancelled') {
      throw new Error('Сбор транскрипции отменен.');
    }

    await patchBatch({
      message: `${meeting.title || meeting.label}: ${state?.progress || 0}%`
    });
    await delay(700);
  }

  if (!state || state.status !== 'done' || !state.text) {
    throw new Error('Тайм-аут ожидания транскрипции.');
  }

  const stamp = dateStampFromUrl(recordingUrl) || meeting.dateStamp || todayStamp();
  const suffix = recordingIndex > 0
    ? `_${String(recordingIndex + 1).padStart(2, '0')}`
    : '';
  const fileName = `${sanitizeFileName(state.title || meeting.title || meeting.label)} - ${stamp}${suffix}.txt`;

  const saved = await nativeMessage({
    action: 'saveTextInFolder',
    folderPath: batchState.folderPath,
    fileName,
    text: state.text
  });

  await appendOperation('TRANSCRIPT_SAVED', {
    meetingId: meeting.id,
    recordingUrl,
    recordingIndex,
    fileName,
    path: saved.path,
    chars: state.text.length,
    items: state.items || 0
  });

  return {
    path: saved.path,
    fileName,
    chars: state.text.length,
    items: state.items || 0
  };
}

async function processRecordingUrl(url, meeting, index, reuseTabId = null) {
  let tabId = reuseTabId;
  let ownTab = false;

  try {
    if (!tabId) {
      const tab = await chrome.tabs.create({ url, active: false });
      tabId = tab.id;
      ownTab = true;
      await waitTabComplete(tabId, 30000);
      await delay(800);
    } else {
      const tab = await chrome.tabs.get(tabId);
      if (tab.url !== url) {
        await chrome.tabs.update(tabId, { url, active: false });
        await waitTabComplete(tabId, 30000);
        await delay(800);
      }
    }

    return await extractRecordingFromTab(tabId, meeting, url, index);
  } finally {
    if (ownTab && tabId) {
      try { await chrome.tabs.remove(tabId); } catch (_) {}
    }
  }
}

async function discoverRecordingUrls(tempTabId) {
  let tab = await chrome.tabs.get(tempTabId);
  await appendOperation('DISCOVER_RECORDING_START', { tabId: tempTabId, url: tab.url });
  if (isRecordingUrl(tab.url)) return [tab.url];

  let page = await findActionsWithWait(tempTabId, 6000);
  await appendOperation('DISCOVER_ACTIONS', {
    tabId: tempTabId,
    url: tab.url,
    actions: (page.actions || []).map(a => ({ kind: a.kind, label: a.label, href: a.href || '' })).slice(0, 40),
    recordingLinks: (page.recordingLinks || []).slice(0, 20)
  });
  let urls = (page.recordingLinks || []).map(x => x.href).filter(Boolean);
  if (urls.length) return [...new Set(urls)];

  const direct = (page.actions || [])
    .find(a => ['recap', 'recording'].includes(a.kind) && a.href);

  if (direct?.href) {
    if (isRecordingUrl(direct.href)) return [direct.href];

    await chrome.tabs.update(tempTabId, { url: direct.href, active: false });
    await waitTabComplete(tempTabId, 30000);
    await delay(900);

    tab = await chrome.tabs.get(tempTabId);
    if (isRecordingUrl(tab.url)) return [tab.url];

    page = await findActionsWithWait(tempTabId, 7000);
    urls = (page.recordingLinks || []).map(x => x.href).filter(Boolean);
    if (urls.length) return [...new Set(urls)];
  }

  const clickableDirect = (page.actions || [])
    .find(a => ['recap', 'recording'].includes(a.kind) && !a.href);

  if (clickableDirect) {
    await triggerPreferredAction(tempTabId, page.actions, ['recap', 'recording']);
    await delay(1200);

    tab = await chrome.tabs.get(tempTabId);
    if (isRecordingUrl(tab.url)) return [tab.url];

    page = await findActionsWithWait(tempTabId, 7000);
    urls = (page.recordingLinks || []).map(x => x.href).filter(Boolean);
    if (urls.length) return [...new Set(urls)];
  }

  page = page || await findActionsWithWait(tempTabId, 4000);
  const chatAction = (page.actions || []).find(a => a.kind === 'chat');
  if (!chatAction) return [];

  await triggerPreferredAction(tempTabId, page.actions, ['chat']);
  await delay(1200);

  tab = await chrome.tabs.get(tempTabId);
  if (isRecordingUrl(tab.url)) return [tab.url];

  page = await findActionsWithWait(tempTabId, 10000);
  urls = (page.recordingLinks || []).map(x => x.href).filter(Boolean);
  if (urls.length) return [...new Set(urls)];

  const recapInChat = (page.actions || [])
    .find(a => ['recap', 'recording'].includes(a.kind));

  if (recapInChat) {
    await triggerPreferredAction(tempTabId, page.actions, ['recap', 'recording']);
    await delay(1500);

    tab = await chrome.tabs.get(tempTabId);
    if (isRecordingUrl(tab.url)) return [tab.url];

    page = await findActionsWithWait(tempTabId, 8000);
    urls = (page.recordingLinks || []).map(x => x.href).filter(Boolean);
  }

  return [...new Set(urls)];
}

async function restoreCalendarForMeeting(tabId, meeting, timeoutMs = 30000) {
  const started = Date.now();
  let backAttempts = 0;

  while (Date.now() - started < timeoutMs) {
    try {
      const response = await sendTab(tabId, { type: 'CALENDAR_SCAN' }, 4000);
      const found = (response?.meetings || []).some(m =>
        m.id === meeting.id ||
        (m.label && meeting.label && m.label === meeting.label) ||
        (m.title === meeting.title && m.dateStamp === meeting.dateStamp && m.startTime === meeting.startTime)
      );
      if (response?.ok && found) {
        return { ok: true, url: response.url || '' };
      }
    } catch (_) {}

    if (backAttempts >= 6) break;

    try {
      await appendOperation('CALENDAR_RESTORE_BACK', {
        tabId,
        meetingId: meeting.id,
        attempt: backAttempts + 1
      });
      await chrome.tabs.goBack(tabId);
      backAttempts++;
      await waitTabComplete(tabId, 20000);
      await delay(1200);
    } catch (_) {
      break;
    }
  }

  return { ok: false };
}

async function processMeeting(meeting, index) {
  await appendOperation('MEETING_START', {
    meetingId: meeting.id,
    index,
    title: meeting.title || meeting.label,
    label: meeting.label,
    dateStamp: meeting.dateStamp || '',
    startTime: meeting.startTime || '',
    dom: meeting.dom || {},
    rect: meeting.rect || {}
  });

  await logMeeting(meeting.id, {
    status: 'running',
    message: 'Открываю встречу...',
    files: []
  });

  await patchBatch({
    currentIndex: index,
    message: `Обрабатываю ${index + 1}/${batchState.selectedIds.length}: ${meeting.title || meeting.label}`
  });

  const calendarTabId = batchState.calendarTabId;

  try {
    const restoredBefore = await restoreCalendarForMeeting(calendarTabId, meeting, 20000);
    if (!restoredBefore.ok) {
      await appendOperation('MEETING_NOT_VISIBLE_IN_CALENDAR', {
        meetingId: meeting.id,
        calendarTabId
      }, 'ERROR');
      throw new Error('Не удалось восстановить календарь с выбранной встречей. Обнови Calendar, заново нажми "Считать календарь" и повтори запуск.');
    }

    const opened = await sendTab(
      calendarTabId,
      { type: 'CALENDAR_OPEN_MEETING', meeting },
      12000
    );

    if (!opened?.ok) {
      await appendOperation('OPEN_MEETING_FAILED', {
        meetingId: meeting.id,
        calendarTabId,
        error: opened?.error || ''
      }, 'ERROR');
      throw new Error(opened?.error || 'Не удалось открыть карточку встречи в календаре.');
    }

    await appendOperation('OPEN_MEETING_OK', {
      meetingId: meeting.id,
      calendarTabId,
      url: opened.url || '',
      target: opened.target || {}
    });

    await delay(1600);

    const urls = await discoverRecordingUrls(calendarTabId);

    await appendOperation('RECORDING_URLS_RESULT', {
      meetingId: meeting.id,
      calendarTabId,
      count: urls.length,
      urls
    });

    if (!urls.length) {
      await appendOperation('MEETING_SKIP', {
        meetingId: meeting.id,
        reason: 'Запись/Recap не найдены.'
      }, 'WARN');

      await logMeeting(meeting.id, {
        status: 'skip',
        message: 'Запись/Recap не найдены.',
        files: []
      });
      return;
    }

    const files = [];

    for (let i = 0; i < urls.length; i++) {
      if (stopRequested) throw new Error('Остановлено пользователем.');

      const file = await processRecordingUrl(urls[i], meeting, i, null);
      files.push(file);
    }

    await appendOperation('MEETING_DONE', {
      meetingId: meeting.id,
      files: files.map(f => ({
        fileName: f.fileName,
        path: f.path,
        chars: f.chars,
        items: f.items
      }))
    });

    await logMeeting(meeting.id, {
      status: 'done',
      message: `Сохранено файлов: ${files.length}.`,
      files
    });
  } catch (e) {
    if (String(e?.message || e).includes('Остановлено пользователем')) throw e;

    await appendOperation('MEETING_ERROR', {
      meetingId: meeting.id,
      error: e?.message || String(e)
    }, 'ERROR');

    await logMeeting(meeting.id, {
      status: 'error',
      message: e?.message || String(e),
      files: []
    });
  } finally {
    const restored = await restoreCalendarForMeeting(calendarTabId, meeting, 30000);
    await appendOperation('CALENDAR_RESTORE_RESULT', {
      meetingId: meeting.id,
      calendarTabId,
      restored: restored.ok,
      url: restored.url || ''
    }, restored.ok ? 'INFO' : 'WARN');
  }
}

function reportText() {
  const lines = [
    'Teams Recap Transcript Exporter - batch report',
    `Started: ${batchState.startedAt ? new Date(batchState.startedAt).toISOString() : '-'}`,
    `Finished: ${batchState.finishedAt ? new Date(batchState.finishedAt).toISOString() : '-'}`,
    `Folder: ${batchState.folderPath || '-'}`,
    ''
  ];

  for (const item of batchState.logs || []) {
    const status = String(item.status || 'pending').toUpperCase().padEnd(7, ' ');
    lines.push(
      `${status} ${item.title || item.meetingId}${item.message ? ` - ${item.message}` : ''}`
    );

    for (const file of item.files || []) {
      lines.push(`        ${file.fileName || file.path}`);
    }
  }

  return lines.join('\n');
}

async function runBatch() {
  stopRequested = false;

  try {
    const ping = await nativeMessage({ action: 'ping' });
    if (!ping?.ok || ping.version !== '2.0.0') {
      throw new Error('Нужен Windows helper v2.0. Переустанови его через Install-Windows-Integration.cmd.');
    }

    const folder = await nativeMessage({ action: 'createBatchFolder' });

    await patchBatch({
      folderPath: folder.path,
      status: 'running',
      startedAt: Date.now(),
      finishedAt: 0,
      message: 'Пакетный сбор запущен.'
    });

    const selected = batchState.selectedIds || [];

    for (let i = 0; i < selected.length; i++) {
      if (stopRequested) break;

      const meeting = (batchState.meetings || [])
        .find(m => m.id === selected[i]);

      if (!meeting) continue;

      const existing = (batchState.logs || [])
        .find(x => x.meetingId === meeting.id);

      if (existing?.status === 'done') continue;

      await processMeeting(meeting, i);
    }

    const finalStatus = stopRequested ? 'stopped' : 'done';

    await patchBatch({
      status: finalStatus,
      finishedAt: Date.now(),
      message: stopRequested
        ? 'Пакетный сбор остановлен.'
        : 'Пакетный сбор завершен.'
    });

    await nativeMessage({
      action: 'saveTextInFolder',
      folderPath: batchState.folderPath,
      fileName: 'batch-report.txt',
      text: reportText()
    });
  } catch (e) {
    await patchBatch({
      status: 'error',
      finishedAt: Date.now(),
      message: e?.message || String(e)
    });
  } finally {
    runPromise = null;
  }
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  const type = message?.type;

  if (type === 'BATCH_GET_STATE') {
    sendResponse({ ok: true, state: batchState });
    return;
  }

  if (type === 'BATCH_SCAN_CALENDAR') {
    (async () => {
      try {
        const meetings = await scanCalendar(message.tabId);
        sendResponse({ ok: true, meetings, state: batchState });
      } catch (e) {
        sendResponse({ ok: false, error: e?.message || String(e) });
      }
    })();
    return true;
  }

  if (type === 'BATCH_SET_SELECTION') {
    (async () => {
      try {
        const validIds = new Set((batchState.meetings || []).map(m => m.id));
        const selectedIds = (Array.isArray(message.selectedIds) ? message.selectedIds : [])
          .filter(id => validIds.has(id));
        await patchBatch({ selectedIds });
        sendResponse({ ok: true, selectedIds, state: batchState });
      } catch (e) {
        sendResponse({ ok: false, error: e?.message || String(e) });
      }
    })();
    return true;
  }

  if (type === 'BATCH_START') {
    (async () => {
      try {
        if (runPromise) {
          sendResponse({ ok: false, error: 'Пакетный сбор уже выполняется.' });
          return;
        }

        const selectedIds = Array.isArray(message.selectedIds)
          ? message.selectedIds
          : [];

        if (!selectedIds.length) {
          sendResponse({ ok: false, error: 'Не выбрана ни одна встреча.' });
          return;
        }

        await patchBatch({
          selectedIds,
          logs: (batchState.logs || []).map(x =>
            selectedIds.includes(x.meetingId)
              ? { ...x, status: 'pending', message: '', files: [] }
              : x
          )
        });

        runPromise = runBatch();
        sendResponse({ ok: true, state: batchState });
      } catch (e) {
        sendResponse({ ok: false, error: e?.message || String(e) });
      }
    })();
    return true;
  }

  if (type === 'BATCH_STOP') {
    stopRequested = true;
    sendResponse({ ok: true });
    return;
  }

  if (type === 'BATCH_OPEN_FOLDER') {
    (async () => {
      try {
        if (!batchState.folderPath) {
          throw new Error('Папка пакетного запуска еще не создана.');
        }

        const response = await nativeMessage({
          action: 'openDirectory',
          path: batchState.folderPath
        });

        sendResponse(response);
      } catch (e) {
        sendResponse({ ok: false, error: e?.message || String(e) });
      }
    })();
    return true;
  }

  if (type === 'BATCH_CALENDAR_DIAGNOSTIC') {
    (async () => {
      try {
        const response = await sendTab(
          message.tabId || batchState.calendarTabId,
          { type: 'CALENDAR_DEBUG' },
          10000
        );

        if (response?.ok && response.text) {
          const ops = (batchState.operationLog || []).map(x =>
            `[${x.ts}] ${x.level} ${x.step} ${JSON.stringify(x)}`
          ).join('\n');
          response.text += '\n\n=== BATCH OPERATION LOG ===\n' + ops;
        }

        sendResponse(response);
      } catch (e) {
        sendResponse({ ok: false, error: e?.message || String(e) });
      }
    })();
    return true;
  }
});

restoreBatchState();