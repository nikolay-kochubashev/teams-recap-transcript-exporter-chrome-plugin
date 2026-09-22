const NATIVE_HOST = 'com.openai.teams_recap_transcript_exporter';
const BATCH_KEY = 'teamsTranscriptBatchStateV5';

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

async function sendFrame(tabId, frameId, message, timeoutMs = 12000) {
  const started = Date.now();
  let lastError = null;
  while (Date.now() - started < timeoutMs) {
    try {
      return await chrome.tabs.sendMessage(tabId, message, { frameId });
    } catch (e) {
      lastError = e;
      await delay(250);
    }
  }
  throw new Error(lastError?.message || `Frame ${frameId} did not answer extension.`);
}

async function probeTranscriptFrames(tabId, timeoutMs = 9000) {
  const started = Date.now();
  let last = { selected: null, frames: [] };

  while (Date.now() - started < timeoutMs) {
    let frames = [];
    try {
      frames = await chrome.webNavigation.getAllFrames({ tabId }) || [];
    } catch (_) {
      frames = [{ frameId: 0, url: (await chrome.tabs.get(tabId)).url || '' }];
    }

    const results = [];
    for (const frame of frames) {
      try {
        const probe = await sendFrame(tabId, frame.frameId, { type: 'PROBE_TRANSCRIPT' }, 1800);
        results.push({
          frameId: frame.frameId,
          parentFrameId: frame.parentFrameId ?? -1,
          frameUrl: frame.url || '',
          ...probe
        });
      } catch (_) {}
    }

    results.sort((a, b) => {
      const as = a.best ? (a.best.strong ? 10000 : 0) + (a.best.score || 0) : (a.transcriptVisible ? 100 : 0);
      const bs = b.best ? (b.best.strong ? 10000 : 0) + (b.best.score || 0) : (b.transcriptVisible ? 100 : 0);
      return bs - as;
    });

    // Do not accept the Teams page shell merely because it contains the
    // word "Transcript". Right after opening Recap, the top frame exposes the
    // Transcript tab before the lazy-loaded SharePoint transcript frame is ready.
    // A real transcript scroller has semantic evidence such as transcript entry
    // votes, timestamp rows or the Teams transcript warning/accessibility text.
    const selected = results.find(x =>
      x.best && x.best.strong && (
        (x.best.votes || 0) > 0 ||
        (x.best.times || 0) > 0 ||
        (x.best.reasons || []).includes('ai-warning') ||
        (x.best.reasons || []).includes('transcript-a11y')
      )
    ) || null;

    last = { selected, frames: results };
    if (selected) return last;
    await delay(350);
  }

  return last;
}

async function collectFrameDiagnostics(tabId) {
  let frames = [];
  try {
    frames = await chrome.webNavigation.getAllFrames({ tabId }) || [];
  } catch (_) {
    frames = [{ frameId: 0, url: (await chrome.tabs.get(tabId)).url || '' }];
  }

  const out = [];
  for (const frame of frames) {
    try {
      const debug = await sendFrame(tabId, frame.frameId, { type: 'GET_DEBUG' }, 2200);
      out.push({
        frameId: frame.frameId,
        parentFrameId: frame.parentFrameId ?? -1,
        frameUrl: frame.url || '',
        debug: String(debug?.text || '').slice(0, 12000)
      });
    } catch (_) {}
  }
  return out;
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

  // Never navigate the Calendar anchor tab through an href.
  // Href targets are resolved in their own background tab.
  if (action.href && /^https?:/i.test(action.href)) {
    return action;
  }

  await sendTab(tabId, { type: 'PAGE_TRIGGER_ACTION', id: action.id }, 5000);
  await delay(1200);
  return action;
}

async function triggerActionAndCaptureNewTabs(tabId, actionId) {
  const source = await chrome.tabs.get(tabId);
  const before = await chrome.tabs.query({ windowId: source.windowId });
  const beforeIds = new Set(before.map(t => t.id));

  await sendTab(tabId, { type: 'PAGE_TRIGGER_ACTION', id: actionId }, 5000);
  await delay(1600);

  const after = await chrome.tabs.query({ windowId: source.windowId });
  const created = after.filter(t => !beforeIds.has(t.id));

  return {
    current: await chrome.tabs.get(tabId),
    created
  };
}

async function resolveHrefForRecordings(url, meeting) {
  if (!url || !/^https?:/i.test(url)) return [];
  if (isRecordingUrl(url)) return [url];

  let tabId = null;
  try {
    const tab = await chrome.tabs.create({ url, active: false });
    tabId = tab.id;
    await waitTabComplete(tabId, 30000);
    await delay(1000);

    const loaded = await chrome.tabs.get(tabId);
    if (isRecordingUrl(loaded.url)) return [loaded.url];

    let page = await findActionsWithWait(tabId, 8000);
    let urls = (page?.recordingLinks || []).map(x => x.href).filter(Boolean);
    if (urls.length) return [...new Set(urls)];

    const cards = await waitForRecapCards(tabId, meeting?.dateStamp, 5000);
    if (cards.length) {
      const target = cards[0];
      if (target.href && isRecordingUrl(target.href)) return [target.href];
      if (!target.href) {
        await sendTab(tabId, { type: 'PAGE_TRIGGER_ACTION', id: target.id }, 5000);
        await delay(1500);
        const after = await chrome.tabs.get(tabId);
        if (isRecordingUrl(after.url)) return [after.url];
        page = await findActionsWithWait(tabId, 8000);
        urls = (page?.recordingLinks || []).map(x => x.href).filter(Boolean);
        if (urls.length) return [...new Set(urls)];
      }
    }
  } catch (_) {
  } finally {
    if (tabId) {
      try { await chrome.tabs.remove(tabId); } catch (_) {}
    }
  }
  return [];
}

async function waitForRecapCards(tabId, dateStamp, timeoutMs = 12000) {
  const started = Date.now();
  let last = [];
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await sendTab(tabId, { type: 'PAGE_FIND_RECAP_CARDS' }, 3500);
      last = response?.cards || [];
      const matching = dateStamp ? last.filter(x => x.dateStamp === dateStamp) : last;
      if (matching.length) return matching;
    } catch (_) {}
    await delay(500);
  }
  return dateStamp ? last.filter(x => x.dateStamp === dateStamp) : last;
}

async function extractRecordingFromTab(tabId, meeting, recordingUrl, recordingIndex, preferredFrameId = null, fileMeta = null) {
  await patchBatch({
    message: `Собираю транскрипцию: ${meeting.title || meeting.label}`
  });

  let frameId = preferredFrameId;
  if (frameId === null || frameId === undefined) {
    const probe = await probeTranscriptFrames(tabId, 6000);
    frameId = probe.selected?.frameId ?? 0;

    await appendOperation('TRANSCRIPT_FRAME_LOOKUP', {
      meetingId: meeting.id,
      selectedFrameId: frameId,
      frames: probe.frames.slice(0, 12).map(x => ({
        frameId: x.frameId,
        parentFrameId: x.parentFrameId,
        frameUrl: x.frameUrl,
        url: x.url || '',
        title: x.title || '',
        isTop: !!x.isTop,
        transcriptVisible: !!x.transcriptVisible,
        best: x.best || null
      }))
    }, probe.selected ? 'INFO' : 'WARN');
  }

  await sendFrame(tabId, frameId, { type: 'START_EXTRACT' }, 20000);

  const started = Date.now();
  let state = null;

  while (Date.now() - started < 12 * 60 * 1000) {
    if (stopRequested) throw new Error('Остановлено пользователем.');

    const response = await sendFrame(tabId, frameId, { type: 'GET_STATE' }, 5000);
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

  // Guard against a false-positive extraction from the Teams Recap page shell.
  // Real transcript output always contains at least one standalone timestamp;
  // the shell snapshot contains only UI text such as Speakers/Shared files/Transcript.
  const timestampLines = String(state.text || '').match(/^(?:\d{1,2}:)?\d{1,2}:\d{2}\s*$/gm) || [];
  if (!timestampLines.length) {
    throw new Error('Собран не текст транскрипции, а оболочка Recap. Ожидаю загрузку реальной области Transcript.');
  }

  const stamp = dateStampFromUrl(recordingUrl) || meeting.dateStamp || todayStamp();
  const sessionStart = String(fileMeta?.sessionStartTime || '').replace(':', '');
  const sessionEnd = String(fileMeta?.sessionEndTime || '').replace(':', '');
  const transcriptIndex = Number(fileMeta?.transcriptIndex || 0);

  let suffix = '';
  if (sessionStart) {
    suffix = ` - ${sessionStart}${sessionEnd ? `-${sessionEnd}` : ''}`;
    if (transcriptIndex > 0) {
      suffix += `_${String(transcriptIndex + 1).padStart(2, '0')}`;
    }
  } else if (recordingIndex > 0) {
    suffix = `_${String(recordingIndex + 1).padStart(2, '0')}`;
  }

  const fileName = `${sanitizeFileName(meeting.title || meeting.label || state.title)} - ${stamp}${suffix}.txt`;

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
    frameId,
    fileMeta: fileMeta || null,
    fileName,
    path: saved.path,
    chars: state.text.length,
    items: state.items || 0
  });

  return {
    path: saved.path,
    fileName,
    chars: state.text.length,
    items: state.items || 0,
    frameId
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

async function findMeetingScopedActionsWithWait(tabId, meeting, timeoutMs = 8000) {
  const started = Date.now();
  let last = { ok: true, actions: [], context: null };

  while (Date.now() - started < timeoutMs) {
    try {
      last = await sendTab(tabId, {
        type: 'PAGE_FIND_MEETING_ACTIONS',
        meeting
      }, 3500);

      if ((last?.actions || []).length) return last;
    } catch (_) {}

    await delay(450);
  }

  return last;
}

function conversationMatchesMeeting(context, meeting) {
  const actual = normalizeMeetingKey(context?.conversationTitle || '');
  const expected = normalizeMeetingKey(meeting?.title || '');

  if (!actual) return { ok: true, uncertain: true };
  if (!expected) return { ok: false, actual, expected };

  const ok = actual === expected || actual.includes(expected) || expected.includes(actual);
  return { ok, actual, expected, uncertain: false };
}

async function validateMeetingConversation(tabId, meeting) {
  const response = await sendTab(tabId, { type: 'PAGE_GET_CONTEXT' }, 5000);
  const verdict = conversationMatchesMeeting(response?.context, meeting);

  await appendOperation('MEETING_CONTEXT_CHECK', {
    meetingId: meeting?.id,
    expectedTitle: meeting?.title || '',
    actualTitle: response?.context?.conversationTitle || '',
    pageKind: response?.context?.pageKind || '',
    url: response?.context?.url || '',
    ok: verdict.ok,
    uncertain: !!verdict.uncertain
  }, verdict.ok ? 'INFO' : 'ERROR');

  return { ...verdict, context: response?.context || null };
}

async function waitForExactMeetingRecap(tabId, meeting, timeoutMs = 20000) {
  const started = Date.now();
  let last = null;

  while (Date.now() - started < timeoutMs) {
    try {
      last = await sendTab(tabId, {
        type: 'PAGE_FIND_MEETING_RECAP_EXACT',
        meeting
      }, 4000);

      if (last?.match) return last;
    } catch (_) {}

    await delay(450);
  }

  return last || { ok: true, match: null, matches: [] };
}

async function waitForMeetingDetailsAssets(tabId, meeting, timeoutMs = 7000) {
  const started = Date.now();
  let last = null;

  while (Date.now() - started < timeoutMs) {
    try {
      last = await sendTab(tabId, {
        type: 'PAGE_FIND_MEETING_DETAILS_ASSETS',
        meeting
      }, 4000);

      if (last?.match && (last?.found?.transcript || last?.found?.recording)) {
        if (last?.found?.transcript) return last;
      }
    } catch (_) {}

    await delay(350);
  }

  return last;
}

function isRecurringMeeting(meeting) {
  return /\brecurring meeting\b/i.test(String(meeting?.label || ''));
}

async function ensureExactRecapCard(tabId, meeting, timeoutMs = 12000) {
  let recap = await waitForExactMeetingRecap(tabId, meeting, Math.min(3000, timeoutMs));
  if (recap?.match) return recap;

  // Calendar may open the Recap/Details surface directly. For recurring
  // meetings we must return to Chat, where every occurrence has its own
  // recap card containing the real occurrence date.
  try {
    const details = await sendTab(tabId, {
      type: 'PAGE_FIND_MEETING_DETAILS_ASSETS',
      meeting
    }, 5000);

    const chatActionId = details?.actions?.chatTabActionId;
    if (chatActionId) {
      const trigger = await sendTab(tabId, {
        type: 'PAGE_TRIGGER_ACTION',
        id: chatActionId
      }, 5000);

      await appendOperation('EXACT_RECAP_CHAT_TRIGGER', {
        meetingId: meeting.id,
        actionId: chatActionId,
        trigger
      }, trigger?.ok ? 'INFO' : 'WARN');

      if (trigger?.ok) await delay(900);
    }
  } catch (e) {
    await appendOperation('EXACT_RECAP_CHAT_TRIGGER_ERROR', {
      meetingId: meeting.id,
      error: e?.message || String(e)
    }, 'WARN');
  }

  recap = await waitForExactMeetingRecap(tabId, meeting, timeoutMs);
  return recap;
}

async function extractTranscriptAfterAction(tabId, meeting, actionId, recordingIndex, sourcePrefix, fileMeta = null) {
  let trigger = null;
  try {
    trigger = await sendTab(tabId, {
      type: 'PAGE_TRIGGER_ACTION',
      id: actionId
    }, 5000);
  } catch (e) {
    await appendOperation(`${sourcePrefix}_TRANSCRIPT_TRIGGER_ERROR`, {
      meetingId: meeting.id,
      actionId,
      error: e?.message || String(e)
    }, 'WARN');
    return null;
  }

  await appendOperation(`${sourcePrefix}_TRANSCRIPT_TRIGGER`, {
    meetingId: meeting.id,
    actionId,
    recordingIndex,
    trigger
  }, trigger?.ok ? 'INFO' : 'WARN');

  if (!trigger?.ok) return null;

  await delay(1400);

  try {
    try {
      await sendTab(tabId, { type: 'CLEAR_RESULT' }, 3000);
    } catch (_) {}

    const probe = await probeTranscriptFrames(tabId, 10000);

    await appendOperation(`${sourcePrefix}_TRANSCRIPT_FRAME_LOOKUP`, {
      meetingId: meeting.id,
      recordingIndex,
      selectedFrameId: probe.selected?.frameId ?? null,
      frames: probe.frames.slice(0, 12).map(x => ({
        frameId: x.frameId,
        parentFrameId: x.parentFrameId,
        frameUrl: x.frameUrl,
        url: x.url || '',
        title: x.title || '',
        isTop: !!x.isTop,
        transcriptVisible: !!x.transcriptVisible,
        best: x.best || null
      }))
    }, probe.selected ? 'INFO' : 'WARN');

    if (!probe.selected) {
      const diagnostics = await collectFrameDiagnostics(tabId);
      await appendOperation(`${sourcePrefix}_TRANSCRIPT_FRAME_DIAGNOSTIC`, {
        meetingId: meeting.id,
        recordingIndex,
        diagnostics
      }, 'WARN');
      throw new Error('Transcript открыт, но область транскрипции не найдена ни в одном frame.');
    }

    try {
      await sendFrame(tabId, probe.selected.frameId, { type: 'CLEAR_RESULT' }, 3000);
    } catch (_) {}

    const tab = await chrome.tabs.get(tabId);
    const file = await extractRecordingFromTab(
      tabId,
      meeting,
      tab.url || 'https://teams.microsoft.com/v2/',
      recordingIndex,
      probe.selected.frameId,
      fileMeta
    );

    await appendOperation(`${sourcePrefix}_TRANSCRIPT_EXTRACT_OK`, {
      meetingId: meeting.id,
      recordingIndex,
      frameId: probe.selected.frameId,
      fileName: file?.fileName || '',
      path: file?.path || '',
      chars: file?.chars || 0,
      items: file?.items || 0
    });

    return file;
  } catch (e) {
    let diagnostics = [];
    try { diagnostics = await collectFrameDiagnostics(tabId); } catch (_) {}

    await appendOperation(`${sourcePrefix}_TRANSCRIPT_EXTRACT_FAILED`, {
      meetingId: meeting.id,
      recordingIndex,
      error: e?.message || String(e),
      diagnostics
    }, 'WARN');
    return null;
  }
}

async function tryExtractTranscriptsFromExactRecap(tabId, meeting) {
  let recap = await ensureExactRecapCard(tabId, meeting, 12000);

  await appendOperation('EXACT_TRANSCRIPT_RECAP_LOOKUP', {
    meetingId: meeting.id,
    expectedDate: meeting.dateStamp || '',
    expectedStartTime: meeting.startTime || '',
    expectedEndTime: meeting.endTime || '',
    pageTitle: recap?.pageTitle || '',
    url: recap?.url || '',
    match: recap?.match || null,
    matches: (recap?.matches || []).slice(0, 20)
  }, recap?.match ? 'INFO' : 'WARN');

  const initialMatches = (recap?.matches || []).filter(x => x?.hasTranscript);
  if (!initialMatches.length) return [];

  const targets = [];
  for (const session of initialMatches) {
    const actionIds = session.actions?.transcriptActionIds ||
      (session.actions?.transcriptActionId ? [session.actions.transcriptActionId] : []);

    actionIds.forEach((_, transcriptIndex) => {
      targets.push({
        sessionKey: session.sessionKey,
        sessionStartTime: session.sessionStartTime || '',
        sessionEndTime: session.sessionEndTime || '',
        transcriptIndex
      });
    });
  }

  await appendOperation('EXACT_TRANSCRIPT_TARGETS', {
    meetingId: meeting.id,
    expectedDate: meeting.dateStamp || '',
    expectedStartTime: meeting.startTime || '',
    expectedEndTime: meeting.endTime || '',
    count: targets.length,
    targets
  });

  const files = [];

  for (let fileIndex = 0; fileIndex < targets.length; fileIndex++) {
    if (stopRequested) throw new Error('Остановлено пользователем.');

    const target = targets[fileIndex];

    // Opening a transcript changes the SPA surface and invalidates DOM action
    // references. Re-locate the exact dated/time-bounded recap cards before
    // every extraction, then select the same physical recap session again.
    if (fileIndex > 0) {
      recap = await ensureExactRecapCard(tabId, meeting, 12000);
    }

    const session = (recap?.matches || []).find(x => x.sessionKey === target.sessionKey);
    const actionIds = session?.actions?.transcriptActionIds ||
      (session?.actions?.transcriptActionId ? [session.actions.transcriptActionId] : []);
    const actionId = actionIds[target.transcriptIndex];

    if (!session || !actionId) {
      await appendOperation('EXACT_TRANSCRIPT_ACTION_MISSING', {
        meetingId: meeting.id,
        sessionKey: target.sessionKey,
        sessionStartTime: target.sessionStartTime,
        sessionEndTime: target.sessionEndTime,
        transcriptIndex: target.transcriptIndex,
        availableSessions: (recap?.matches || []).map(x => ({
          sessionKey: x.sessionKey,
          sessionStartTime: x.sessionStartTime || '',
          sessionEndTime: x.sessionEndTime || '',
          transcriptCount: x.transcriptCount || 0
        }))
      }, 'WARN');
      continue;
    }

    const file = await extractTranscriptAfterAction(
      tabId,
      meeting,
      actionId,
      fileIndex,
      'EXACT_RECAP',
      {
        sessionStartTime: target.sessionStartTime,
        sessionEndTime: target.sessionEndTime,
        transcriptIndex: target.transcriptIndex
      }
    );

    if (file) files.push(file);
  }

  return files;
}

async function tryExtractTranscriptFromMeetingDetails(tabId, meeting) {
  // Generic Details/Recap controls are unsafe for recurring meetings because
  // Teams can keep the series chat open while selecting a different occurrence.
  // Recurring occurrences are handled only through an exact dated recap card.
  if (isRecurringMeeting(meeting)) {
    await appendOperation('GENERIC_DETAILS_SKIPPED_FOR_RECURRING', {
      meetingId: meeting.id,
      expectedDate: meeting.dateStamp || ''
    }, 'INFO');
    return null;
  }

  const details = await waitForMeetingDetailsAssets(tabId, meeting, 8000);

  await appendOperation('DETAILS_TRANSCRIPT_LOOKUP', {
    meetingId: meeting.id,
    match: !!details?.match,
    found: details?.found || null,
    meta: details?.meta || null,
    actions: details?.actions || {},
    pageTitle: details?.pageTitle || '',
    url: details?.url || ''
  }, details?.found?.transcript ? 'INFO' : 'WARN');

  const actionId = details?.actions?.transcriptActionId;
  if (!actionId) return null;

  return extractTranscriptAfterAction(
    tabId,
    meeting,
    actionId,
    0,
    'DETAILS'
  );
}

async function triggerExactActionAndFindRecording(tabId, actionId, meeting) {
  const source = await chrome.tabs.get(tabId);
  const before = await chrome.tabs.query({ windowId: source.windowId });
  const beforeIds = new Set(before.map(t => t.id));

  const trigger = await sendTab(tabId, { type: 'PAGE_TRIGGER_ACTION', id: actionId }, 5000);

  await appendOperation('EXACT_RECORDING_ACTION_TRIGGER', {
    meetingId: meeting.id,
    actionId,
    trigger
  }, trigger?.ok ? 'INFO' : 'WARN');

  const started = Date.now();
  let lastTabs = [];
  let lastPage = null;

  while (Date.now() - started < 12000) {
    const tabs = await chrome.tabs.query({ windowId: source.windowId });
    lastTabs = tabs;

    const candidates = [
      ...tabs.filter(t => !beforeIds.has(t.id)),
      ...tabs.filter(t => t.id === tabId)
    ];

    for (const tab of candidates) {
      if (isRecordingUrl(tab.url)) {
        return [tab.url];
      }
    }

    // Some Teams builds keep the recording inside the same SPA page instead
    // of navigating the browser tab. Inspect links/actions that appeared
    // after the click as well.
    try {
      lastPage = await sendTab(tabId, { type: 'PAGE_FIND_RECORDINGS' }, 3500);
      const urls = (lastPage?.recordingLinks || []).map(x => x.href).filter(Boolean);
      if (urls.length) return [...new Set(urls)];
    } catch (_) {}

    await delay(400);
  }

  await appendOperation('EXACT_RECORDING_ACTION_NO_URL', {
    meetingId: meeting.id,
    actionId,
    page: lastPage ? {
      pageKind: lastPage.pageKind || '',
      url: lastPage.url || '',
      actions: (lastPage.actions || []).slice(0, 20),
      recordingLinks: (lastPage.recordingLinks || []).slice(0, 20)
    } : null,
    tabs: lastTabs.slice(0, 20).map(t => ({
      id: t.id,
      url: t.url || '',
      title: t.title || ''
    }))
  }, 'WARN');

  return [];
}

async function discoverRecordingUrls(tabId, meeting) {
  const tab = await chrome.tabs.get(tabId);

  await appendOperation('DISCOVER_RECORDING_START', {
    tabId,
    url: tab.url,
    meetingId: meeting.id,
    meetingDate: meeting.dateStamp || '',
    meetingTitle: meeting.title || ''
  });

  if (isRecordingUrl(tab.url)) return [tab.url];

  // Teams can navigate directly from Calendar to the full meeting Details
  // view. Never use generic Details assets for recurring meetings: the page can
  // display a different occurrence from the calendar item we selected.
  let details = null;
  if (!isRecurringMeeting(meeting)) {
    try {
      details = await sendTab(tabId, {
        type: 'PAGE_FIND_MEETING_DETAILS_ASSETS',
        meeting
      }, 5000);
    } catch (_) {}

    await appendOperation('MEETING_DETAILS_ASSETS', {
      meetingId: meeting.id,
      match: !!details?.match,
      pageTitle: details?.pageTitle || '',
      url: details?.url || '',
      found: details?.found || null,
      actions: details?.actions || {}
    }, details?.match ? 'INFO' : 'WARN');
  } else {
    await appendOperation('MEETING_DETAILS_ASSETS_SKIPPED_FOR_RECURRING', {
      meetingId: meeting.id,
      expectedDate: meeting.dateStamp || ''
    });
  }

  if (details?.match) {
    const actions = details.actions || {};

    if (actions.recordingActionId) {
      const urls = await triggerExactActionAndFindRecording(
        tabId,
        actions.recordingActionId,
        meeting
      );
      if (urls.length) return urls;
    }

    if (actions.recapTabActionId) {
      await sendTab(tabId, {
        type: 'PAGE_TRIGGER_ACTION',
        id: actions.recapTabActionId
      }, 5000);
      await delay(1200);

      const page = await findActionsWithWait(tabId, 10000);
      const directUrls = (page?.recordingLinks || []).map(x => x.href).filter(Boolean);
      if (directUrls.length) return [...new Set(directUrls)];
    }
  }

  const recap = await waitForExactMeetingRecap(tabId, meeting, 20000);

  await appendOperation('EXACT_RECAP_LOOKUP', {
    meetingId: meeting.id,
    pageTitle: recap?.pageTitle || '',
    url: recap?.url || '',
    match: recap?.match || null,
    matches: (recap?.matches || []).slice(0, 10)
  }, recap?.match ? 'INFO' : 'WARN');

  if (!recap?.match) return [];

  const actions = recap.match.actions || {};

  if (actions.recordingActionId) {
    const urls = await triggerExactActionAndFindRecording(
      tabId,
      actions.recordingActionId,
      meeting
    );
    if (urls.length) return urls;
  }

  if (actions.recapActionId) {
    await appendOperation('EXACT_RECAP_FALLBACK', {
      meetingId: meeting.id,
      actionId: actions.recapActionId
    }, 'WARN');

    const nav = await triggerActionAndCaptureNewTabs(tabId, actions.recapActionId);

    const candidates = [
      ...(nav.created || []),
      nav.current
    ].filter(Boolean);

    for (const candidate of candidates) {
      if (isRecordingUrl(candidate.url)) return [candidate.url];

      if (candidate.id) {
        try {
          const page = await findActionsWithWait(candidate.id, 10000);
          const urls = (page?.recordingLinks || []).map(x => x.href).filter(Boolean);
          if (urls.length) return [...new Set(urls)];
        } catch (_) {}
      }
    }
  }

  return [];
}

function normalizeMeetingKey(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[“”"'.,;:()\[\]{}]/g, '')
    .trim();
}

function meetingMatches(candidate, meeting) {
  if (!candidate || !meeting) return false;

  if (candidate.id && meeting.id && candidate.id === meeting.id) return true;
  if (candidate.label && meeting.label && candidate.label === meeting.label) return true;

  const sameDate = !meeting.dateStamp || candidate.dateStamp === meeting.dateStamp;
  const sameTitle = normalizeMeetingKey(candidate.title) === normalizeMeetingKey(meeting.title);
  if (sameDate && sameTitle) return true;

  // Time is a useful discriminator, but should not be mandatory because
  // Teams can change AM/PM formatting between scans.
  const sameTime = !meeting.startTime || candidate.startTime === meeting.startTime;
  if (sameDate && sameTime) {
    const a = normalizeMeetingKey(candidate.title);
    const m = normalizeMeetingKey(meeting.title);
    if (a && m && (a.includes(m) || m.includes(a))) return true;
  }

  return false;
}

async function scanForMeeting(tabId, meeting) {
  const response = await sendTab(tabId, { type: 'CALENDAR_SCAN' }, 5000);
  const candidates = response?.meetings || [];
  const found = candidates.find(m => meetingMatches(m, meeting)) || null;
  return { response, candidates, found };
}

async function restoreCalendarForMeeting(tabId, meeting, timeoutMs = 30000) {
  const started = Date.now();

  while (Date.now() - started < timeoutMs) {
    try {
      // First inspect the page exactly as it is. If the requested week is
      // already visible, do not touch Calendar navigation at all.
      const current = await scanForMeeting(tabId, meeting);

      if (current.response?.ok && current.found) {
        await appendOperation('CALENDAR_MEETING_ALREADY_VISIBLE', {
          meetingId: meeting.id,
          dateStamp: meeting.dateStamp,
          title: meeting.title,
          found: {
            id: current.found.id,
            title: current.found.title,
            dateStamp: current.found.dateStamp,
            startTime: current.found.startTime
          }
        });
        return {
          ok: true,
          url: current.response.url || '',
          found: current.found
        };
      }

      await appendOperation('CALENDAR_CURRENT_SCAN_NO_MATCH', {
        meetingId: meeting.id,
        dateStamp: meeting.dateStamp,
        title: meeting.title,
        candidates: current.candidates.slice(0, 30).map(x => ({
          id: x.id,
          title: x.title,
          dateStamp: x.dateStamp,
          startTime: x.startTime
        }))
      }, 'WARN');

      if (!current.candidates.length) {
        let nav = null;
        try {
          nav = await sendTab(tabId, { type: 'PAGE_OPEN_CALENDAR' }, 12000);
        } catch (e) {
          nav = { ok: false, error: e?.message || String(e) };
        }

        await appendOperation('CALENDAR_NAVIGATION_RESULT', {
          meetingId: meeting.id,
          ok: !!nav?.ok,
          alreadyCalendar: !!nav?.alreadyCalendar,
          url: nav?.url || '',
          error: nav?.error || ''
        }, nav?.ok ? 'INFO' : 'WARN');

        if (nav?.ok) {
          await delay(800);
          continue;
        }
      }

      const ensured = await sendTab(tabId, {
        type: 'CALENDAR_ENSURE_DATE',
        dateStamp: meeting.dateStamp
      }, 15000);

      await appendOperation('CALENDAR_ENSURE_DATE_RESULT', {
        meetingId: meeting.id,
        dateStamp: meeting.dateStamp,
        ok: !!ensured?.ok,
        range: ensured?.range || null,
        attempts: ensured?.attempts ?? null,
        error: ensured?.error || ''
      }, ensured?.ok ? 'INFO' : 'WARN');

      if (!ensured?.ok) {
        await delay(700);
        continue;
      }

      const after = await scanForMeeting(tabId, meeting);

      if (after.response?.ok && after.found) {
        return {
          ok: true,
          url: after.response.url || '',
          range: ensured.range || null,
          found: after.found
        };
      }

      await appendOperation('CALENDAR_SCAN_AFTER_NAV_NO_MATCH', {
        meetingId: meeting.id,
        dateStamp: meeting.dateStamp,
        title: meeting.title,
        range: ensured.range || null,
        candidates: after.candidates.slice(0, 30).map(x => ({
          id: x.id,
          title: x.title,
          dateStamp: x.dateStamp,
          startTime: x.startTime
        }))
      }, 'WARN');
    } catch (e) {
      await appendOperation('CALENDAR_ENSURE_DATE_ERROR', {
        meetingId: meeting.id,
        error: e?.message || String(e)
      }, 'WARN');
    }

    await delay(700);
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
      { type: 'CALENDAR_OPEN_MEETING_CHAT', meeting },
      15000
    );

    if (!opened?.ok) {
      await appendOperation('OPEN_MEETING_FAILED', {
        meetingId: meeting.id,
        calendarTabId,
        error: opened?.error || '',
        attempts: opened?.attempts || [],
        diagnostic: opened?.diagnostic || null,
        target: opened?.target || null
      }, 'ERROR');
      throw new Error(opened?.error || 'Не удалось открыть чат выбранной встречи из Calendar.');
    }

    await appendOperation('OPEN_MEETING_VIEW_OK', {
      meetingId: meeting.id,
      calendarTabId,
      mode: opened.mode || '',
      url: opened.url || '',
      target: opened.target || {},
      details: opened.details || null
    });

    await delay(1200);

    const files = [];

    // First bind extraction to the exact dated recap card. This is mandatory
    // for recurring meetings and also protects one-off meetings from stale UI.
    const exactFiles = await tryExtractTranscriptsFromExactRecap(
      calendarTabId,
      meeting
    );

    if (exactFiles.length) {
      files.push(...exactFiles);
    } else {
      const directFile = await tryExtractTranscriptFromMeetingDetails(
        calendarTabId,
        meeting
      );

      if (directFile) {
        files.push(directFile);
      } else {
        const urls = await discoverRecordingUrls(calendarTabId, meeting);

      await appendOperation('RECORDING_URLS_RESULT', {
        meetingId: meeting.id,
        calendarTabId,
        count: urls.length,
        urls
      });

      if (!urls.length) {
        await appendOperation('MEETING_SKIP', {
          meetingId: meeting.id,
          reason: 'Не удалось открыть Transcript или найти запись выбранной встречи.'
        }, 'WARN');

        await logMeeting(meeting.id, {
          status: 'skip',
          message: 'Не удалось открыть Transcript или найти запись выбранной встречи.',
          files: []
        });
        return;
      }

      for (let i = 0; i < urls.length; i++) {
        if (stopRequested) throw new Error('Остановлено пользователем.');

        const file = await processRecordingUrl(urls[i], meeting, i, null);
        files.push(file);
      }
      }
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

    try { await saveBatchLogs(); } catch (_) {}
  }
}

function operationLogText() {
  const lines = [
    'Teams Recap Transcript Exporter - operation log',
    `Started: ${batchState.startedAt ? new Date(batchState.startedAt).toISOString() : '-'}`,
    `Finished: ${batchState.finishedAt ? new Date(batchState.finishedAt).toISOString() : '-'}`,
    `Folder: ${batchState.folderPath || '-'}`,
    ''
  ];

  for (const op of batchState.operationLog || []) {
    lines.push(`[${op.ts || '-'}] ${op.level || 'INFO'} ${op.step || ''} ${JSON.stringify(op)}`);
  }

  return lines.join('\n');
}

async function saveBatchLogs() {
  if (!batchState.folderPath) return;

  await nativeMessage({
    action: 'saveTextInFolder',
    folderPath: batchState.folderPath,
    fileName: 'batch-report.txt',
    text: reportText()
  });

  await nativeMessage({
    action: 'saveTextInFolder',
    folderPath: batchState.folderPath,
    fileName: 'batch-operation-log.txt',
    text: operationLogText()
  });
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

    await saveBatchLogs();
  } catch (e) {
    await appendOperation('BATCH_ERROR', { error: e?.message || String(e) }, 'ERROR');
    await patchBatch({
      status: 'error',
      finishedAt: Date.now(),
      message: e?.message || String(e)
    });
    try { await saveBatchLogs(); } catch (_) {}
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
        await patchBatch({ operationLog: [] });
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