(() => {
  if (window.__teamsTranscriptCalendarLoaded) return;
  window.__teamsTranscriptCalendarLoaded = true;

  const VERSION = '2.0.4';
  let actionMap = new Map();
  let lastCalendarScanDebug = { rejectedSlots: [], candidateCount: 0, acceptedCount: 0 };

  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
  const normalize = value => (value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/\s*\n\s*/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  function hash(input) {
    let h = 2166136261;
    const s = String(input || '');
    for (let i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    return (h >>> 0).toString(16);
  }

  function isRendered(el) {
    if (!el || !(el instanceof Element) || !el.isConnected) return false;
    const style = getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden' || Number(style.opacity) === 0) return false;
    const r = el.getBoundingClientRect();
    return r.width > 15 && r.height > 12;
  }

  function accessibleText(el) {
    return normalize(
      el?.getAttribute?.('aria-label') ||
      el?.getAttribute?.('title') ||
      el?.innerText ||
      el?.textContent ||
      ''
    );
  }

  const timeRegex = /\b(?:[01]?\d|2[0-3]):[0-5]\d\b|\b(?:1[0-2]|0?[1-9]):[0-5]\d\s*(?:AM|PM)\b/i;
  const dateRegexes = [
    /\b(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\b/,
    /\b(\d{1,2})[-/.](\d{1,2})[-/.](20\d{2})\b/
  ];
  const monthNames = {
    january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
    july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
    январь: 1, февраля: 2, март: 3, апреля: 4, май: 5, июня: 6,
    июля: 7, август: 8, сентября: 9, октябрь: 10, ноября: 11, декабрь: 12
  };

  function monthNumber(value) {
    return monthNames[String(value || '').toLowerCase()] || 0;
  }

  function currentCalendarMonthYear() {
    const body = normalize(document.body?.innerText || '').slice(0, 5000);
    const names = Object.keys(monthNames).sort((a, b) => b.length - a.length).join('|');
    const re = new RegExp('\\b(' + names + ')\\s+(20\\d{2})\\b', 'i');
    const m = body.match(re);
    return m ? { month: monthNumber(m[1]), year: Number(m[2]) } : null;
  }

  function parseNamedDate(text, fallbackYear) {
    const value = normalize(text);
    const names = Object.keys(monthNames).sort((a, b) => b.length - a.length).join('|');
    const re = new RegExp('\\b(\\d{1,2})\\s+(' + names + ')(?:\\s+(20\\d{2}))?\\b', 'i');
    const m = value.match(re);
    if (!m) return null;
    return {
      day: Number(m[1]),
      month: monthNumber(m[2]),
      year: Number(m[3] || fallbackYear || 0)
    };
  }

  function parseStartTime(text) {
    const value = normalize(text);
    let m = value.match(/\b([01]?\d|2[0-3]):([0-5]\d)\b/);
    if (m) return { hour: Number(m[1]), minute: Number(m[2]) };

    m = value.match(/\b(1[0-2]|0?[1-9]):([0-5]\d)\s*(AM|PM)\b/i);
    if (!m) return null;

    let hour = Number(m[1]);
    const minute = Number(m[2]);
    const ap = m[3].toUpperCase();
    if (ap === 'AM' && hour === 12) hour = 0;
    if (ap === 'PM' && hour !== 12) hour += 12;
    return { hour, minute };
  }

  function isNonMeetingControl(text) {
    const value = normalize(text);
    if (!value) return true;

    return /^(?:join with an id|new meeting|meet now|calendar(?:\s*\([^)]*\))?|today|work week|week|day|month|previous|next)$/i.test(value)
      || /^schedule a new meeting\b/i.test(value)
      || /^use alt\+down to schedule different types of events\.?$/i.test(value)
      || /^schedule a new meeting,?\s*use alt\+down to schedule different types of events\.?$/i.test(value)
      || /^calendar\s*\(ctrl\+shift\+6\)$/i.test(value);
  }

  function getVisibleDayColumns() {
    const context = currentCalendarMonthYear();
    if (!context) return [];

    const candidates = [];
    for (const el of document.querySelectorAll('body *')) {
      if (!isRendered(el)) continue;
      const text = normalize(el.innerText || el.textContent || '');
      if (!/^\d{1,2}\s+(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)$/i.test(text)) continue;
      const day = Number(text.match(/^\d{1,2}/)?.[0] || 0);
      if (!day) continue;
      const r = el.getBoundingClientRect();
      candidates.push({
        day,
        month: context.month,
        year: context.year,
        centerX: r.left + r.width / 2,
        left: r.left,
        right: r.right
      });
    }
    return candidates;
  }

  function inferDateFromColumn(rect, dayColumns) {
    if (!dayColumns.length) return null;
    const centerX = rect.left + rect.width / 2;
    const containing = dayColumns.find(d => centerX >= d.left - 8 && centerX <= d.right + 8);
    if (containing) return containing;
    return [...dayColumns].sort((a, b) => Math.abs(a.centerX - centerX) - Math.abs(b.centerX - centerX))[0] || null;
  }

  function toDateStamp(parts) {
    if (!parts?.year || !parts?.month || !parts?.day) return '';
    return `${parts.year}${String(parts.month).padStart(2, '0')}${String(parts.day).padStart(2, '0')}`;
  }

  function toSortKey(dateParts, timeParts, fallbackIndex) {
    const y = dateParts?.year || 9999;
    const m = dateParts?.month || 12;
    const d = dateParts?.day || 31;
    const hh = timeParts?.hour ?? 23;
    const mm = timeParts?.minute ?? 59;
    return y * 100000000 + m * 1000000 + d * 10000 + hh * 100 + mm + (fallbackIndex || 0) / 100000;
  }


  function dateStampFromText(text) {
    let m = String(text || '').match(dateRegexes[0]);
    if (m) return `${m[1]}${String(m[2]).padStart(2, '0')}${String(m[3]).padStart(2, '0')}`;
    m = String(text || '').match(dateRegexes[1]);
    if (m) return `${m[3]}${String(m[2]).padStart(2, '0')}${String(m[1]).padStart(2, '0')}`;
    return '';
  }

  function isCalendarAggregateSlot(text) {
    const value = normalize(text);
    if (!value) return false;

    // Teams calendar time-grid cells expose accessibility labels such as:
    // "21 September 00:00 to 21 September 00:30. 0 events".
    // They describe a slot, not an actual meeting card.
    if (/\b\d+\s+events?\b/i.test(value)) return true;

    if (
      /\b\d{1,2}\s+[A-Za-zА-Яа-яЁё]+\s+\d{1,2}:\d{2}\s+to\s+\d{1,2}\s+[A-Za-zА-Яа-яЁё]+\s+\d{1,2}:\d{2}\b/i.test(value)
    ) return true;

    return false;
  }

  function meetingCandidateScore(el, text) {
    const attr = normalize([
      el.getAttribute('data-tid'),
      el.getAttribute('data-testid'),
      el.getAttribute('role'),
      el.className,
      el.getAttribute('aria-label')
    ].filter(Boolean).join(' '));
    const lower = `${attr} ${text}`.toLowerCase();
    if (isCalendarAggregateSlot(text) || isNonMeetingControl(text)) return -1000;
    let score = 0;

    if (/calendar|event|appointment|meeting/.test(lower)) score += 80;
    if (timeRegex.test(text)) score += 55;
    if (/organizer|attendees|busy|free|tentative|accepted|meeting/i.test(text)) score += 15;
    if (el.matches('button,a,[role="button"],[role="link"],[role="gridcell"]')) score += 20;
    if (el.closest('[role="grid"],[role="main"],main')) score += 10;

    if (/today|tomorrow|previous|next|work week|week|day|month|new event|meet now|calendar settings|join$/i.test(text)) score -= 90;
    if (/calendar/i.test(text) && text.length < 18) score -= 80;
    if (/search|activity|chat|calls|onedrive|copilot|apps/i.test(text) && text.length < 35) score -= 80;
    if (text.length < 3 || text.length > 500) score -= 100;

    const r = el.getBoundingClientRect();
    if (r.width > innerWidth * 0.85 && r.height > innerHeight * 0.5) score -= 120;
    if (r.height > 260) score -= 30;
    return score;
  }

  function cleanMeetingTitle(text) {
    let s = normalize(text);
    if (!s) return '';

    const names = Object.keys(monthNames).sort((a, b) => b.length - a.length).join('|');
    const dateTail = new RegExp(',?\\s*\\d{1,2}\\s+(?:' + names + ')\\s+20\\d{2}\\b.*$', 'i');

    s = s.replace(dateTail, '');
    s = s.replace(/,?\s*location\s*:\s*.*$/i, '');
    s = s.replace(/,?\s*(?:organised|organized) by\b.*$/i, '');
    s = s.replace(/,?\s*(?:Recurring meeting|Microsoft Teams meeting|Teams meeting).*$/i, '');
    s = s.replace(/,?\s*Press Shift\+F10 for more options.*$/i, '');
    s = s.replace(/[\s,;:-]+$/g, '').replace(/\s+/g, ' ').trim();

    if (s.length > 180) s = s.slice(0, 180).trim();
    return s || normalize(text).slice(0, 180);
  }

  function scanCalendarMeetings() {
    const pool = Array.from(document.querySelectorAll([
      '[data-tid*="calendar"]',
      '[data-tid*="event"]',
      '[data-testid*="calendar"]',
      '[data-testid*="event"]',
      '[role="gridcell"]',
      '[role="button"]',
      'button',
      'a[href]'
    ].join(',')));

    const raw = [];
    const rejectedSlots = [];
    for (const el of pool) {
      if (!isRendered(el)) continue;
      const text = accessibleText(el);

      if (isCalendarAggregateSlot(text) || isNonMeetingControl(text)) {
        if (rejectedSlots.length < 80) {
          rejectedSlots.push({
            text: text.slice(0, 300),
            tag: el.tagName,
            role: el.getAttribute('role') || '',
            dataTid: el.getAttribute('data-tid') || '',
            dataTestId: el.getAttribute('data-testid') || '',
            ariaLabel: (el.getAttribute('aria-label') || '').slice(0, 300),
            className: typeof el.className === 'string' ? el.className.slice(0, 300) : ''
          });
        }
        continue;
      }

      const score = meetingCandidateScore(el, text);
      if (score < 55) continue;

      const r = el.getBoundingClientRect();
      raw.push({ el, text, score, r });
    }

    const dayColumns = getVisibleDayColumns();

    raw.sort((a, b) => a.r.top - b.r.top || a.r.left - b.r.left || b.score - a.score);

    const labelCounts = new Map();
    const meetings = [];
    const seenBoxes = [];

    for (const item of raw) {
      const duplicateBox = seenBoxes.some(b =>
        Math.abs(b.left - item.r.left) < 3 &&
        Math.abs(b.top - item.r.top) < 3 &&
        Math.abs(b.width - item.r.width) < 5 &&
        Math.abs(b.height - item.r.height) < 5
      );
      if (duplicateBox) continue;

      const label = item.text;
      const ordinal = labelCounts.get(label) || 0;
      labelCounts.set(label, ordinal + 1);
      const id = hash(`${label}|${ordinal}`);
      const href = item.el.href || item.el.closest('a[href]')?.href || '';
      const parentText = normalize(item.el.parentElement?.innerText || '').slice(0, 700);
      const calendarContext = currentCalendarMonthYear();
      const explicitDate = parseNamedDate(`${label} ${parentText}`, calendarContext?.year)
        || (() => {
          const stamp = dateStampFromText(`${label} ${parentText}`);
          if (!stamp) return null;
          return { year: Number(stamp.slice(0, 4)), month: Number(stamp.slice(4, 6)), day: Number(stamp.slice(6, 8)) };
        })();
      const inferredDate = explicitDate || inferDateFromColumn(item.r, dayColumns);
      const startTime = parseStartTime(`${label} ${parentText}`);
      const dateStamp = toDateStamp(inferredDate);
      const startTimeText = startTime
        ? `${String(startTime.hour).padStart(2, '0')}:${String(startTime.minute).padStart(2, '0')}`
        : '';

      meetings.push({
        id,
        label,
        title: cleanMeetingTitle(label),
        dateStamp,
        startTime: startTimeText,
        sortKey: toSortKey(inferredDate, startTime, meetings.length),
        href,
        ordinal,
        score: item.score,
        dom: {
          tag: item.el.tagName,
          role: item.el.getAttribute('role') || '',
          dataTid: item.el.getAttribute('data-tid') || '',
          dataTestId: item.el.getAttribute('data-testid') || '',
          ariaLabel: item.el.getAttribute('aria-label') || '',
          className: typeof item.el.className === 'string' ? item.el.className.slice(0, 300) : ''
        },
        rect: {
          top: Math.round(item.r.top),
          left: Math.round(item.r.left),
          width: Math.round(item.r.width),
          height: Math.round(item.r.height)
        }
      });
      seenBoxes.push(item.r);
    }

    meetings.sort((a, b) => a.sortKey - b.sortKey || a.title.localeCompare(b.title, 'ru'));
    const result = meetings.slice(0, 120);
    lastCalendarScanDebug = {
      rejectedSlots,
      candidateCount: raw.length,
      acceptedCount: result.length
    };
    return result;
  }

  function findMeetingElementById(id) {
    const meetings = scanCalendarMeetings();
    const match = meetings.find(m => m.id === id);
    if (!match) return null;

    const pool = Array.from(document.querySelectorAll([
      '[data-tid*="calendar"]', '[data-tid*="event"]', '[data-testid*="calendar"]', '[data-testid*="event"]',
      '[role="gridcell"]', '[role="button"]', 'button', 'a[href]'
    ].join(',')));
    const candidates = pool.filter(el => isRendered(el) && accessibleText(el) === match.label);
    return candidates[match.ordinal] || candidates[0] || null;
  }

  async function openCalendarMeeting(id) {
    const el = findMeetingElementById(id);
    if (!el) return { ok: false, error: 'Meeting element not found in current calendar view.' };

    const target = el.matches('button,a,[role="button"],[role="link"],[tabindex]')
      ? el
      : (el.closest('button,a,[role="button"],[role="link"],[tabindex]') || el);

    try { target.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(120);

    try { target.focus({ preventScroll: true }); } catch (_) {}

    for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
      try {
        const EventCtor = type.startsWith('pointer') && window.PointerEvent ? PointerEvent : MouseEvent;
        target.dispatchEvent(new EventCtor(type, {
          bubbles: true,
          cancelable: true,
          view: window,
          button: 0,
          buttons: type.endsWith('down') ? 1 : 0
        }));
      } catch (_) {}
    }

    try { target.click(); } catch (_) {}
    await sleep(700);

    return {
      ok: true,
      url: location.href,
      target: {
        tag: target.tagName,
        role: target.getAttribute('role') || '',
        dataTid: target.getAttribute('data-tid') || '',
        dataTestId: target.getAttribute('data-testid') || '',
        ariaLabel: target.getAttribute('aria-label') || '',
        text: accessibleText(target).slice(0, 500)
      }
    };
  }

  function classifyAction(label, href) {
    const text = `${label || ''} ${href || ''}`.toLowerCase();
    if (/recap|meeting recap|summary/.test(text)) return 'recap';
    if (/recording|recorded|watch recording|meeting recording|stream\.aspx|\.mp4(?:\?|$)/.test(text)) return 'recording';
    if (/meeting chat|\bchat\b/.test(text)) return 'chat';
    if (/transcript/.test(text)) return 'transcript';
    return '';
  }

  function findActions() {
    actionMap = new Map();
    const result = [];
    const els = Array.from(document.querySelectorAll('a[href],button,[role="button"],[role="link"],[tabindex]'));
    let seq = 0;

    for (const el of els) {
      if (!isRendered(el)) continue;
      const label = accessibleText(el);
      const href = el.href || el.closest('a[href]')?.href || '';
      const kind = classifyAction(label, href);
      if (!kind) continue;
      const id = `a${++seq}-${hash(`${kind}|${label}|${href}`)}`;
      actionMap.set(id, el);
      result.push({ id, kind, label, href });
    }

    return result.slice(0, 80);
  }

  async function triggerAction(id) {
    const el = actionMap.get(id);
    if (!el || !el.isConnected) {
      findActions();
    }
    const target = actionMap.get(id);
    if (!target) return { ok: false, error: 'Action element not found.' };
    try { target.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(80);
    target.click();
    return { ok: true };
  }

  function findRecordingLinks() {
    const links = [];
    const seen = new Set();
    for (const a of document.querySelectorAll('a[href]')) {
      const href = a.href || '';
      const label = accessibleText(a);
      const lower = `${href} ${label}`.toLowerCase();
      if (!/sharepoint\.com|stream\.aspx|meeting recording|recording|\.mp4(?:\?|$)/.test(lower)) continue;
      if (!/^https?:/i.test(href)) continue;
      if (seen.has(href)) continue;
      seen.add(href);
      links.push({ href, label, kind: classifyAction(label, href) || 'recording' });
    }
    return links.slice(0, 20);
  }

  function pageKind() {
    const url = location.href.toLowerCase();
    const body = normalize(document.body?.innerText || '').slice(0, 4000).toLowerCase();
    if (/calendar/.test(url) || /\bcalendar\b/.test(body)) return 'calendar';
    if (/sharepoint\.com/.test(url) || /stream\.aspx/.test(url) || /meeting recording/.test(body)) return 'recording';
    if (/chat/.test(url) || /meeting chat/.test(body)) return 'chat';
    return 'teams';
  }

  function buildDiagnostic() {
    const meetings = scanCalendarMeetings();
    const actions = findActions();
    const links = findRecordingLinks();
    const lines = [
      'Teams Recap Transcript Exporter Calendar diagnostic',
      `Version: ${VERSION}`,
      `URL: ${location.href}`,
      `Title: ${document.title}`,
      `Viewport: ${innerWidth}x${innerHeight}`,
      `PageKind: ${pageKind()}`,
      '',
      `Meetings found: ${meetings.length}`
    ];
    meetings.forEach((m, i) => lines.push(`[M${i + 1}] id=${m.id} score=${m.score} date=${m.dateStamp || '-'} time=${m.startTime || '-'} rect=${JSON.stringify(m.rect)} dom=${JSON.stringify(m.dom || {})} label=${m.label}`));
    lines.push(
      '',
      `Calendar scan: candidates=${lastCalendarScanDebug.candidateCount || 0}, accepted=${lastCalendarScanDebug.acceptedCount || 0}, rejectedSlots=${(lastCalendarScanDebug.rejectedSlots || []).length}`
    );
    (lastCalendarScanDebug.rejectedSlots || []).slice(0, 40).forEach((x, i) =>
      lines.push(`[REJECTED_SLOT ${i + 1}] ${JSON.stringify(x)}`)
    );
    lines.push('', `Actions found: ${actions.length}`);
    actions.forEach((a, i) => lines.push(`[A${i + 1}] kind=${a.kind} label=${a.label} href=${a.href || '-'}`));
    lines.push('', `Recording links found: ${links.length}`);
    links.forEach((l, i) => lines.push(`[R${i + 1}] kind=${l.kind} label=${l.label} href=${l.href}`));
    lines.push('', 'Visible page text sample:', normalize(document.body?.innerText || '').slice(0, 12000));
    return lines.join('\n');
  }

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    const type = message?.type;

    if (type === 'CALENDAR_SCAN') {
      sendResponse({ ok: true, meetings: scanCalendarMeetings(), pageKind: pageKind() });
      return;
    }
    if (type === 'CALENDAR_OPEN_MEETING') {
      openCalendarMeeting(message.id).then(sendResponse);
      return true;
    }
    if (type === 'CALENDAR_FIND_ACTIONS' || type === 'PAGE_FIND_ACTIONS') {
      sendResponse({ ok: true, actions: findActions(), recordingLinks: findRecordingLinks(), pageKind: pageKind(), url: location.href });
      return;
    }
    if (type === 'CALENDAR_TRIGGER_ACTION' || type === 'PAGE_TRIGGER_ACTION') {
      triggerAction(message.id).then(sendResponse);
      return true;
    }
    if (type === 'PAGE_FIND_RECORDINGS') {
      sendResponse({ ok: true, recordingLinks: findRecordingLinks(), actions: findActions(), pageKind: pageKind(), url: location.href });
      return;
    }
    if (type === 'CALENDAR_DEBUG') {
      sendResponse({ ok: true, text: buildDiagnostic() });
      return;
    }
  });
})();
, 'i');
    s = s.replace(dateTail, '');

    s = s.replace(/\b(?:[01]?\d|2[0-3]):[0-5]\d\b\s*(?:-|–|—|to)?\s*\b(?:[01]?\d|2[0-3]):[0-5]\d\b/gi, ' ');
    s = s.replace(/\b(?:1[0-2]|0?[1-9]):[0-5]\d\s*(?:AM|PM)\b\s*(?:-|–|—|to)?\s*\b(?:1[0-2]|0?[1-9]):[0-5]\d\s*(?:AM|PM)\b/gi, ' ');
    s = s.replace(/\b(?:Accepted|Tentative|Busy|Free|Organizer)\b/gi, ' ');
    s = s.replace(/[\s,;:-]+$/g, '').replace(/\s+/g, ' ').trim();

    if (s.length > 180) s = s.slice(0, 180).trim();
    return s || normalize(text).slice(0, 180);
  }

  function scanCalendarMeetings() {
    const pool = Array.from(document.querySelectorAll([
      '[data-tid*="calendar"]',
      '[data-tid*="event"]',
      '[data-testid*="calendar"]',
      '[data-testid*="event"]',
      '[role="gridcell"]',
      '[role="button"]',
      'button',
      'a[href]'
    ].join(',')));

    const raw = [];
    const rejectedSlots = [];
    for (const el of pool) {
      if (!isRendered(el)) continue;
      const text = accessibleText(el);

      if (isCalendarAggregateSlot(text) || isNonMeetingControl(text)) {
        if (rejectedSlots.length < 80) {
          rejectedSlots.push({
            text: text.slice(0, 300),
            tag: el.tagName,
            role: el.getAttribute('role') || '',
            dataTid: el.getAttribute('data-tid') || '',
            dataTestId: el.getAttribute('data-testid') || '',
            ariaLabel: (el.getAttribute('aria-label') || '').slice(0, 300),
            className: typeof el.className === 'string' ? el.className.slice(0, 300) : ''
          });
        }
        continue;
      }

      const score = meetingCandidateScore(el, text);
      if (score < 55) continue;

      const r = el.getBoundingClientRect();
      raw.push({ el, text, score, r });
    }

    const dayColumns = getVisibleDayColumns();

    raw.sort((a, b) => a.r.top - b.r.top || a.r.left - b.r.left || b.score - a.score);

    const labelCounts = new Map();
    const meetings = [];
    const seenBoxes = [];

    for (const item of raw) {
      const duplicateBox = seenBoxes.some(b =>
        Math.abs(b.left - item.r.left) < 3 &&
        Math.abs(b.top - item.r.top) < 3 &&
        Math.abs(b.width - item.r.width) < 5 &&
        Math.abs(b.height - item.r.height) < 5
      );
      if (duplicateBox) continue;

      const label = item.text;
      const ordinal = labelCounts.get(label) || 0;
      labelCounts.set(label, ordinal + 1);
      const id = hash(`${label}|${ordinal}`);
      const href = item.el.href || item.el.closest('a[href]')?.href || '';
      const parentText = normalize(item.el.parentElement?.innerText || '').slice(0, 700);
      const calendarContext = currentCalendarMonthYear();
      const explicitDate = parseNamedDate(`${label} ${parentText}`, calendarContext?.year)
        || (() => {
          const stamp = dateStampFromText(`${label} ${parentText}`);
          if (!stamp) return null;
          return { year: Number(stamp.slice(0, 4)), month: Number(stamp.slice(4, 6)), day: Number(stamp.slice(6, 8)) };
        })();
      const inferredDate = explicitDate || inferDateFromColumn(item.r, dayColumns);
      const startTime = parseStartTime(`${label} ${parentText}`);
      const dateStamp = toDateStamp(inferredDate);
      const startTimeText = startTime
        ? `${String(startTime.hour).padStart(2, '0')}:${String(startTime.minute).padStart(2, '0')}`
        : '';

      meetings.push({
        id,
        label,
        title: cleanMeetingTitle(label),
        dateStamp,
        startTime: startTimeText,
        sortKey: toSortKey(inferredDate, startTime, meetings.length),
        href,
        ordinal,
        score: item.score,
        dom: {
          tag: item.el.tagName,
          role: item.el.getAttribute('role') || '',
          dataTid: item.el.getAttribute('data-tid') || '',
          dataTestId: item.el.getAttribute('data-testid') || '',
          ariaLabel: item.el.getAttribute('aria-label') || '',
          className: typeof item.el.className === 'string' ? item.el.className.slice(0, 300) : ''
        },
        rect: {
          top: Math.round(item.r.top),
          left: Math.round(item.r.left),
          width: Math.round(item.r.width),
          height: Math.round(item.r.height)
        }
      });
      seenBoxes.push(item.r);
    }

    meetings.sort((a, b) => a.sortKey - b.sortKey || a.title.localeCompare(b.title, 'ru'));
    const result = meetings.slice(0, 120);
    lastCalendarScanDebug = {
      rejectedSlots,
      candidateCount: raw.length,
      acceptedCount: result.length
    };
    return result;
  }

  function findMeetingElementById(id) {
    const meetings = scanCalendarMeetings();
    const match = meetings.find(m => m.id === id);
    if (!match) return null;

    const pool = Array.from(document.querySelectorAll([
      '[data-tid*="calendar"]', '[data-tid*="event"]', '[data-testid*="calendar"]', '[data-testid*="event"]',
      '[role="gridcell"]', '[role="button"]', 'button', 'a[href]'
    ].join(',')));
    const candidates = pool.filter(el => isRendered(el) && accessibleText(el) === match.label);
    return candidates[match.ordinal] || candidates[0] || null;
  }

  async function openCalendarMeeting(id) {
    const el = findMeetingElementById(id);
    if (!el) return { ok: false, error: 'Meeting element not found in current calendar view.' };

    const target = el.matches('button,a,[role="button"],[role="link"],[tabindex]')
      ? el
      : (el.closest('button,a,[role="button"],[role="link"],[tabindex]') || el);

    try { target.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(120);

    try { target.focus({ preventScroll: true }); } catch (_) {}

    for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
      try {
        const EventCtor = type.startsWith('pointer') && window.PointerEvent ? PointerEvent : MouseEvent;
        target.dispatchEvent(new EventCtor(type, {
          bubbles: true,
          cancelable: true,
          view: window,
          button: 0,
          buttons: type.endsWith('down') ? 1 : 0
        }));
      } catch (_) {}
    }

    try { target.click(); } catch (_) {}
    await sleep(700);

    return {
      ok: true,
      url: location.href,
      target: {
        tag: target.tagName,
        role: target.getAttribute('role') || '',
        dataTid: target.getAttribute('data-tid') || '',
        dataTestId: target.getAttribute('data-testid') || '',
        ariaLabel: target.getAttribute('aria-label') || '',
        text: accessibleText(target).slice(0, 500)
      }
    };
  }

  function classifyAction(label, href) {
    const text = `${label || ''} ${href || ''}`.toLowerCase();
    if (/recap|meeting recap|summary/.test(text)) return 'recap';
    if (/recording|recorded|watch recording|meeting recording|stream\.aspx|\.mp4(?:\?|$)/.test(text)) return 'recording';
    if (/meeting chat|\bchat\b/.test(text)) return 'chat';
    if (/transcript/.test(text)) return 'transcript';
    return '';
  }

  function findActions() {
    actionMap = new Map();
    const result = [];
    const els = Array.from(document.querySelectorAll('a[href],button,[role="button"],[role="link"],[tabindex]'));
    let seq = 0;

    for (const el of els) {
      if (!isRendered(el)) continue;
      const label = accessibleText(el);
      const href = el.href || el.closest('a[href]')?.href || '';
      const kind = classifyAction(label, href);
      if (!kind) continue;
      const id = `a${++seq}-${hash(`${kind}|${label}|${href}`)}`;
      actionMap.set(id, el);
      result.push({ id, kind, label, href });
    }

    return result.slice(0, 80);
  }

  async function triggerAction(id) {
    const el = actionMap.get(id);
    if (!el || !el.isConnected) {
      findActions();
    }
    const target = actionMap.get(id);
    if (!target) return { ok: false, error: 'Action element not found.' };
    try { target.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(80);
    target.click();
    return { ok: true };
  }

  function findRecordingLinks() {
    const links = [];
    const seen = new Set();
    for (const a of document.querySelectorAll('a[href]')) {
      const href = a.href || '';
      const label = accessibleText(a);
      const lower = `${href} ${label}`.toLowerCase();
      if (!/sharepoint\.com|stream\.aspx|meeting recording|recording|\.mp4(?:\?|$)/.test(lower)) continue;
      if (!/^https?:/i.test(href)) continue;
      if (seen.has(href)) continue;
      seen.add(href);
      links.push({ href, label, kind: classifyAction(label, href) || 'recording' });
    }
    return links.slice(0, 20);
  }

  function pageKind() {
    const url = location.href.toLowerCase();
    const body = normalize(document.body?.innerText || '').slice(0, 4000).toLowerCase();
    if (/calendar/.test(url) || /\bcalendar\b/.test(body)) return 'calendar';
    if (/sharepoint\.com/.test(url) || /stream\.aspx/.test(url) || /meeting recording/.test(body)) return 'recording';
    if (/chat/.test(url) || /meeting chat/.test(body)) return 'chat';
    return 'teams';
  }

  function buildDiagnostic() {
    const meetings = scanCalendarMeetings();
    const actions = findActions();
    const links = findRecordingLinks();
    const lines = [
      'Teams Recap Transcript Exporter Calendar diagnostic',
      `Version: ${VERSION}`,
      `URL: ${location.href}`,
      `Title: ${document.title}`,
      `Viewport: ${innerWidth}x${innerHeight}`,
      `PageKind: ${pageKind()}`,
      '',
      `Meetings found: ${meetings.length}`
    ];
    meetings.forEach((m, i) => lines.push(`[M${i + 1}] id=${m.id} score=${m.score} date=${m.dateStamp || '-'} time=${m.startTime || '-'} rect=${JSON.stringify(m.rect)} dom=${JSON.stringify(m.dom || {})} label=${m.label}`));
    lines.push(
      '',
      `Calendar scan: candidates=${lastCalendarScanDebug.candidateCount || 0}, accepted=${lastCalendarScanDebug.acceptedCount || 0}, rejectedSlots=${(lastCalendarScanDebug.rejectedSlots || []).length}`
    );
    (lastCalendarScanDebug.rejectedSlots || []).slice(0, 40).forEach((x, i) =>
      lines.push(`[REJECTED_SLOT ${i + 1}] ${JSON.stringify(x)}`)
    );
    lines.push('', `Actions found: ${actions.length}`);
    actions.forEach((a, i) => lines.push(`[A${i + 1}] kind=${a.kind} label=${a.label} href=${a.href || '-'}`));
    lines.push('', `Recording links found: ${links.length}`);
    links.forEach((l, i) => lines.push(`[R${i + 1}] kind=${l.kind} label=${l.label} href=${l.href}`));
    lines.push('', 'Visible page text sample:', normalize(document.body?.innerText || '').slice(0, 12000));
    return lines.join('\n');
  }

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    const type = message?.type;

    if (type === 'CALENDAR_SCAN') {
      sendResponse({ ok: true, meetings: scanCalendarMeetings(), pageKind: pageKind() });
      return;
    }
    if (type === 'CALENDAR_OPEN_MEETING') {
      openCalendarMeeting(message.id).then(sendResponse);
      return true;
    }
    if (type === 'CALENDAR_FIND_ACTIONS' || type === 'PAGE_FIND_ACTIONS') {
      sendResponse({ ok: true, actions: findActions(), recordingLinks: findRecordingLinks(), pageKind: pageKind(), url: location.href });
      return;
    }
    if (type === 'CALENDAR_TRIGGER_ACTION' || type === 'PAGE_TRIGGER_ACTION') {
      triggerAction(message.id).then(sendResponse);
      return true;
    }
    if (type === 'PAGE_FIND_RECORDINGS') {
      sendResponse({ ok: true, recordingLinks: findRecordingLinks(), actions: findActions(), pageKind: pageKind(), url: location.href });
      return;
    }
    if (type === 'CALENDAR_DEBUG') {
      sendResponse({ ok: true, text: buildDiagnostic() });
      return;
    }
  });
})();
