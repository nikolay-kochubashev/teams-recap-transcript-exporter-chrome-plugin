(() => {
  if (window.__teamsTranscriptCalendarLoaded) return;
  window.__teamsTranscriptCalendarLoaded = true;

  const VERSION = '2.0.21';
  let actionMap = new Map();
  let lastCalendarScanDebug = { rejected: [], candidates: [], acceptedCount: 0 };

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

  const monthNames = {
    january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
    july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
    января: 1, февраля: 2, марта: 3, апреля: 4, мая: 5, июня: 6,
    июля: 7, августа: 8, сентября: 9, октября: 10, ноября: 11, декабря: 12,
    январь: 1, февраль: 2, март: 3, апрель: 4, май: 5, июнь: 6,
    июль: 7, август: 8, сентябрь: 9, октябрь: 10, ноябрь: 11, декабрь: 12
  };

  const monthPattern = Object.keys(monthNames).sort((a, b) => b.length - a.length).join('|');
  const timeRegex = /\b(?:[01]?\d|2[0-3]):[0-5]\d\b|\b(?:1[0-2]|0?[1-9]):[0-5]\d\s*(?:AM|PM)\b/i;

  function monthNumber(value) {
    return monthNames[String(value || '').toLowerCase()] || 0;
  }

  function currentCalendarMonthYear() {
    const body = normalize(document.body?.innerText || '').slice(0, 7000);
    const re = new RegExp('\\b(' + monthPattern + ')\\s+(20\\d{2})\\b', 'i');
    const m = body.match(re);
    return m ? { month: monthNumber(m[1]), year: Number(m[2]) } : null;
  }

  function parseNamedDate(text, fallbackYear) {
    const value = normalize(text);
    const re = new RegExp('\\b(\\d{1,2})\\s+(' + monthPattern + ')(?:\\s+(20\\d{2}))?\\b', 'i');
    const m = value.match(re);
    if (!m) return null;
    return {
      day: Number(m[1]),
      month: monthNumber(m[2]),
      year: Number(m[3] || fallbackYear || 0)
    };
  }

  function parseFlexibleDate(text, fallbackYear) {
    const value = normalize(text);
    const dayFirst = parseNamedDate(value, fallbackYear);
    if (dayFirst) return dayFirst;

    const re = new RegExp('\\b(' + monthPattern + ')\\s+(\\d{1,2})(?:,)?(?:\\s+(20\\d{2}))?\\b', 'i');
    const m = value.match(re);
    if (!m) return null;
    return {
      day: Number(m[2]),
      month: monthNumber(m[1]),
      year: Number(m[3] || fallbackYear || 0)
    };
  }

  function parseDateStampParts(stamp) {
    const value = String(stamp ?? '').trim().replace(/[^0-9]/g, '');
    if (!/^[0-9]{8}$/.test(value)) return null;
    const year = Number(value.slice(0, 4));
    const month = Number(value.slice(4, 6));
    const day = Number(value.slice(6, 8));
    if (year < 2000 || month < 1 || month > 12 || day < 1 || day > 31) return null;
    return { year, month, day };
  }

  function datePartsToUtc(parts) {
    if (!parts?.year || !parts?.month || !parts?.day) return NaN;
    return Date.UTC(parts.year, parts.month - 1, parts.day);
  }

  function parseStartTime(text) {
    const value = normalize(text);

    // Parse 12-hour clock first. Otherwise "3:00 PM" is prematurely
    // interpreted by the generic 24-hour regex as 03:00.
    let m = value.match(/\b(1[0-2]|0?[1-9]):([0-5]\d)\s*(AM|PM)\b/i);
    if (m) {
      let hour = Number(m[1]);
      const minute = Number(m[2]);
      const ap = m[3].toUpperCase();
      if (ap === 'AM' && hour === 12) hour = 0;
      if (ap === 'PM' && hour !== 12) hour += 12;
      return { hour, minute };
    }

    m = value.match(/\b([01]?\d|2[0-3]):([0-5]\d)\b/);
    if (m) return { hour: Number(m[1]), minute: Number(m[2]) };

    return null;
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

  function isCalendarAggregateSlot(text) {
    const value = normalize(text);
    if (!value) return false;
    if (/\b\d+\s+events?\b/i.test(value)) return true;
    if (/\b\d{1,2}\s+[A-Za-zА-Яа-яЁё]+\s+\d{1,2}:\d{2}\s+to\s+\d{1,2}\s+[A-Za-zА-Яа-яЁё]+\s+\d{1,2}:\d{2}\b/i.test(value)) return true;
    return false;
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

  function cleanMeetingTitle(text) {
    let s = normalize(text);
    if (!s) return '';

    const dayFirstTail = /,?\s*\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December|января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+20\d{2}\b.*$/i;
    const monthFirstTail = /,?\s*(?:(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s*)?(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+20\d{2}\b.*$/i;

    s = s.replace(dayFirstTail, '');
    s = s.replace(monthFirstTail, '');
    s = s.replace(/,?\s*location\s*:\s*.*$/i, '');
    s = s.replace(/,?\s*(?:organised|organized) by\b.*$/i, '');
    s = s.replace(/,?\s*(?:Recurring meeting|Microsoft Teams meeting|Teams meeting).*$/i, '');
    s = s.replace(/,?\s*Press Shift\+F10 for more options.*$/i, '');
    s = s.replace(/[\s,;:-]+$/g, '').replace(/\s+/g, ' ').trim();

    return (s || normalize(text)).slice(0, 180);
  }

  function getVisibleDayColumns() {
    const context = currentCalendarMonthYear();
    if (!context) return [];

    const columns = [];
    for (const el of document.querySelectorAll('body *')) {
      if (!isRendered(el)) continue;
      const text = normalize(el.innerText || el.textContent || '');
      if (!/^\d{1,2}\s+(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)$/i.test(text)) continue;
      const day = Number(text.match(/^\d{1,2}/)?.[0] || 0);
      if (!day) continue;
      const r = el.getBoundingClientRect();
      columns.push({
        day,
        month: context.month,
        year: context.year,
        centerX: r.left + r.width / 2,
        left: r.left,
        right: r.right
      });
    }
    return columns;
  }

  function inferDateFromColumn(rect, dayColumns) {
    if (!dayColumns.length) return null;
    const centerX = rect.left + rect.width / 2;
    const containing = dayColumns.find(d => centerX >= d.left - 8 && centerX <= d.right + 8);
    if (containing) return containing;
    return [...dayColumns].sort((a, b) => Math.abs(a.centerX - centerX) - Math.abs(b.centerX - centerX))[0] || null;
  }

  function meetingCandidateScore(el, text) {
    if (isCalendarAggregateSlot(text) || isNonMeetingControl(text)) return -1000;

    const attr = normalize([
      el.getAttribute('data-tid'),
      el.getAttribute('data-testid'),
      el.getAttribute('role'),
      typeof el.className === 'string' ? el.className : '',
      el.getAttribute('aria-label')
    ].filter(Boolean).join(' '));
    const lower = `${attr} ${text}`.toLowerCase();

    let score = 0;
    if (/calendar|event|appointment|meeting/.test(lower)) score += 80;
    if (timeRegex.test(text)) score += 55;
    if (/organizer|organised|organized|attendees|busy|free|tentative|accepted|recurring meeting/i.test(text)) score += 25;
    if (el.matches('button,a,[role="button"],[role="link"],[role="gridcell"]')) score += 20;
    if (el.closest('[role="grid"],[role="main"],main')) score += 10;

    if (/today|tomorrow|previous|next|work week|week|day|month|new event|meet now|calendar settings|join$/i.test(text)) score -= 90;
    if (/search|activity|chat|calls|onedrive|copilot|apps/i.test(text) && text.length < 35) score -= 80;
    if (text.length < 3 || text.length > 650) score -= 100;

    const r = el.getBoundingClientRect();
    if (r.width > innerWidth * 0.85 && r.height > innerHeight * 0.5) score -= 120;
    if (r.height > 300) score -= 35;
    return score;
  }

  function scanCalendarMeetings() {
    const cards = Array.from(
      document.querySelectorAll('[data-testid="calendar-in-day-event-card"]')
    ).filter(isRendered);

    const calendarContext = currentCalendarMonthYear();
    const meetings = [];
    const debugCandidates = [];

    for (const el of cards) {
      const label = normalize(el.getAttribute('aria-label') || accessibleText(el));
      if (!label) continue;

      const date = parseFlexibleDate(label, calendarContext?.year);
      const startTime = parseStartTime(label);
      const dateStamp = toDateStamp(date);
      const startTimeText = startTime
        ? `${String(startTime.hour).padStart(2, '0')}:${String(startTime.minute).padStart(2, '0')}`
        : '';
      const title = cleanMeetingTitle(label);
      const r = el.getBoundingClientRect();
      const elementId = el.id || '';
      const id = elementId || hash(`${label}|${dateStamp}|${startTimeText}`);

      const meeting = {
        id,
        label,
        title,
        dateStamp,
        startTime: startTimeText,
        sortKey: toSortKey(date, startTime, meetings.length),
        href: '',
        score: 1000,
        dom: {
          adapter: 'calendar-in-day-event-card',
          elementId,
          dataTestId: 'calendar-in-day-event-card',
          ariaLabel: label
        },
        rect: {
          top: Math.round(r.top),
          left: Math.round(r.left),
          width: Math.round(r.width),
          height: Math.round(r.height)
        }
      };

      meetings.push(meeting);
      if (debugCandidates.length < 80) {
        debugCandidates.push({
          id,
          title,
          dateStamp,
          startTime: startTimeText,
          elementId,
          ariaLabel: label
        });
      }
    }

    meetings.sort((a, b) => a.sortKey - b.sortKey || a.title.localeCompare(b.title, 'ru'));

    lastCalendarScanDebug = {
      rejected: [],
      candidates: debugCandidates,
      acceptedCount: meetings.length
    };

    return meetings.slice(0, 120);
  }

  function findMeetingElement(meeting) {
    if (!meeting) return null;

    const byId = meeting.dom?.elementId
      ? document.getElementById(meeting.dom.elementId)
      : null;

    if (
      byId &&
      byId.matches?.('[data-testid="calendar-in-day-event-card"]') &&
      isRendered(byId)
    ) {
      return byId;
    }

    const cards = Array.from(
      document.querySelectorAll('[data-testid="calendar-in-day-event-card"]')
    ).filter(isRendered);

    const exact = cards.find(el =>
      normalize(el.getAttribute('aria-label') || '') === normalize(meeting.label)
    );
    if (exact) return exact;

    const expectedTitle = normalizedComparable(meeting.title);
    return cards.find(el => {
      const label = normalize(el.getAttribute('aria-label') || '');
      const title = normalizedComparable(cleanMeetingTitle(label));
      const dateStamp = toDateStamp(parseFlexibleDate(label, currentCalendarMonthYear()?.year));
      return title === expectedTitle && dateStamp === meeting.dateStamp;
    }) || null;
  }

  function interactiveText(el) {
    return normalize(
      el?.getAttribute?.('aria-label') ||
      el?.getAttribute?.('title') ||
      el?.innerText ||
      el?.textContent ||
      ''
    );
  }

  function findChatWithParticipants(meeting) {
    const pattern = /^(?:chat with participants|чат с участниками)$/i;
    const expectedTitle = normalizedComparable(meeting?.title || '');

    const candidates = Array.from(document.querySelectorAll(
      'a[href],button,[role="button"],[role="link"],[tabindex]'
    ))
      .filter(isRendered)
      .filter(el => pattern.test(interactiveText(el)));

    if (!candidates.length) return null;
    if (candidates.length === 1) return candidates[0];

    const scoped = candidates.find(el => {
      let node = el;
      for (let i = 0; i < 8 && node; i++, node = node.parentElement) {
        const text = normalizedComparable(node.innerText || node.textContent || '');
        if (expectedTitle && text.includes(expectedTitle)) return true;
      }
      return false;
    });

    return scoped || candidates[0];
  }

  function meetingDetailsViewMatches(meeting) {
    const expected = normalizedComparable(meeting?.title || '');
    if (!expected) return false;

    const titleText = normalizedComparable(document.title || '');
    const bodyText = normalizedComparable(
      (document.body?.innerText || document.body?.textContent || '').slice(0, 12000)
    );

    const titleMatches = titleText.includes(expected);
    const bodyMatches = bodyText.includes(expected);
    if (!titleMatches && !bodyMatches) return false;

    const controls = Array.from(document.querySelectorAll(
      'button,a,[role="button"],[role="link"],[role="tab"],[tabindex]'
    ))
      .filter(isRendered)
      .map(interactiveText)
      .filter(Boolean);

    const hasDetailsSignals = controls.some(text =>
      /^(?:details|recap|recording|transcript|close)$/i.test(text)
    );

    const hasVisibleCalendarCards = Array.from(document.querySelectorAll(
      '[data-testid="calendar-in-day-event-card"]'
    )).some(isRendered);

    // Never accept a stale previous meeting merely because the requested title
    // is somewhere in the left chat list/body. The document title is the most
    // stable identity signal in this Teams build.
    if (titleMatches && !hasVisibleCalendarCards) return true;

    const headingMatches = Array.from(document.querySelectorAll(
      'h1,h2,h3,[role="heading"]'
    ))
      .filter(isRendered)
      .some(el => normalizedComparable(el.innerText || el.textContent || '').includes(expected));

    if (headingMatches && !hasVisibleCalendarCards && hasDetailsSignals) return true;

    return false;
  }

  function clickableAncestor(el, maxDepth = 6) {
    let node = el;
    for (let i = 0; node && i <= maxDepth; i++, node = node.parentElement) {
      if (!(node instanceof Element)) break;
      if (
        node.matches('button,a[href],[role="button"],[role="link"],[tabindex]') &&
        isRendered(node)
      ) {
        return node;
      }
    }
    return null;
  }

  function findClickableByExactText(pattern) {
    const direct = Array.from(document.querySelectorAll(
      'button,a[href],[role="button"],[role="link"],[role="tab"],[tabindex]'
    ))
      .filter(isRendered)
      .find(el => pattern.test(interactiveText(el)));
    if (direct) return direct;

    const all = Array.from(document.querySelectorAll('body *'))
      .filter(isRendered)
      .filter(el => pattern.test(normalize(el.innerText || el.textContent || '')))
      .sort((a, b) => {
        const ar = a.getBoundingClientRect();
        const br = b.getBoundingClientRect();
        return (ar.width * ar.height) - (br.width * br.height);
      });

    for (const el of all) {
      const clickable = clickableAncestor(el);
      if (clickable) return clickable;
    }
    return null;
  }

  function findClickableContainingText(pattern) {
    const candidates = Array.from(document.querySelectorAll(
      'button,a[href],[role="button"],[role="link"],[role="tab"],[tabindex]'
    ))
      .filter(isRendered)
      .map(el => ({ el, text: interactiveText(el) }))
      .filter(x => x.text && x.text.length <= 180 && pattern.test(x.text))
      .sort((a, b) => {
        const ar = a.el.getBoundingClientRect();
        const br = b.el.getBoundingClientRect();
        return (ar.width * ar.height) - (br.width * br.height);
      });
    if (candidates.length) return candidates[0].el;

    const nodes = Array.from(document.querySelectorAll('body *'))
      .filter(isRendered)
      .map(el => ({ el, text: normalize(el.innerText || el.textContent || '') }))
      .filter(x => x.text && x.text.length <= 180 && pattern.test(x.text))
      .sort((a, b) => {
        const ar = a.el.getBoundingClientRect();
        const br = b.el.getBoundingClientRect();
        return (ar.width * ar.height) - (br.width * br.height);
      });

    for (const x of nodes) {
      const clickable = clickableAncestor(x.el, 8);
      if (clickable) return clickable;
    }
    return null;
  }

  function findMeetingDetailsAssets(meeting) {
    actionMap = new Map();

    const matches = meetingDetailsViewMatches(meeting);
    if (!matches) {
      return {
        ok: true,
        match: false,
        pageTitle: document.title,
        url: location.href,
        actions: {}
      };
    }

    const recording =
      findClickableByExactText(/^(?:recording|запись)(?:\s|$)/i) ||
      findClickableContainingText(/\b(?:recording|запись)\b/i);
    const transcript =
      findClickableByExactText(/^(?:transcript|транскрипт|расшифровка)(?:\s|$)/i) ||
      findClickableContainingText(/\b(?:transcript|транскрипт|расшифровка)\b/i);
    const recapTab = findClickableByExactText(/^recap$/i);
    const chatTab = findClickableByExactText(/^chat$/i);

    const actions = {};
    for (const [kind, el] of [
      ['recording', recording],
      ['transcript', transcript],
      ['recapTab', recapTab],
      ['chatTab', chatTab]
    ]) {
      if (!el) continue;
      const id = `details-${kind}-${hash(`${meeting?.id || ''}|${kind}|${interactiveText(el)}`)}`;
      actionMap.set(id, el);
      actions[`${kind}ActionId`] = id;
    }

    return {
      ok: true,
      match: true,
      pageTitle: document.title,
      url: location.href,
      actions,
      found: {
        recording: !!recording,
        transcript: !!transcript,
        recapTab: !!recapTab,
        chatTab: !!chatTab
      },
      meta: {
        recordingText: recording ? interactiveText(recording) : '',
        transcriptText: transcript ? interactiveText(transcript) : '',
        recordingTag: recording?.tagName || '',
        transcriptTag: transcript?.tagName || '',
        recordingRole: recording?.getAttribute?.('role') || '',
        transcriptRole: transcript?.getAttribute?.('role') || ''
      }
    };
  }

  async function openCalendarNavigation() {
    if (document.querySelector('[data-testid="calendar-in-day-event-card"]')) {
      return { ok: true, alreadyCalendar: true, url: location.href };
    }

    const candidates = Array.from(document.querySelectorAll(
      'a[href],button,[role="button"],[role="link"],[tabindex]'
    )).filter(isRendered);

    const calendar = candidates.find(el =>
      /^(?:calendar|calendar \(ctrl\+shift\+6\)|календарь)$/i.test(interactiveText(el))
    );

    if (!calendar) {
      return {
        ok: false,
        error: 'Calendar navigation control not found.',
        url: location.href,
        title: document.title
      };
    }

    try { calendar.click(); } catch (_) { dispatchPointerSequence(calendar); }

    const started = Date.now();
    while (Date.now() - started < 10000) {
      if (
        document.querySelector('[data-testid="calendar-in-day-event-card"]') ||
        /^Calendar$/i.test(normalize(document.querySelector('h1,h2,[role="heading"]')?.textContent || ''))
      ) {
        return { ok: true, alreadyCalendar: false, url: location.href };
      }
      await sleep(200);
    }

    return {
      ok: false,
      error: 'Calendar did not become ready after navigation click.',
      url: location.href,
      title: document.title
    };
  }

  function popupDiagnostic(meeting) {
    const chatCandidates = Array.from(document.querySelectorAll(
      'a[href],button,[role="button"],[role="link"],[tabindex]'
    ))
      .filter(isRendered)
      .map(el => interactiveText(el))
      .filter(Boolean)
      .filter(text => /chat|participants|join|edit/i.test(text))
      .slice(0, 30);

    const peeks = Array.from(document.querySelectorAll('[data-tid="peek-body-container"]'))
      .filter(isRendered)
      .map(el => normalize(el.innerText || el.textContent || '').slice(0, 1200));

    const visibleMeetingCards = Array.from(
      document.querySelectorAll('[data-testid="calendar-in-day-event-card"]')
    ).filter(isRendered).length;

    const controls = Array.from(document.querySelectorAll(
      'button,a,[role="button"],[role="link"],[role="tab"],[tabindex]'
    ))
      .filter(isRendered)
      .map(interactiveText)
      .filter(Boolean)
      .slice(0, 80);

    return {
      expectedTitle: meeting?.title || '',
      documentTitle: document.title,
      url: location.href,
      visibleMeetingCards,
      detailsMatch: meetingDetailsViewMatches(meeting),
      peekCount: peeks.length,
      peeks,
      chatCandidates,
      controls
    };
  }

  async function waitForChatAction(meeting, timeoutMs = 1800) {
    const started = Date.now();
    while (Date.now() - started < timeoutMs) {
      const action = findChatWithParticipants(meeting);
      if (action) return action;
      await sleep(120);
    }
    return null;
  }

  function dispatchPointerSequence(el) {
    const r = el.getBoundingClientRect();
    const opts = {
      bubbles: true,
      cancelable: true,
      composed: true,
      view: window,
      clientX: r.left + Math.max(1, r.width / 2),
      clientY: r.top + Math.max(1, r.height / 2),
      button: 0,
      buttons: 1,
      detail: 1
    };

    try { el.dispatchEvent(new PointerEvent('pointerdown', { ...opts, pointerId: 1, pointerType: 'mouse', isPrimary: true })); } catch (_) {}
    try { el.dispatchEvent(new MouseEvent('mousedown', opts)); } catch (_) {}
    try { el.dispatchEvent(new PointerEvent('pointerup', { ...opts, pointerId: 1, pointerType: 'mouse', isPrimary: true, buttons: 0 })); } catch (_) {}
    try { el.dispatchEvent(new MouseEvent('mouseup', { ...opts, buttons: 0 })); } catch (_) {}
    try { el.dispatchEvent(new MouseEvent('click', { ...opts, buttons: 0 })); } catch (_) {}
  }

  function describeActivationTarget(el, name) {
    if (!el) return null;
    const r = el.getBoundingClientRect();
    return {
      name,
      tag: el.tagName || '',
      role: el.getAttribute?.('role') || '',
      className: typeof el.className === 'string' ? el.className.slice(0, 220) : '',
      text: normalize(el.innerText || el.textContent || '').slice(0, 180),
      rect: {
        left: Math.round(r.left),
        top: Math.round(r.top),
        width: Math.round(r.width),
        height: Math.round(r.height)
      }
    };
  }

  function calendarActivationTargets(card) {
    const result = [];
    const seen = new Set();

    const add = (name, el) => {
      if (!el || !(el instanceof Element) || seen.has(el) || !isRendered(el)) return;
      seen.add(el);
      result.push({ name, el });
    };

    add('card', card);

    // Fluent UI event cards wrap the actual body in a stable ui-card__body
    // element. Recurring meetings can react to the body while a synthetic
    // click on the role=group wrapper itself is ignored.
    add('card-body', card.querySelector('[class*="ui-card__body"]'));

    const r = card.getBoundingClientRect();
    if (r.width > 2 && r.height > 2) {
      const points = [
        ['center-hit', r.left + r.width / 2, r.top + r.height / 2],
        ['title-hit', r.left + Math.min(40, r.width * 0.25), r.top + Math.min(24, r.height * 0.3)]
      ];
      for (const [name, x, y] of points) {
        const hit = document.elementFromPoint(
          Math.max(0, Math.min(innerWidth - 1, x)),
          Math.max(0, Math.min(innerHeight - 1, y))
        );
        if (hit && (hit === card || card.contains(hit))) add(name, hit);
      }
    }

    // In the captured Teams DOM the event card itself is inside
    // <div class="fui-Primitive" draggable="true">.
    const wrapper = card.parentElement;
    if (wrapper?.getAttribute?.('draggable') === 'true') add('draggable-wrapper', wrapper);

    return result;
  }

  function dispatchKeyboardActivation(el, key) {
    const code = key === ' ' ? 'Space' : 'Enter';
    const keyCode = key === ' ' ? 32 : 13;
    try { el.focus({ preventScroll: true }); } catch (_) {}
    try {
      el.dispatchEvent(new KeyboardEvent('keydown', {
        key,
        code,
        keyCode,
        which: keyCode,
        bubbles: true,
        cancelable: true,
        composed: true
      }));
      el.dispatchEvent(new KeyboardEvent('keyup', {
        key,
        code,
        keyCode,
        which: keyCode,
        bubbles: true,
        cancelable: true,
        composed: true
      }));
    } catch (_) {}
  }

  function dispatchDoubleClick(el) {
    const r = el.getBoundingClientRect();
    const opts = {
      bubbles: true,
      cancelable: true,
      composed: true,
      view: window,
      clientX: r.left + Math.max(1, r.width / 2),
      clientY: r.top + Math.max(1, r.height / 2),
      button: 0,
      buttons: 0,
      detail: 2
    };
    try { el.dispatchEvent(new MouseEvent('dblclick', opts)); } catch (_) {}
  }

  async function openCalendarMeeting(meeting) {
    const attempts = [];
    let lastTarget = null;

    async function waitForOpenedState(timeoutMs) {
      const started = Date.now();
      while (Date.now() - started < timeoutMs) {
        const chatAction = findChatWithParticipants(meeting);
        if (chatAction) return { mode: 'popup', chatAction };
        if (meetingDetailsViewMatches(meeting)) return { mode: 'details', chatAction: null };
        await sleep(120);
      }
      return null;
    }

    async function prepareFreshCard() {
      if (!visibleCalendarRange()) {
        const nav = await navigateToCalendar();
        if (!nav.ok) return null;
      }

      if (meeting?.dateStamp) {
        const ensured = await ensureCalendarDate(meeting.dateStamp);
        if (!ensured.ok) return null;
      }

      await sleep(300);
      const card = findMeetingElement(meeting);
      if (!card) return null;

      try { card.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
      await sleep(180);
      return findMeetingElement(meeting) || card;
    }

    async function restoreForNextAttempt() {
      try {
        if (!visibleCalendarRange()) await navigateToCalendar();
        if (meeting?.dateStamp) await ensureCalendarDate(meeting.dateStamp);
        await sleep(300);
      } catch (_) {}
    }

    // Try several DOM targets. The real Teams markup uses role=group inside a
    // draggable Fluent wrapper, and recurring events do not always react to
    // synthetic activation on the outer role=group element.
    const strategies = [
      { type: 'click' },
      { type: 'pointer' },
      { type: 'keyboard-enter', cardOnly: true },
      { type: 'keyboard-space', cardOnly: true },
      { type: 'double-click' }
    ];

    for (const strategy of strategies) {
      const card = await prepareFreshCard();
      if (!card) {
        return {
          ok: false,
          error: 'Meeting card not found in current Calendar view.',
          attempts,
          diagnostic: popupDiagnostic(meeting)
        };
      }

      lastTarget = card;
      const targets = strategy.cardOnly
        ? [{ name: 'card', el: card }]
        : calendarActivationTargets(card);

      for (const target of targets) {
        try { target.el.scrollIntoView({ block: 'nearest', inline: 'nearest' }); } catch (_) {}
        await sleep(80);

        if (strategy.type === 'click') {
          try { target.el.click(); } catch (_) { dispatchPointerSequence(target.el); }
        } else if (strategy.type === 'pointer') {
          dispatchPointerSequence(target.el);
        } else if (strategy.type === 'keyboard-enter') {
          dispatchKeyboardActivation(target.el, 'Enter');
        } else if (strategy.type === 'keyboard-space') {
          dispatchKeyboardActivation(target.el, ' ');
        } else if (strategy.type === 'double-click') {
          dispatchDoubleClick(target.el);
        }

        attempts.push({
          strategy: strategy.type,
          target: describeActivationTarget(target.el, target.name)
        });

        const openedState = await waitForOpenedState(1700);
        if (openedState) {
          return {
            ok: true,
            mode: openedState.mode,
            url: location.href,
            attempts,
            activation: attempts[attempts.length - 1],
            target: {
              tag: card.tagName,
              role: card.getAttribute('role') || '',
              dataTestId: card.getAttribute('data-testid') || '',
              elementId: card.id || '',
              ariaLabel: card.getAttribute('aria-label') || ''
            }
          };
        }

        // If Teams navigated somewhere that is neither Calendar nor the
        // requested meeting, restore Calendar before trying another target.
        if (!visibleCalendarRange()) {
          await restoreForNextAttempt();
          break;
        }
      }

      await restoreForNextAttempt();
    }

    return {
      ok: false,
      error: 'Meeting view did not become ready after activating Calendar card.',
      attempts,
      diagnostic: popupDiagnostic(meeting),
      target: lastTarget ? {
        tag: lastTarget.tagName,
        role: lastTarget.getAttribute('role') || '',
        dataTestId: lastTarget.getAttribute('data-testid') || '',
        elementId: lastTarget.id || '',
        ariaLabel: lastTarget.getAttribute('aria-label') || ''
      } : {}
    };
  }

  async function openMeetingChatFromCalendar(meeting) {
    const opened = await openCalendarMeeting(meeting);
    if (!opened.ok) return opened;

    if (opened.mode === 'details') {
      return {
        ok: true,
        mode: 'details',
        url: location.href,
        attempts: opened.attempts || [],
        target: opened.target || {},
        details: findMeetingDetailsAssets(meeting)
      };
    }

    const chatButton = findChatWithParticipants(meeting);
    if (!chatButton) {
      return {
        ok: false,
        error: '"Chat with participants" action disappeared after meeting popup opened.',
        attempts: opened.attempts || [],
        diagnostic: popupDiagnostic(meeting)
      };
    }

    const popupText = (() => {
      let node = chatButton;
      for (let i = 0; i < 8 && node; i++, node = node.parentElement) {
        if (node.matches?.('[data-tid="peek-body-container"],[role="dialog"]')) {
          return normalize(node.innerText || node.textContent || '').slice(0, 1200);
        }
      }
      return '';
    })();

    try { chatButton.scrollIntoView({ block: 'nearest', inline: 'nearest' }); } catch (_) {}
    try { chatButton.focus({ preventScroll: true }); } catch (_) {}
    try { chatButton.click(); } catch (_) { dispatchPointerSequence(chatButton); }

    await sleep(900);

    return {
      ok: true,
      url: location.href,
      attempts: opened.attempts || [],
      popup: {
        text: popupText,
        actionTag: chatButton.tagName,
        actionRole: chatButton.getAttribute('role') || '',
        actionAriaLabel: chatButton.getAttribute('aria-label') || '',
        actionText: interactiveText(chatButton)
      }
    };
  }

  function visibleCalendarRange() {
    const columns = getVisibleDayColumns();
    if (!columns.length) return null;
    const dates = columns
      .map(x => ({ ...x, utc: datePartsToUtc(x) }))
      .filter(x => Number.isFinite(x.utc))
      .sort((a, b) => a.utc - b.utc);
    if (!dates.length) return null;
    return { min: dates[0], max: dates[dates.length - 1], columns: dates };
  }

  function calendarToolbarNavigationByGeometry(kind, els) {
    if (kind !== 'previous' && kind !== 'next') return null;

    const today = els.find(el => {
      const text = normalize(`${accessibleText(el)} ${el.getAttribute('aria-label') || ''} ${el.getAttribute('title') || ''}`);
      return /^(?:today|сегодня)(?:\s|$)/i.test(text);
    });
    if (!today) return null;

    const tr = today.getBoundingClientRect();
    const ty = tr.top + tr.height / 2;

    // In Teams Web the two date-range arrows are the first two compact buttons
    // immediately to the right of Today. Their accessible names can be empty.
    const nearby = els
      .filter(el => el !== today)
      .map(el => ({ el, r: el.getBoundingClientRect() }))
      .filter(x => {
        const cy = x.r.top + x.r.height / 2;
        return (
          x.r.left >= tr.right - 2 &&
          x.r.left <= tr.right + 190 &&
          Math.abs(cy - ty) <= 26 &&
          x.r.width > 12 && x.r.width <= 64 &&
          x.r.height > 12 && x.r.height <= 64
        );
      })
      .sort((a, b) => a.r.left - b.r.left);

    if (nearby.length < 2) return null;
    return kind === 'previous' ? nearby[0].el : nearby[1].el;
  }

  function calendarNavDiagnostic(els = null) {
    const controls = els || Array.from(
      document.querySelectorAll('button,[role="button"],a[href],[tabindex]')
    ).filter(isRendered);

    return controls.slice(0, 80).map(el => {
      const r = el.getBoundingClientRect();
      return {
        text: accessibleText(el),
        aria: el.getAttribute('aria-label') || '',
        title: el.getAttribute('title') || '',
        tid: el.getAttribute('data-tid') || el.getAttribute('data-testid') || '',
        rect: {
          left: Math.round(r.left),
          top: Math.round(r.top),
          width: Math.round(r.width),
          height: Math.round(r.height)
        }
      };
    });
  }

  function findCalendarNavControl(kind) {
    const els = Array.from(
      document.querySelectorAll('button,[role="button"],a[href],[tabindex]')
    ).filter(isRendered);

    const patterns = kind === 'calendar'
      ? [/^calendar(?:\s*\([^)]*\))?$/i, /calendar/i]
      : kind === 'previous'
        ? [/previous/i, /prev/i, /предыдущ/i, /назад/i]
        : [/next/i, /следующ/i, /впер[её]д/i, /далее/i];

    const semantic = els.find(el => {
      const text = accessibleText(el);
      const tid = normalize(`${el.getAttribute('data-tid') || ''} ${el.getAttribute('data-testid') || ''} ${el.getAttribute('aria-label') || ''} ${el.getAttribute('title') || ''}`);
      return patterns.some(re => re.test(text) || re.test(tid));
    });
    if (semantic) return semantic;

    return calendarToolbarNavigationByGeometry(kind, els);
  }

  async function navigateToCalendar() {
    if (visibleCalendarRange()) return { ok: true, alreadyThere: true };

    const control = findCalendarNavControl('calendar');
    if (!control) return { ok: false, error: 'Calendar navigation control not found.' };

    try { control.click(); } catch (_) {
      control.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window, button: 0 }));
    }

    for (let i = 0; i < 30; i++) {
      await sleep(300);
      if (visibleCalendarRange()) return { ok: true, alreadyThere: false };
    }
    return { ok: false, error: 'Calendar did not become ready.' };
  }

  async function ensureCalendarDate(dateStamp) {
    const target = parseDateStampParts(dateStamp);
    if (!target) return { ok: false, error: 'Meeting date is unavailable.' };

    const nav = await navigateToCalendar();
    if (!nav.ok) return nav;

    const targetUtc = datePartsToUtc(target);

    for (let attempt = 0; attempt < 24; attempt++) {
      const range = visibleCalendarRange();
      if (!range) {
        await sleep(350);
        continue;
      }

      if (targetUtc >= range.min.utc && targetUtc <= range.max.utc) {
        return {
          ok: true,
          range: {
            from: toDateStamp(range.min),
            to: toDateStamp(range.max)
          },
          attempts: attempt
        };
      }

      const direction = targetUtc < range.min.utc ? 'previous' : 'next';
      const button = findCalendarNavControl(direction);
      if (!button) {
        return {
          ok: false,
          error: `Calendar ${direction} control not found.`,
          currentRange: {
            from: toDateStamp(range.min),
            to: toDateStamp(range.max)
          },
          navDiagnostic: calendarNavDiagnostic()
        };
      }

      const beforeKey = `${toDateStamp(range.min)}|${toDateStamp(range.max)}`;
      try { button.click(); } catch (_) {
        dispatchPointerSequence(button);
      }

      // Do not rely on a fixed sleep. Teams updates the Calendar grid
      // asynchronously and the old week can remain in DOM for a while.
      const waitStarted = Date.now();
      while (Date.now() - waitStarted < 3500) {
        await sleep(140);
        const nextRange = visibleCalendarRange();
        if (!nextRange) continue;
        const nextKey = `${toDateStamp(nextRange.min)}|${toDateStamp(nextRange.max)}`;
        if (nextKey !== beforeKey) break;
      }
    }

    const finalRange = visibleCalendarRange();
    return {
      ok: false,
      error: 'Unable to navigate Calendar to the meeting week.',
      currentRange: finalRange ? {
        from: toDateStamp(finalRange.min),
        to: toDateStamp(finalRange.max)
      } : null,
      navDiagnostic: calendarNavDiagnostic()
    };
  }

  function recapCardContext(card) {
    const message = card.closest('[data-tid="control-message-renderer"]');
    const headingId = message?.getAttribute('aria-labelledby') || '';
    const heading = headingId ? document.getElementById(headingId) : null;
    const headingText = normalize(
      heading?.innerText ||
      heading?.textContent ||
      message?.innerText ||
      message?.textContent ||
      card.innerText ||
      card.textContent ||
      ''
    );

    return { message, heading, headingText };
  }

  function findMeetingRecap(meeting) {
    actionMap = new Map();

    const cards = Array.from(document.querySelectorAll(
      '[data-testid="meeting-recap-object"], [data-testid="meeting-recap-chiclet"]'
    )).filter(isRendered);

    const expectedTitle = normalizedComparable(meeting?.title || '');
    const expectedDate = meeting?.dateStamp || '';
    const matches = [];

    for (const card of cards) {
      const { headingText } = recapCardContext(card);
      const cardText = normalize(card.innerText || card.textContent || '');
      const fullContext = normalize(`${headingText} ${cardText}`);
      const dateStamp = toDateStamp(
        parseFlexibleDate(fullContext, currentCalendarMonthYear()?.year)
      );

      const titleNode = card.querySelector('[data-tid="meeting-title"]');
      const cardTitle = normalize(
        titleNode?.innerText ||
        titleNode?.textContent ||
        card.querySelector('[data-testid="meeting-recap-chiclet-top-container"]')?.innerText ||
        cardText
      );
      const comparableCard = normalizedComparable(cardTitle);
      const comparableHeading = normalizedComparable(headingText);
      const comparablePage = normalizedComparable(document.title);

      const titleMatches =
        !!expectedTitle &&
        (
          comparableCard.includes(expectedTitle) ||
          comparableHeading.includes(expectedTitle) ||
          comparablePage.includes(expectedTitle)
        );
      const dateMatches = !!expectedDate && dateStamp === expectedDate;

      if (!titleMatches || !dateMatches) continue;

      const recordingButton = card.querySelector(
        '[data-testid="meeting-recap-chiclet-recording-image"]'
      );
      const transcriptButton = Array.from(card.querySelectorAll('button')).find(btn =>
        normalize(btn.getAttribute('aria-label') || btn.innerText || '') === 'Transcript'
      );
      const viewRecapButton = card.querySelector(
        '[data-testid="view-meeting-recap-button"], [data-testid="meeting-recap-chiclet-view-recap-button"]'
      );

      const actions = {};
      for (const [kind, el] of [
        ['recording', recordingButton],
        ['transcript', transcriptButton],
        ['recap', viewRecapButton]
      ]) {
        if (!el) continue;
        const actionId = `exact-${kind}-${hash(`${meeting.id}|${kind}|${headingText}`)}`;
        actionMap.set(actionId, el);
        actions[`${kind}ActionId`] = actionId;
      }

      matches.push({
        kind: card.getAttribute('data-testid') || '',
        dateStamp,
        headingText: headingText.slice(0, 1200),
        cardTitle: cleanMeetingTitle(headingText),
        actions,
        hasRecording: !!recordingButton,
        hasTranscript: !!transcriptButton,
        hasViewRecap: !!viewRecapButton
      });
    }

    return {
      ok: true,
      match: matches[0] || null,
      matches,
      pageTitle: document.title,
      url: location.href
    };
  }

  function findRecapCards() {
    actionMap = new Map();
    const result = [];
    let seq = 0;
    const buttons = Array.from(document.querySelectorAll('button,[role="button"],a[href]')).filter(isRendered);

    for (const button of buttons) {
      const label = accessibleText(button);
      if (!/view\s+recap|recap/i.test(label)) continue;

      let card = button;
      for (let depth = 0; depth < 7 && card?.parentElement; depth++) {
        const parent = card.parentElement;
        const text = normalize(parent.innerText || parent.textContent || '');
        if (/\bTranscript\b/i.test(text) && timeRegex.test(text) && text.length < 2200) {
          card = parent;
          break;
        }
        card = parent;
      }

      const cardText = normalize(card?.innerText || card?.textContent || label);
      const date = parseFlexibleDate(cardText, currentCalendarMonthYear()?.year);
      const dateStamp = toDateStamp(date);
      const id = `recap-${++seq}-${hash(`${dateStamp}|${cardText}`)}`;
      actionMap.set(id, button);

      result.push({
        id,
        kind: 'recap-card',
        label,
        dateStamp,
        text: cardText.slice(0, 900),
        href: button.href || button.closest('a[href]')?.href || ''
      });
    }
    return result;
  }

  function classifyAction(label, href) {
    const text = `${label || ''} ${href || ''}`.toLowerCase();
    if (/recap|meeting recap|summary/.test(text)) return 'recap';
    if (/recording|recorded|watch recording|meeting recording|stream\.aspx|\.mp4(?:\?|$)/.test(text)) return 'recording';
    if (/meeting chat|\bchat\b/.test(text)) return 'chat';
    if (/transcript/.test(text)) return 'transcript';
    return '';
  }

  function normalizedComparable(value) {
    return normalize(value)
      .toLowerCase()
      .replace(/[“”"'.,;:()[\]{}]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function elementBelongsToMeeting(el, meeting) {
    if (!el || !meeting?.title) return false;
    const expected = normalizedComparable(meeting.title);
    if (!expected) return false;

    let node = el;
    for (let depth = 0; node && depth < 8; depth++, node = node.parentElement) {
      if (!(node instanceof Element)) break;
      const text = normalize(node.innerText || node.textContent || '');
      if (!text || text.length > 6500) continue;
      const comparable = normalizedComparable(text);

      if (comparable.includes(expected)) {
        return true;
      }

      // Do not walk all the way into the global Teams shell. Otherwise a
      // left-navigation "Chat" button eventually inherits the whole page text
      // and incorrectly looks related to the meeting.
      if (
        node.matches('body,main,[role="main"],nav,[role="navigation"]') ||
        text.length > 4500
      ) {
        break;
      }
    }
    return false;
  }

  function classifyMeetingScopedAction(label, href) {
    const value = normalize(label);
    const lower = `${value} ${href || ''}`.toLowerCase();

    if (/view\s+recap|open\s+recap|meeting\s+recap/i.test(value)) return 'recap';
    if (/watch\s+recording|open\s+recording|meeting\s+recording|recorded/i.test(value) || /stream\.aspx|\.mp4(?:\?|$)/i.test(href || '')) return 'recording';
    if (/^(?:chat|meeting chat|open chat|chat with participants)$/i.test(value)) return 'chat';
    if (/^transcript$/i.test(value)) return 'transcript';

    // A direct SharePoint/Stream link inside the meeting details is also safe.
    if (/sharepoint\.com|stream\.aspx/i.test(lower)) return 'recording';
    return '';
  }

  function findMeetingScopedActions(meeting) {
    actionMap = new Map();
    const result = [];
    let seq = 0;
    const els = Array.from(document.querySelectorAll(
      'a[href],button,[role="button"],[role="link"],[tabindex]'
    ));

    for (const el of els) {
      if (!isRendered(el)) continue;
      if (!elementBelongsToMeeting(el, meeting)) continue;

      const label = accessibleText(el);
      const href = el.href || el.closest('a[href]')?.href || '';
      const kind = classifyMeetingScopedAction(label, href);
      if (!kind) continue;

      const id = `meeting-${++seq}-${hash(`${kind}|${label}|${href}`)}`;
      actionMap.set(id, el);
      result.push({
        id,
        kind,
        label,
        href,
        contextTitle: meeting.title
      });
    }

    return result.slice(0, 40);
  }

  function getCurrentConversationTitle() {
    const candidates = Array.from(document.querySelectorAll(
      'h1,h2,h3,[role="heading"]'
    ))
      .filter(isRendered)
      .map(el => {
        const r = el.getBoundingClientRect();
        return {
          text: normalize(el.innerText || el.textContent || ''),
          top: r.top,
          left: r.left,
          width: r.width
        };
      })
      .filter(x =>
        x.text &&
        x.text.length >= 3 &&
        x.text.length <= 220 &&
        x.top >= 80 &&
        x.top <= 320 &&
        !/^(?:chat|recap|attendance|shared|calendar)$/i.test(x.text)
      )
      .sort((a, b) => a.top - b.top || a.left - b.left);

    return candidates[0]?.text || '';
  }

  function getPageContext() {
    return {
      url: location.href,
      pageKind: pageKind(),
      conversationTitle: getCurrentConversationTitle(),
      documentTitle: document.title
    };
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
    return result.slice(0, 100);
  }

  async function triggerAction(id) {
    if (!actionMap.has(id)) findActions();
    const target = actionMap.get(id);
    if (!target || !target.isConnected) return { ok: false, error: 'Action element not found.' };

    const before = {
      url: location.href,
      title: document.title,
      text: interactiveText(target),
      tag: target.tagName,
      role: target.getAttribute?.('role') || '',
      href: target.href || target.closest?.('a[href]')?.href || ''
    };

    try { target.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(100);
    try { target.focus({ preventScroll: true }); } catch (_) {}

    let method = 'click';
    try {
      target.click();
    } catch (_) {
      method = 'pointer';
      dispatchPointerSequence(target);
    }

    await sleep(450);

    return {
      ok: true,
      method,
      before,
      after: {
        url: location.href,
        title: document.title
      }
    };
  }

  function findRecordingLinks() {
    const links = [];
    const seen = new Set();
    for (const a of document.querySelectorAll('a[href]')) {
      const href = a.href || '';
      const label = accessibleText(a);
      const lower = `${href} ${label}`.toLowerCase();
      if (!/sharepoint\.com|stream\.aspx|meeting recording|recording|\.mp4(?:\?|$)/.test(lower)) continue;
      if (!/^https?:/i.test(href) || seen.has(href)) continue;
      seen.add(href);
      links.push({ href, label, kind: classifyAction(label, href) || 'recording' });
    }
    return links.slice(0, 30);
  }

  function pageKind() {
    const url = location.href.toLowerCase();
    const body = normalize(document.body?.innerText || '').slice(0, 5000).toLowerCase();
    if (/calendar/.test(url) || /\bcalendar\b/.test(body)) return 'calendar';
    if (/sharepoint\.com/.test(url) || /stream\.aspx/.test(url) || /meeting recording/.test(body)) return 'recording';
    if (/chat/.test(url) || /meeting chat/.test(body)) return 'chat';
    return 'teams';
  }

  function buildDiagnostic() {
    const meetings = scanCalendarMeetings();
    const actions = findActions();
    const links = findRecordingLinks();
    const recapCards = findRecapCards();
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

    meetings.forEach((m, i) => lines.push(
      `[M${i + 1}] id=${m.id} date=${m.dateStamp || '-'} time=${m.startTime || '-'} score=${m.score} rect=${JSON.stringify(m.rect)} dom=${JSON.stringify(m.dom || {})} title=${m.title} label=${m.label}`
    ));

    lines.push('', `Candidates: ${(lastCalendarScanDebug.candidates || []).length}`);
    (lastCalendarScanDebug.candidates || []).slice(0, 50).forEach((x, i) =>
      lines.push(`[CANDIDATE ${i + 1}] ${JSON.stringify(x)}`)
    );

    lines.push('', `Rejected controls/slots: ${(lastCalendarScanDebug.rejected || []).length}`);
    (lastCalendarScanDebug.rejected || []).slice(0, 50).forEach((x, i) =>
      lines.push(`[REJECTED ${i + 1}] ${JSON.stringify(x)}`)
    );

    lines.push('', `Actions found: ${actions.length}`);
    actions.forEach((a, i) => lines.push(`[A${i + 1}] kind=${a.kind} label=${a.label} href=${a.href || '-'}`));
    lines.push('', `Recap cards found: ${recapCards.length}`);
    recapCards.forEach((x, i) => lines.push(`[RC${i + 1}] date=${x.dateStamp || '-'} label=${x.label} text=${x.text}`));
    lines.push('', `Recording links found: ${links.length}`);
    links.forEach((l, i) => lines.push(`[R${i + 1}] kind=${l.kind} label=${l.label} href=${l.href}`));
    lines.push('', 'Visible page text sample:', normalize(document.body?.innerText || '').slice(0, 14000));
    return lines.join('\n');
  }

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    const type = message?.type;

    if (type === 'CALENDAR_SCAN') {
      sendResponse({ ok: true, meetings: scanCalendarMeetings(), pageKind: pageKind(), url: location.href });
      return;
    }
    if (type === 'CALENDAR_OPEN_MEETING_CHAT') {
      openMeetingChatFromCalendar(message.meeting).then(sendResponse);
      return true;
    }
    if (type === 'PAGE_FIND_MEETING_DETAILS_ASSETS') {
      sendResponse(findMeetingDetailsAssets(message.meeting));
      return;
    }
    if (type === 'PAGE_OPEN_CALENDAR') {
      openCalendarNavigation().then(sendResponse);
      return true;
    }
    if (type === 'PAGE_FIND_MEETING_RECAP_EXACT') {
      sendResponse(findMeetingRecap(message.meeting));
      return;
    }
    if (type === 'CALENDAR_OPEN_MEETING') {
      openCalendarMeeting(message.meeting).then(sendResponse);
      return true;
    }
    if (type === 'CALENDAR_ENSURE_DATE') {
      ensureCalendarDate(message.dateStamp).then(sendResponse);
      return true;
    }
    if (type === 'PAGE_FIND_RECAP_CARDS') {
      sendResponse({ ok: true, cards: findRecapCards(), url: location.href, pageKind: pageKind() });
      return;
    }
    if (type === 'PAGE_FIND_MEETING_ACTIONS') {
      sendResponse({
        ok: true,
        actions: findMeetingScopedActions(message.meeting),
        context: getPageContext()
      });
      return;
    }
    if (type === 'PAGE_GET_CONTEXT') {
      sendResponse({ ok: true, context: getPageContext() });
      return;
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
