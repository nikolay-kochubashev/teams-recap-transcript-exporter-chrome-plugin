(() => {
  if (window.__teamsTranscriptCalendarLoaded) return;
  window.__teamsTranscriptCalendarLoaded = true;

  const VERSION = '2.0.8';
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
    if (!/^\\d{8}$/.test(String(stamp || ''))) return null;
    return {
      year: Number(stamp.slice(0, 4)),
      month: Number(stamp.slice(4, 6)),
      day: Number(stamp.slice(6, 8))
    };
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
    const rejected = [];
    const debugCandidates = [];

    for (const el of pool) {
      if (!isRendered(el)) continue;
      const text = accessibleText(el);
      const score = meetingCandidateScore(el, text);

      const dom = {
        tag: el.tagName,
        role: el.getAttribute('role') || '',
        dataTid: el.getAttribute('data-tid') || '',
        dataTestId: el.getAttribute('data-testid') || '',
        ariaLabel: (el.getAttribute('aria-label') || '').slice(0, 500),
        className: typeof el.className === 'string' ? el.className.slice(0, 300) : ''
      };

      if (score < 55) {
        if ((isCalendarAggregateSlot(text) || isNonMeetingControl(text)) && rejected.length < 80) {
          rejected.push({ text: text.slice(0, 350), score, dom });
        }
        continue;
      }

      const r = el.getBoundingClientRect();
      raw.push({ el, text, score, r, dom });
      if (debugCandidates.length < 80) debugCandidates.push({ text: text.slice(0, 500), score, dom });
    }

    const dayColumns = getVisibleDayColumns();
    const calendarContext = currentCalendarMonthYear();
    const byId = new Map();

    for (const item of raw) {
      const label = item.text;
      const parentText = normalize(item.el.parentElement?.innerText || '').slice(0, 1000);
      const explicitDate = parseFlexibleDate(`${label} ${parentText}`, calendarContext?.year);
      const inferredDate = explicitDate || inferDateFromColumn(item.r, dayColumns);
      const startTime = parseStartTime(`${label} ${parentText}`);
      const dateStamp = toDateStamp(inferredDate);
      const startTimeText = startTime
        ? `${String(startTime.hour).padStart(2, '0')}:${String(startTime.minute).padStart(2, '0')}`
        : '';
      const title = cleanMeetingTitle(label);
      const id = hash(`${normalize(label)}|${dateStamp}|${startTimeText}`);
      const href = item.el.href || item.el.closest('a[href]')?.href || '';

      const meeting = {
        id,
        label,
        title,
        dateStamp,
        startTime: startTimeText,
        sortKey: toSortKey(inferredDate, startTime, byId.size),
        href,
        score: item.score,
        dom: item.dom,
        rect: {
          top: Math.round(item.r.top),
          left: Math.round(item.r.left),
          width: Math.round(item.r.width),
          height: Math.round(item.r.height)
        }
      };

      const existing = byId.get(id);
      if (!existing || meeting.score > existing.score) byId.set(id, meeting);
    }

    const meetings = Array.from(byId.values())
      .sort((a, b) => a.sortKey - b.sortKey || a.title.localeCompare(b.title, 'ru'))
      .slice(0, 120);

    lastCalendarScanDebug = {
      rejected,
      candidates: debugCandidates,
      acceptedCount: meetings.length
    };
    return meetings;
  }

  function findMeetingElement(meeting) {
    if (!meeting) return null;
    const pool = Array.from(document.querySelectorAll([
      '[data-tid*="calendar"]',
      '[data-tid*="event"]',
      '[data-testid*="calendar"]',
      '[data-testid*="event"]',
      '[role="gridcell"]',
      '[role="button"]',
      'button',
      'a[href]'
    ].join(','))).filter(isRendered);

    const exact = pool.filter(el => accessibleText(el) === meeting.label);
    if (exact.length) {
      return exact.sort((a, b) => meetingCandidateScore(b, accessibleText(b)) - meetingCandidateScore(a, accessibleText(a)))[0];
    }

    const title = normalize(meeting.title).toLowerCase();
    const fallback = pool
      .map(el => ({ el, text: accessibleText(el), r: el.getBoundingClientRect() }))
      .filter(x => x.text && normalize(x.text).toLowerCase().startsWith(title))
      .sort((a, b) => {
        const ad = Math.abs(a.r.left - (meeting.rect?.left || a.r.left)) + Math.abs(a.r.top - (meeting.rect?.top || a.r.top));
        const bd = Math.abs(b.r.left - (meeting.rect?.left || b.r.left)) + Math.abs(b.r.top - (meeting.rect?.top || b.r.top));
        return ad - bd;
      });
    return fallback[0]?.el || null;
  }

  async function openCalendarMeeting(meeting) {
    const el = findMeetingElement(meeting);
    if (!el) {
      return { ok: false, error: 'Meeting element not found in current calendar view.' };
    }

    const target = el.matches('button,a,[role="button"],[role="link"],[tabindex]')
      ? el
      : (el.querySelector('button,a,[role="button"],[role="link"],[tabindex]')
        || el.closest('button,a,[role="button"],[role="link"],[tabindex]')
        || el);

    try { target.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(120);
    try { target.focus({ preventScroll: true }); } catch (_) {}

    try {
      target.click();
    } catch (_) {
      target.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window, button: 0 }));
    }

    await sleep(1000);
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

  function findCalendarNavControl(kind) {
    const els = Array.from(document.querySelectorAll('button,[role="button"],a[href],[tabindex]')).filter(isRendered);
    const patterns = kind === 'calendar'
      ? [/^calendar(?:\s*\([^)]*\))?$/i, /calendar/i]
      : kind === 'previous'
        ? [/previous\s*(?:week|period|date)?/i, /предыдущ/i, /назад/i]
        : [/next\s*(?:week|period|date)?/i, /следующ/i, /впер[её]д/i, /далее/i];

    return els.find(el => {
      const text = accessibleText(el);
      const tid = normalize(`${el.getAttribute('data-tid') || ''} ${el.getAttribute('data-testid') || ''} ${el.getAttribute('aria-label') || ''} ${el.getAttribute('title') || ''}`);
      return patterns.some(re => re.test(text) || re.test(tid));
    }) || null;
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
        return { ok: false, error: `Calendar ${direction} control not found.` };
      }

      try { button.click(); } catch (_) {
        button.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window, button: 0 }));
      }
      await sleep(650);
    }

    return { ok: false, error: 'Unable to navigate Calendar to the meeting week.' };
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
    try { target.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(80);
    target.click();
    return { ok: true, url: location.href };
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
    const monthFirstTail = new RegExp(',?\\s*(?:(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\\s*)?(?:' + monthPattern + ')\\s+\\d{1,2},?\\s+20\\d{2}\\b.*

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
    const rejected = [];
    const debugCandidates = [];

    for (const el of pool) {
      if (!isRendered(el)) continue;
      const text = accessibleText(el);
      const score = meetingCandidateScore(el, text);

      const dom = {
        tag: el.tagName,
        role: el.getAttribute('role') || '',
        dataTid: el.getAttribute('data-tid') || '',
        dataTestId: el.getAttribute('data-testid') || '',
        ariaLabel: (el.getAttribute('aria-label') || '').slice(0, 500),
        className: typeof el.className === 'string' ? el.className.slice(0, 300) : ''
      };

      if (score < 55) {
        if ((isCalendarAggregateSlot(text) || isNonMeetingControl(text)) && rejected.length < 80) {
          rejected.push({ text: text.slice(0, 350), score, dom });
        }
        continue;
      }

      const r = el.getBoundingClientRect();
      raw.push({ el, text, score, r, dom });
      if (debugCandidates.length < 80) debugCandidates.push({ text: text.slice(0, 500), score, dom });
    }

    const dayColumns = getVisibleDayColumns();
    const calendarContext = currentCalendarMonthYear();
    const byId = new Map();

    for (const item of raw) {
      const label = item.text;
      const parentText = normalize(item.el.parentElement?.innerText || '').slice(0, 1000);
      const explicitDate = parseFlexibleDate(`${label} ${parentText}`, calendarContext?.year);
      const inferredDate = explicitDate || inferDateFromColumn(item.r, dayColumns);
      const startTime = parseStartTime(`${label} ${parentText}`);
      const dateStamp = toDateStamp(inferredDate);
      const startTimeText = startTime
        ? `${String(startTime.hour).padStart(2, '0')}:${String(startTime.minute).padStart(2, '0')}`
        : '';
      const title = cleanMeetingTitle(label);
      const id = hash(`${normalize(label)}|${dateStamp}|${startTimeText}`);
      const href = item.el.href || item.el.closest('a[href]')?.href || '';

      const meeting = {
        id,
        label,
        title,
        dateStamp,
        startTime: startTimeText,
        sortKey: toSortKey(inferredDate, startTime, byId.size),
        href,
        score: item.score,
        dom: item.dom,
        rect: {
          top: Math.round(item.r.top),
          left: Math.round(item.r.left),
          width: Math.round(item.r.width),
          height: Math.round(item.r.height)
        }
      };

      const existing = byId.get(id);
      if (!existing || meeting.score > existing.score) byId.set(id, meeting);
    }

    const meetings = Array.from(byId.values())
      .sort((a, b) => a.sortKey - b.sortKey || a.title.localeCompare(b.title, 'ru'))
      .slice(0, 120);

    lastCalendarScanDebug = {
      rejected,
      candidates: debugCandidates,
      acceptedCount: meetings.length
    };
    return meetings;
  }

  function findMeetingElement(meeting) {
    if (!meeting) return null;
    const pool = Array.from(document.querySelectorAll([
      '[data-tid*="calendar"]',
      '[data-tid*="event"]',
      '[data-testid*="calendar"]',
      '[data-testid*="event"]',
      '[role="gridcell"]',
      '[role="button"]',
      'button',
      'a[href]'
    ].join(','))).filter(isRendered);

    const exact = pool.filter(el => accessibleText(el) === meeting.label);
    if (exact.length) {
      return exact.sort((a, b) => meetingCandidateScore(b, accessibleText(b)) - meetingCandidateScore(a, accessibleText(a)))[0];
    }

    const title = normalize(meeting.title).toLowerCase();
    const fallback = pool
      .map(el => ({ el, text: accessibleText(el), r: el.getBoundingClientRect() }))
      .filter(x => x.text && normalize(x.text).toLowerCase().startsWith(title))
      .sort((a, b) => {
        const ad = Math.abs(a.r.left - (meeting.rect?.left || a.r.left)) + Math.abs(a.r.top - (meeting.rect?.top || a.r.top));
        const bd = Math.abs(b.r.left - (meeting.rect?.left || b.r.left)) + Math.abs(b.r.top - (meeting.rect?.top || b.r.top));
        return ad - bd;
      });
    return fallback[0]?.el || null;
  }

  async function openCalendarMeeting(meeting) {
    const el = findMeetingElement(meeting);
    if (!el) {
      return { ok: false, error: 'Meeting element not found in current calendar view.' };
    }

    const target = el.matches('button,a,[role="button"],[role="link"],[tabindex]')
      ? el
      : (el.querySelector('button,a,[role="button"],[role="link"],[tabindex]')
        || el.closest('button,a,[role="button"],[role="link"],[tabindex]')
        || el);

    try { target.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(120);
    try { target.focus({ preventScroll: true }); } catch (_) {}

    try {
      target.click();
    } catch (_) {
      target.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window, button: 0 }));
    }

    await sleep(1000);
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

  function findCalendarNavControl(kind) {
    const els = Array.from(document.querySelectorAll('button,[role="button"],a[href],[tabindex]')).filter(isRendered);
    const patterns = kind === 'calendar'
      ? [/^calendar(?:\s*\([^)]*\))?$/i, /calendar/i]
      : kind === 'previous'
        ? [/previous\s*(?:week|period|date)?/i, /предыдущ/i, /назад/i]
        : [/next\s*(?:week|period|date)?/i, /следующ/i, /впер[её]д/i, /далее/i];

    return els.find(el => {
      const text = accessibleText(el);
      const tid = normalize(`${el.getAttribute('data-tid') || ''} ${el.getAttribute('data-testid') || ''} ${el.getAttribute('aria-label') || ''} ${el.getAttribute('title') || ''}`);
      return patterns.some(re => re.test(text) || re.test(tid));
    }) || null;
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
        return { ok: false, error: `Calendar ${direction} control not found.` };
      }

      try { button.click(); } catch (_) {
        button.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window, button: 0 }));
      }
      await sleep(650);
    }

    return { ok: false, error: 'Unable to navigate Calendar to the meeting week.' };
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
    try { target.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(80);
    target.click();
    return { ok: true, url: location.href };
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
    const rejected = [];
    const debugCandidates = [];

    for (const el of pool) {
      if (!isRendered(el)) continue;
      const text = accessibleText(el);
      const score = meetingCandidateScore(el, text);

      const dom = {
        tag: el.tagName,
        role: el.getAttribute('role') || '',
        dataTid: el.getAttribute('data-tid') || '',
        dataTestId: el.getAttribute('data-testid') || '',
        ariaLabel: (el.getAttribute('aria-label') || '').slice(0, 500),
        className: typeof el.className === 'string' ? el.className.slice(0, 300) : ''
      };

      if (score < 55) {
        if ((isCalendarAggregateSlot(text) || isNonMeetingControl(text)) && rejected.length < 80) {
          rejected.push({ text: text.slice(0, 350), score, dom });
        }
        continue;
      }

      const r = el.getBoundingClientRect();
      raw.push({ el, text, score, r, dom });
      if (debugCandidates.length < 80) debugCandidates.push({ text: text.slice(0, 500), score, dom });
    }

    const dayColumns = getVisibleDayColumns();
    const calendarContext = currentCalendarMonthYear();
    const byId = new Map();

    for (const item of raw) {
      const label = item.text;
      const parentText = normalize(item.el.parentElement?.innerText || '').slice(0, 1000);
      const explicitDate = parseFlexibleDate(`${label} ${parentText}`, calendarContext?.year);
      const inferredDate = explicitDate || inferDateFromColumn(item.r, dayColumns);
      const startTime = parseStartTime(`${label} ${parentText}`);
      const dateStamp = toDateStamp(inferredDate);
      const startTimeText = startTime
        ? `${String(startTime.hour).padStart(2, '0')}:${String(startTime.minute).padStart(2, '0')}`
        : '';
      const title = cleanMeetingTitle(label);
      const id = hash(`${normalize(label)}|${dateStamp}|${startTimeText}`);
      const href = item.el.href || item.el.closest('a[href]')?.href || '';

      const meeting = {
        id,
        label,
        title,
        dateStamp,
        startTime: startTimeText,
        sortKey: toSortKey(inferredDate, startTime, byId.size),
        href,
        score: item.score,
        dom: item.dom,
        rect: {
          top: Math.round(item.r.top),
          left: Math.round(item.r.left),
          width: Math.round(item.r.width),
          height: Math.round(item.r.height)
        }
      };

      const existing = byId.get(id);
      if (!existing || meeting.score > existing.score) byId.set(id, meeting);
    }

    const meetings = Array.from(byId.values())
      .sort((a, b) => a.sortKey - b.sortKey || a.title.localeCompare(b.title, 'ru'))
      .slice(0, 120);

    lastCalendarScanDebug = {
      rejected,
      candidates: debugCandidates,
      acceptedCount: meetings.length
    };
    return meetings;
  }

  function findMeetingElement(meeting) {
    if (!meeting) return null;
    const pool = Array.from(document.querySelectorAll([
      '[data-tid*="calendar"]',
      '[data-tid*="event"]',
      '[data-testid*="calendar"]',
      '[data-testid*="event"]',
      '[role="gridcell"]',
      '[role="button"]',
      'button',
      'a[href]'
    ].join(','))).filter(isRendered);

    const exact = pool.filter(el => accessibleText(el) === meeting.label);
    if (exact.length) {
      return exact.sort((a, b) => meetingCandidateScore(b, accessibleText(b)) - meetingCandidateScore(a, accessibleText(a)))[0];
    }

    const title = normalize(meeting.title).toLowerCase();
    const fallback = pool
      .map(el => ({ el, text: accessibleText(el), r: el.getBoundingClientRect() }))
      .filter(x => x.text && normalize(x.text).toLowerCase().startsWith(title))
      .sort((a, b) => {
        const ad = Math.abs(a.r.left - (meeting.rect?.left || a.r.left)) + Math.abs(a.r.top - (meeting.rect?.top || a.r.top));
        const bd = Math.abs(b.r.left - (meeting.rect?.left || b.r.left)) + Math.abs(b.r.top - (meeting.rect?.top || b.r.top));
        return ad - bd;
      });
    return fallback[0]?.el || null;
  }

  async function openCalendarMeeting(meeting) {
    const el = findMeetingElement(meeting);
    if (!el) {
      return { ok: false, error: 'Meeting element not found in current calendar view.' };
    }

    const target = el.matches('button,a,[role="button"],[role="link"],[tabindex]')
      ? el
      : (el.querySelector('button,a,[role="button"],[role="link"],[tabindex]')
        || el.closest('button,a,[role="button"],[role="link"],[tabindex]')
        || el);

    try { target.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(120);
    try { target.focus({ preventScroll: true }); } catch (_) {}

    try {
      target.click();
    } catch (_) {
      target.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window, button: 0 }));
    }

    await sleep(1000);
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

  function findCalendarNavControl(kind) {
    const els = Array.from(document.querySelectorAll('button,[role="button"],a[href],[tabindex]')).filter(isRendered);
    const patterns = kind === 'calendar'
      ? [/^calendar(?:\s*\([^)]*\))?$/i, /calendar/i]
      : kind === 'previous'
        ? [/previous\s*(?:week|period|date)?/i, /предыдущ/i, /назад/i]
        : [/next\s*(?:week|period|date)?/i, /следующ/i, /впер[её]д/i, /далее/i];

    return els.find(el => {
      const text = accessibleText(el);
      const tid = normalize(`${el.getAttribute('data-tid') || ''} ${el.getAttribute('data-testid') || ''} ${el.getAttribute('aria-label') || ''} ${el.getAttribute('title') || ''}`);
      return patterns.some(re => re.test(text) || re.test(tid));
    }) || null;
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
        return { ok: false, error: `Calendar ${direction} control not found.` };
      }

      try { button.click(); } catch (_) {
        button.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window, button: 0 }));
      }
      await sleep(650);
    }

    return { ok: false, error: 'Unable to navigate Calendar to the meeting week.' };
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
    try { target.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(80);
    target.click();
    return { ok: true, url: location.href };
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
