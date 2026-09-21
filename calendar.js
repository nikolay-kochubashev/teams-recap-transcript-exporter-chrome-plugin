(() => {
  if (window.__teamsTranscriptCalendarLoaded) return;
  window.__teamsTranscriptCalendarLoaded = true;

  const VERSION = '2.0.1';
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
    if (isCalendarAggregateSlot(text)) return -1000;
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
    s = s.replace(/\b(?:[01]?\d|2[0-3]):[0-5]\d\b\s*(?:-|–|—|to)?\s*\b(?:[01]?\d|2[0-3]):[0-5]\d\b/gi, ' ');
    s = s.replace(/\b(?:1[0-2]|0?[1-9]):[0-5]\d\s*(?:AM|PM)\b\s*(?:-|–|—|to)?\s*\b(?:1[0-2]|0?[1-9]):[0-5]\d\s*(?:AM|PM)\b/gi, ' ');
    s = s.replace(/\b(?:Accepted|Tentative|Busy|Free|Organizer|Microsoft Teams meeting|Teams meeting)\b/gi, ' ');
    s = s.replace(/\s+/g, ' ').trim();
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

      if (isCalendarAggregateSlot(text)) {
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
      const dateStamp = dateStampFromText(`${label} ${parentText}`);

      meetings.push({
        id,
        label,
        title: cleanMeetingTitle(label),
        dateStamp,
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
    try { el.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (_) {}
    await sleep(120);
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
    el.click();
    await sleep(550);
    return { ok: true };
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
    meetings.forEach((m, i) => lines.push(`[M${i + 1}] id=${m.id} score=${m.score} date=${m.dateStamp || '-'} rect=${JSON.stringify(m.rect)} dom=${JSON.stringify(m.dom || {})} label=${m.label}`));
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
