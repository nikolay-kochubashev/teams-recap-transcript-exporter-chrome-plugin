(() => {
  const VERSION = '2.4.0';
  const debugState = {
    stage: 'idle',
    lastQuery: '',
    queryCandidates: [],
    peopleActions: 0,
    error: ''
  };

  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

  function normalize(value) {
    return String(value || '').replace(/\s+/g, ' ').trim();
  }

  function lines(value) {
    return String(value || '').split(/\r?\n/).map(normalize).filter(Boolean);
  }

  function isRendered(el) {
    if (!el) return false;
    const r = el.getBoundingClientRect();
    const style = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && style.display !== 'none' && style.visibility !== 'hidden';
  }

  async function waitFor(factory, timeout = 15000, interval = 120) {
    const started = Date.now();
    let lastError = null;
    while (Date.now() - started < timeout) {
      try {
        const value = factory();
        if (value) return value;
      } catch (e) {
        lastError = e;
      }
      await sleep(interval);
    }
    if (lastError) throw lastError;
    throw new Error('Timeout waiting for Teams UI.');
  }

  function clickElement(el) {
    if (!el) throw new Error('Teams control was not found.');
    el.scrollIntoView({ block: 'center', inline: 'nearest' });
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
    el.click();
  }

  function setInputValue(input, value) {
    const descriptor = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
    if (descriptor?.set) descriptor.set.call(input, value);
    else input.value = value;

    input.dispatchEvent(new InputEvent('input', {
      bubbles: true,
      cancelable: false,
      data: value,
      inputType: 'insertText'
    }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function parseSelfName(value) {
    const text = normalize(value);
    const match = text.match(/([^|,\n]{3,160}?)\s*\(You\)/i);
    return match ? normalize(match[1]) : '';
  }

  function detectSelfName() {
    const titleMatch = String(document.title || '').match(/\|\s*([^|]+?)\s*\(You\)\s*\|/i);
    if (titleMatch) return normalize(titleMatch[1]);

    const candidates = Array.from(document.querySelectorAll(
      '[aria-label*="(You)"],[title*="(You)"],[data-tid^="AUTOSUGGEST_SUGGESTION_PEOPLE"]'
    ));
    for (const el of candidates) {
      const name = parseSelfName(
        el.getAttribute('aria-label') || el.getAttribute('title') || el.innerText || el.textContent
      );
      if (name) return name;
    }

    const bodyText = document.body?.innerText || '';
    return parseSelfName(bodyText);
  }

  function formatDateStamp(stamp) {
    if (!/^\d{8}$/.test(String(stamp || ''))) throw new Error('Invalid date stamp: ' + stamp);
    const yyyy = Number(stamp.slice(0, 4));
    const mm = Number(stamp.slice(4, 6));
    const dd = Number(stamp.slice(6, 8));
    const date = new Date(yyyy, mm - 1, dd);
    const locale = document.documentElement.lang || 'en-GB';
    return new Intl.DateTimeFormat(locale, {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    }).format(date);
  }

  function findSearchInput() {
    const candidates = [
      document.querySelector('#ms-searchux-input'),
      document.querySelector('[data-tid="AUTOSUGGEST_INPUT"]'),
      document.querySelector('input[placeholder*="people"][type="search"]')
    ];
    return candidates.find(isRendered) || null;
  }

  async function ensureSearchInput() {
    let input = findSearchInput();
    if (input) return input;

    const toggle = Array.from(document.querySelectorAll(
      '[data-tid="title-bar-toggle-search-btn"],button[aria-label*="Search"],button[title*="Search"]'
    )).find(isRendered);

    if (toggle) clickElement(toggle);
    input = await waitFor(findSearchInput, 8000);
    return input;
  }

  function personSearchActions() {
    // Do NOT require getBoundingClientRect() here. In Teams autosuggest some
    // icon buttons are visually rendered by Fluent UI while the button itself
    // can temporarily report a zero-sized rect. The DOM node is still the
    // correct actionable control.
    const selectors = [
      'button[data-tid="AUTOSUGGEST_ACTION_PEOPLECENTRICSEARCH"]',
      'button[aria-label^="All results from "]',
      '[aria-label^="All results from "] > button',
      '[aria-label^="All results from "] button'
    ];

    const seen = new Set();
    const result = [];

    for (const selector of selectors) {
      for (const el of document.querySelectorAll(selector)) {
        if (seen.has(el)) continue;
        seen.add(el);
        result.push(el);
      }
    }

    return result;
  }

  function selfSuggestionName() {
    const suggestions = Array.from(document.querySelectorAll(
      '[data-tid^="AUTOSUGGEST_SUGGESTION_PEOPLE"]'
    )).filter(isRendered);

    for (const el of suggestions) {
      const name = parseSelfName(el.getAttribute('aria-label') || el.innerText || el.textContent);
      if (name) return name;
    }
    return '';
  }

  function peopleQueryCandidates(query) {
    const full = normalize(query);
    const first = normalize(full.split(/\s+/)[0] || '');
    const short = first.length > 8 ? first.slice(0, 8) : first;

    return Array.from(new Set([first, short, full].filter(x => x.length >= 3)));
  }

  async function waitForPeopleAction(desired, timeout = 4500) {
    const started = Date.now();

    while (Date.now() - started < timeout) {
      const actions = personSearchActions();
      debugState.peopleActions = actions.length;

      let action = actions.find(el => {
        const label = normalize(
          el.getAttribute('aria-label') ||
          el.closest('[aria-label^="All results from "]')?.getAttribute('aria-label') ||
          el.parentElement?.getAttribute('aria-label') ||
          ''
        ).toLocaleLowerCase();
        return desired && label.includes(desired.toLocaleLowerCase());
      });

      if (!action) {
        const detected = selfSuggestionName();
        if (detected) {
          action = actions.find(el =>
            normalize(el.getAttribute('aria-label')).toLocaleLowerCase()
              .includes(detected.toLocaleLowerCase())
          );
        }
      }

      if (!action && actions.length === 1) action = actions[0];
      if (action) return action;

      await sleep(120);
    }

    return null;
  }

  async function openPeopleCentricSearch(query) {
    debugState.stage = 'people-search';
    debugState.error = '';

    const input = await ensureSearchInput();
    const candidates = peopleQueryCandidates(query);
    debugState.queryCandidates = candidates;

    if (!candidates.length) {
      throw new Error('Не удалось сформировать поисковый запрос для People.');
    }

    let action = null;
    let usedQuery = '';

    for (const candidate of candidates) {
      debugState.lastQuery = candidate;
      usedQuery = candidate;

      input.focus();
      setInputValue(input, '');
      await sleep(120);
      setInputValue(input, candidate);

      action = await waitForPeopleAction(candidate, 4500);
      if (action) break;
    }

    if (!action) {
      debugState.error = 'people-autosuggest-not-found';
      throw new Error(
        'Teams не показал People autosuggest для текущего пользователя. ' +
        'Пробовал: ' + candidates.join(', ') + '.'
      );
    }

    const label = normalize(
      action.getAttribute('aria-label') ||
      action.closest('[aria-label^="All results from "]')?.getAttribute('aria-label') ||
      action.parentElement?.getAttribute('aria-label') ||
      ''
    );
    const author = label.replace(/^All results from\s+/i, '').trim() ||
      selfSuggestionName() || query;

    debugState.stage = 'open-people-result';
    clickElement(action);

    try {
      await waitFor(() => {
        const content = document.querySelector('[data-tid="search-content"]');
        const filter = document.querySelector('button[data-tid="search-people-filter"]');
        return content && filter && isRendered(content) ? content : null;
      }, 18000);
    } catch (_) {
      debugState.error = 'search-results-not-opened';
      throw new Error(
        'Teams принял People result, но страница Search Results не открылась. ' +
        'Последний запрос: ' + usedQuery + '.'
      );
    }

    debugState.stage = 'people-filter-ready';
    return author;
  }

  function findDatePopup() {
    return Array.from(document.querySelectorAll('[data-tid="search-date-filter"]'))
      .find(el => el.querySelector('input[placeholder="Select a Date"]') && isRendered(el)) || null;
  }

  async function applyDateRange(startStamp, endStamp) {
    const dateButton = Array.from(document.querySelectorAll('button[data-tid="search-date-filter"]'))
      .find(isRendered);
    if (!dateButton) throw new Error('Фильтр Date в Teams Search не найден.');

    clickElement(dateButton);
    const popup = await waitFor(findDatePopup, 8000);
    const dateInputs = Array.from(popup.querySelectorAll('input[placeholder="Select a Date"]'));

    let fromInput = dateInputs.find(x => /^From/i.test(x.getAttribute('aria-label') || '')) || dateInputs[0];
    let toInput = dateInputs.find(x => /^To/i.test(x.getAttribute('aria-label') || '')) || dateInputs[1];

    if (!fromInput || !toInput) throw new Error('Поля From/To фильтра Date не найдены.');

    const fromText = formatDateStamp(startStamp);
    const toText = formatDateStamp(endStamp);

    fromInput.removeAttribute('readonly');
    toInput.removeAttribute('readonly');
    fromInput.focus();
    setInputValue(fromInput, fromText);
    fromInput.dispatchEvent(new Event('blur', { bubbles: true }));
    await sleep(150);
    toInput.focus();
    setInputValue(toInput, toText);
    toInput.dispatchEvent(new Event('blur', { bubbles: true }));
    await sleep(250);

    const apply = Array.from(popup.querySelectorAll('button')).find(btn => {
      const aria = normalize(btn.getAttribute('aria-label'));
      const text = normalize(btn.innerText || btn.textContent);
      return /Apply the Selected Date Range/i.test(aria) || /^Apply$/i.test(text);
    });

    if (!apply) throw new Error('Кнопка Apply в фильтре Date не найдена.');

    await waitFor(() => !apply.disabled ? apply : null, 4000).catch(() => {
      throw new Error('Teams не принял диапазон дат. Нужна диагностика Date filter.');
    });

    clickElement(apply);

    const appliedButton = await waitFor(() => {
      const visiblePopup = findDatePopup();
      const button = Array.from(document.querySelectorAll('button[data-tid="search-date-filter"]'))
        .find(isRendered);
      const label = normalize(button?.getAttribute('aria-label'));
      const hasFrom = label.toLocaleLowerCase().includes(fromText.toLocaleLowerCase());
      const hasTo = label.toLocaleLowerCase().includes(toText.toLocaleLowerCase());
      return !visiblePopup && button && hasFrom && hasTo ? button : null;
    }, 10000).catch(() => null);

    if (!appliedButton) {
      throw new Error('Teams закрыл Date filter, но примененный диапазон не подтвержден. Сбор остановлен, чтобы не сохранить сообщения за неверные даты.');
    }

    return { fromText, toText };
  }

  async function openMessagesResults() {
    await sleep(500);

    const moreMessages = document.querySelector('button[data-tid="more-Messages"]');
    if (moreMessages && isRendered(moreMessages)) {
      clickElement(moreMessages);
    } else {
      const messagesTab = document.querySelector('[data-tid="messages-tab"]');
      if (messagesTab && isRendered(messagesTab)) clickElement(messagesTab);
    }

    await waitFor(() => {
      const cards = document.querySelectorAll('[data-tid="search-card"]');
      const pagination = document.querySelector('[data-tid="search-pagination-previous-next"]');
      return cards.length || pagination ? true : false;
    }, 15000);
  }

  function directCells(row, role) {
    return Array.from(row.children || []).filter(el => el.getAttribute?.('role') === role);
  }

  function messageTextWithLinks(cell) {
    if (!cell) return '';
    const clone = cell.cloneNode(true);
    const links = [];
    for (const a of Array.from(clone.querySelectorAll('a[href]'))) {
      const href = a.href || a.getAttribute('href') || '';
      const label = normalize(a.innerText || a.textContent);
      if (/^https?:/i.test(href)) {
        links.push(href);
        if (label && label !== href) a.textContent = label + ' (' + href + ')';
        else a.textContent = href;
      }
    }
    return {
      text: normalize(clone.innerText || clone.textContent),
      links: Array.from(new Set(links))
    };
  }

  function parseMessageRow(row) {
    if (!row?.querySelector('[data-tid="search-message-card"]')) return null;

    const rowHeader = directCells(row, 'rowheader')[0] || row.querySelector('[role="rowheader"]');
    const cardRoot = row.closest('[data-tid="search-card"]');
    const appHeader = row.querySelector('[data-tid="message-app-card-header"]') ||
      cardRoot?.querySelector('[data-tid="message-app-card-header"]');
    const conversationTitle =
      rowHeader?.querySelector('[title]')?.getAttribute('title') ||
      appHeader?.querySelector('[title]')?.getAttribute('title') ||
      '';
    const headerLines = lines(
      rowHeader?.innerText ||
      appHeader?.innerText ||
      rowHeader?.textContent ||
      appHeader?.textContent ||
      ''
    );
    const conversation = normalize(
      conversationTitle ||
      headerLines.find(x => !/^Send feedback$/i.test(x)) ||
      ''
    );

    const cells = directCells(row, 'gridcell');
    const senderLines = lines(cells[0]?.innerText || cells[0]?.textContent || '');
    const sender = senderLines[0] || '';
    const timestamp = senderLines.slice(1).join(' ') || '';

    const body = messageTextWithLinks(cells[1]);
    const contentNode = row.querySelector('[id^="serp-message-card-content-"]');
    const messageId = contentNode?.id?.replace(/^serp-message-card-content-/, '') || '';

    const aria = normalize(row.getAttribute('aria-label'));
    const message = body.text || aria
      .replace(/^Message from .*?\./i, '')
      .replace(/\.\s*Press Enter\/Spacebar.*$/i, '')
      .trim();

    if (!message && !conversation) return null;

    const key = messageId || [conversation, sender, timestamp, message].join('|');
    return {
      id: messageId,
      key,
      conversation,
      sender,
      timestamp,
      message,
      links: body.links,
      aria
    };
  }

  function currentMessageRows() {
    const searchContent = document.querySelector('[data-tid="search-content"]') || document;
    const rows = Array.from(searchContent.querySelectorAll('[data-tid="search-card"] [role="row"]'));
    const unique = [];
    const seen = new Set();

    for (const row of rows) {
      const item = parseMessageRow(row);
      if (!item || seen.has(item.key)) continue;
      seen.add(item.key);
      unique.push(item);
    }
    return unique;
  }

  function pageSignature() {
    const items = currentMessageRows();
    return items.slice(0, 3).concat(items.slice(-2)).map(x => x.key).join('||');
  }

  function paginationNext() {
    const root = document.querySelector('[data-tid="search-pagination-previous-next"]');
    if (!root) return null;
    const buttons = Array.from(root.querySelectorAll('button'));
    return buttons.find(btn => /^(Next|Далее)$/i.test(normalize(btn.innerText || btn.textContent))) || null;
  }

  function hasNextPage() {
    const btn = paginationNext();
    return !!btn && !btn.disabled && btn.getAttribute('aria-disabled') !== 'true' && isRendered(btn);
  }

  async function goNextPage() {
    const button = paginationNext();
    if (!button || button.disabled || button.getAttribute('aria-disabled') === 'true') {
      return { ok: true, moved: false };
    }

    const before = pageSignature();
    clickElement(button);

    await waitFor(() => {
      const after = pageSignature();
      return after && after !== before ? after : null;
    }, 15000);

    return { ok: true, moved: true, signature: pageSignature() };
  }

  async function prepareSearch(request) {
    debugState.stage = 'prepare';
    debugState.error = '';
    const requested = normalize(request?.authorQuery);
    const autoDetected = detectSelfName();
    const query = requested || autoDetected;

    if (!query) {
      throw new Error('Не удалось автоматически определить твое имя в Teams. Укажи имя или фамилию в поле расширения.');
    }

    const author = await openPeopleCentricSearch(query);
    debugState.stage = 'date-filter';
    const dateRange = await applyDateRange(request.startDate, request.endDate);
    debugState.stage = 'messages-results';
    await openMessagesResults();
    debugState.stage = 'ready';

    return {
      ok: true,
      author,
      query,
      startDate: request.startDate,
      endDate: request.endDate,
      dateRange,
      pageTitle: document.title,
      url: location.href
    };
  }

  function collectPage() {
    const items = currentMessageRows();
    return {
      ok: true,
      items,
      count: items.length,
      signature: pageSignature(),
      hasNext: hasNextPage(),
      pageTitle: document.title,
      url: location.href
    };
  }


  function localDateStamp(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
    return String(date.getFullYear()) +
      String(date.getMonth() + 1).padStart(2, '0') +
      String(date.getDate()).padStart(2, '0');
  }

  function localDateTimeText(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
    return String(date.getDate()).padStart(2, '0') + '.' +
      String(date.getMonth() + 1).padStart(2, '0') + '.' +
      String(date.getFullYear()) + ' ' +
      String(date.getHours()).padStart(2, '0') + ':' +
      String(date.getMinutes()).padStart(2, '0');
  }

  function cleanMultiline(value) {
    const rows = String(value || '')
      .replace(/\r/g, '')
      .split('\n')
      .map(x => x.replace(/[ \t\u00a0]+/g, ' ').trim());

    const out = [];
    let blank = false;
    for (const row of rows) {
      if (!row) {
        if (!blank && out.length) out.push('');
        blank = true;
        continue;
      }
      blank = false;
      out.push(row);
    }
    while (out.length && !out[out.length - 1]) out.pop();
    return out.join('\n').trim();
  }

  function chatViewport() {
    return document.querySelector(
      '[data-tid="right-rail-message-pane-body"] [data-tid="message-pane-list-viewport"]'
    ) || document.querySelector('[data-tid="message-pane-list-viewport"]');
  }

  function chatTitle() {
    const title = document.querySelector('[data-tid="chat-title"]');
    return normalize(title?.innerText || title?.textContent || '');
  }

  function searchRowForItem(item) {
    if (item?.id) {
      const content = document.getElementById('serp-message-card-content-' + item.id);
      const row = content?.closest('[role="row"]');
      if (row) return row;
    }

    const wantedConversation = normalize(item?.conversation).toLocaleLowerCase();
    const wantedMessage = normalize(item?.message).slice(0, 80).toLocaleLowerCase();

    for (const row of Array.from(document.querySelectorAll(
      '[data-tid="search-content"] [data-tid="search-card"] [role="row"]'
    ))) {
      const parsed = parseMessageRow(row);
      if (!parsed) continue;
      const sameConversation = wantedConversation &&
        normalize(parsed.conversation).toLocaleLowerCase() === wantedConversation;
      const sameMessage = wantedMessage &&
        normalize(parsed.message).toLocaleLowerCase().includes(wantedMessage);
      if (sameConversation && (!wantedMessage || sameMessage)) return row;
    }
    return null;
  }

  function chatMessageText(contentNode) {
    if (!contentNode) return { text: '', links: [], quote: null };

    const quoteNode = contentNode.querySelector('[data-tid="quoted-reply-card"]');
    let quote = null;
    if (quoteNode) {
      const quoteTimestamp = normalize(
        quoteNode.querySelector('[data-tid="quoted-reply-timestamp"]')?.innerText ||
        quoteNode.querySelector('[data-tid="quoted-reply-timestamp"]')?.textContent ||
        ''
      );
      const quotePreview = cleanMultiline(
        quoteNode.querySelector('[data-tid="quoted-reply-preview-content"]')?.innerText ||
        quoteNode.querySelector('[data-tid="quoted-reply-preview-content"]')?.textContent ||
        ''
      );
      const quoteLines = lines(quoteNode.innerText || quoteNode.textContent || '');
      const quoteAuthor = quoteLines.find(x => x !== quoteTimestamp && x !== quotePreview) || '';
      quote = {
        author: normalize(quoteAuthor),
        timestamp: quoteTimestamp,
        preview: quotePreview
      };
    }

    const clone = contentNode.cloneNode(true);
    clone.querySelectorAll('[data-tid="quoted-reply-card"]').forEach(x => x.remove());
    clone.querySelectorAll('button,svg,[role="button"]').forEach(x => x.remove());

    const links = [];
    for (const a of Array.from(clone.querySelectorAll('a[href]'))) {
      const href = a.href || a.getAttribute('href') || '';
      const label = normalize(a.innerText || a.textContent);
      if (/^https?:/i.test(href)) {
        links.push(href);
        a.textContent = label && label !== href ? label + ' (' + href + ')' : href;
      }
    }

    return {
      text: cleanMultiline(clone.innerText || clone.textContent || ''),
      links: Array.from(new Set(links)),
      quote
    };
  }

  function parseChatItem(item) {
    const body = item?.querySelector('[data-tid="chat-pane-message"]');
    if (!body) return null;

    const mid = normalize(body.getAttribute('data-mid'));
    const time = item.querySelector('time[datetime]');
    let epoch = Number(mid);
    if (!Number.isFinite(epoch) || epoch < 1000000000000) {
      epoch = Date.parse(time?.getAttribute('datetime') || '');
    }
    if (!Number.isFinite(epoch)) return null;

    const date = new Date(epoch);
    const content = item.querySelector('[data-message-content]') || body;
    const extracted = chatMessageText(content);
    const author = normalize(
      item.querySelector('[data-tid="message-author-name"]')?.innerText ||
      item.querySelector('[data-tid="message-author-name"]')?.textContent ||
      ''
    );

    const key = mid || [
      epoch,
      author,
      extracted.text,
      extracted.quote?.preview || ''
    ].join('|');

    return {
      id: mid,
      key,
      epoch,
      stamp: localDateStamp(date),
      dateTime: date.toISOString(),
      timestamp: localDateTimeText(date),
      author,
      message: extracted.text,
      quote: extracted.quote,
      links: extracted.links
    };
  }

  function visibleChatMessages(viewport) {
    const root = viewport || chatViewport();
    if (!root) return [];

    const out = [];
    const seen = new Set();
    for (const item of Array.from(root.querySelectorAll('[data-tid="chat-pane-item"]'))) {
      const parsed = parseChatItem(item);
      if (!parsed || seen.has(parsed.key)) continue;
      seen.add(parsed.key);
      out.push(parsed);
    }
    return out;
  }

  function viewportSignature(viewport) {
    const items = visibleChatMessages(viewport);
    return items.slice(0, 2).concat(items.slice(-2))
      .map(x => x.key)
      .join('||') + '::' + Math.round(viewport?.scrollTop || 0);
  }

  async function nudgeChatViewport(viewport, direction) {
    const beforeSignature = viewportSignature(viewport);
    const beforeTop = Number(viewport.scrollTop || 0);
    const step = Math.max(260, Math.floor((viewport.clientHeight || 600) * 0.68));
    const target = Math.max(0, beforeTop + (direction < 0 ? -step : step));

    viewport.scrollTop = target;
    viewport.dispatchEvent(new Event('scroll', { bubbles: true }));

    for (let i = 0; i < 12; i++) {
      await sleep(i === 0 ? 220 : 120);
      const afterSignature = viewportSignature(viewport);
      const afterTop = Number(viewport.scrollTop || 0);
      if (afterSignature !== beforeSignature || Math.abs(afterTop - beforeTop) > 1) {
        return { moved: true, beforeTop, afterTop };
      }
    }

    return { moved: false, beforeTop, afterTop: Number(viewport.scrollTop || 0) };
  }

  function titleLooksLike(actual, expected) {
    const a = normalize(actual).toLocaleLowerCase();
    const e = normalize(expected).toLocaleLowerCase();
    if (!a || !e) return false;
    return a === e || a.includes(e) || e.includes(a);
  }

  function paneContainsMessage(viewport, message) {
    const needle = normalize(message).slice(0, 70).toLocaleLowerCase();
    if (!needle) return false;
    return visibleChatMessages(viewport).some(x =>
      normalize(x.message).toLocaleLowerCase().includes(needle)
    );
  }

  async function openConversationFromSearch(item) {
    const row = searchRowForItem(item);
    if (!row) throw new Error('Search result row для переписки не найден.');

    const beforeTitle = chatTitle();
    const beforeViewport = chatViewport();
    const beforeSignature = beforeViewport ? viewportSignature(beforeViewport) : '';

    clickElement(row);

    let viewport = await waitFor(() => {
      const current = chatViewport();
      if (!current) return null;

      const currentTitle = chatTitle();
      if (titleLooksLike(currentTitle, item?.conversation)) return current;
      if (paneContainsMessage(current, item?.message)) return current;

      const currentSignature = viewportSignature(current);
      if (currentTitle && currentTitle !== beforeTitle && currentSignature !== beforeSignature) {
        return current;
      }
      return null;
    }, 10000).catch(() => null);

    if (!viewport) {
      try {
        row.focus();
        row.dispatchEvent(new KeyboardEvent('keydown', {
          key: 'Enter',
          code: 'Enter',
          bubbles: true,
          cancelable: true
        }));
      } catch (_) {}

      viewport = await waitFor(() => {
        const current = chatViewport();
        if (!current) return null;
        if (titleLooksLike(chatTitle(), item?.conversation)) return current;
        if (paneContainsMessage(current, item?.message)) return current;
        return null;
      }, 8000).catch(() => null);
    }

    if (!viewport) {
      throw new Error('Teams не открыл контекст переписки для выбранного Search result.');
    }

    await sleep(350);
    return viewport;
  }

  async function collectConversationContext(request) {
    const item = request?.item || {};
    const startDate = String(request?.startDate || '');
    const endDate = String(request?.endDate || '');

    if (!/^\d{8}$/.test(startDate) || !/^\d{8}$/.test(endDate)) {
      throw new Error('Некорректный период для чтения переписки.');
    }

    debugState.stage = 'conversation-open';
    const viewport = await openConversationFromSearch(item);
    const actualTitle = chatTitle() || normalize(item.conversation) || 'Conversation';

    debugState.stage = 'conversation-scroll';
    const collected = new Map();

    const addVisible = () => {
      for (const message of visibleChatMessages(viewport)) {
        if (!message.key) continue;
        collected.set(message.key, message);
      }
    };

    const minStamp = () => {
      const values = Array.from(collected.values()).map(x => x.stamp).filter(Boolean).sort();
      return values[0] || '';
    };

    const maxStamp = () => {
      const values = Array.from(collected.values()).map(x => x.stamp).filter(Boolean).sort();
      return values.length ? values[values.length - 1] : '';
    };

    addVisible();

    let upStable = 0;
    let reachedStart = false;
    for (let i = 0; i < 180; i++) {
      if (minStamp() && minStamp() < startDate) {
        reachedStart = true;
        break;
      }

      const moved = await nudgeChatViewport(viewport, -1);
      addVisible();

      if (!moved.moved || Number(viewport.scrollTop || 0) <= 1) upStable++;
      else upStable = 0;

      if (upStable >= 3) {
        reachedStart = true;
        break;
      }
    }

    let downStable = 0;
    let reachedEnd = false;
    for (let i = 0; i < 260; i++) {
      if (maxStamp() && maxStamp() > endDate) {
        reachedEnd = true;
        break;
      }

      const moved = await nudgeChatViewport(viewport, 1);
      addVisible();

      const atBottom =
        Math.abs(
          Number(viewport.scrollHeight || 0) -
          Number(viewport.clientHeight || 0) -
          Number(viewport.scrollTop || 0)
        ) <= 3;

      if (!moved.moved || atBottom) downStable++;
      else downStable = 0;

      if (downStable >= 3) {
        reachedEnd = true;
        break;
      }
    }

    addVisible();

    const messages = Array.from(collected.values())
      .filter(x => x.stamp >= startDate && x.stamp <= endDate)
      .sort((a, b) => a.epoch - b.epoch || String(a.key).localeCompare(String(b.key)));

    debugState.stage = 'ready';

    return {
      ok: true,
      conversation: actualTitle,
      requestedConversation: normalize(item.conversation),
      messages,
      count: messages.length,
      loaded: collected.size,
      reachedStart,
      reachedEnd
    };
  }

  function diagnostic() {
    const dateButton = Array.from(document.querySelectorAll('button[data-tid="search-date-filter"]')).find(isRendered);
    const peopleButton = Array.from(document.querySelectorAll('button[data-tid="search-people-filter"]')).find(isRendered);
    const next = paginationNext();
    const items = currentMessageRows();

    const info = {
      version: VERSION,
      stage: debugState.stage,
      lastQuery: debugState.lastQuery,
      queryCandidates: debugState.queryCandidates,
      peopleActions: debugState.peopleActions,
      peopleActionsRaw: personSearchActions().length,
      peopleActionsRendered: personSearchActions().filter(isRendered).length,
      allResultsWrappers: Array.from(document.querySelectorAll('[aria-label^="All results from "]'))
        .filter(isRendered)
        .slice(0, 10)
        .map(el => ({
          tag: el.tagName,
          aria: normalize(el.getAttribute('aria-label')),
          dataTid: el.getAttribute('data-tid') || ''
        })),
      adapterError: debugState.error,
      url: location.href,
      title: document.title,
      lang: document.documentElement.lang || '',
      selfDetected: detectSelfName(),
      searchInput: !!findSearchInput(),
      searchContent: !!document.querySelector('[data-tid="search-content"]'),
      peopleFilter: peopleButton?.getAttribute('aria-label') || '',
      dateFilter: dateButton?.getAttribute('aria-label') || '',
      messageCards: document.querySelectorAll('[data-tid="search-card"]').length,
      parsedMessages: items.length,
      pagination: !!document.querySelector('[data-tid="search-pagination-previous-next"]'),
      nextText: normalize(next?.innerText || ''),
      nextDisabled: next ? !!next.disabled : null,
      sample: items.slice(0, 10),
      chatContext: {
        title: chatTitle(),
        viewport: !!chatViewport(),
        visibleMessages: visibleChatMessages(chatViewport()).length,
        firstVisible: visibleChatMessages(chatViewport())[0] || null,
        lastVisible: visibleChatMessages(chatViewport()).slice(-1)[0] || null
      }
    };

    return JSON.stringify(info, null, 2);
  }

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    const type = message?.type;

    if (type === 'CHAT_DETECT_SELF') {
      sendResponse({ ok: true, author: detectSelfName(), version: VERSION });
      return;
    }

    if (type === 'CHAT_PREPARE_SEARCH') {
      prepareSearch(message).then(sendResponse).catch(error => {
        sendResponse({ ok: false, error: error.message || String(error), diagnostic: diagnostic() });
      });
      return true;
    }

    if (type === 'CHAT_COLLECT_PAGE') {
      try {
        sendResponse(collectPage());
      } catch (error) {
        sendResponse({ ok: false, error: error.message || String(error), diagnostic: diagnostic() });
      }
      return;
    }

    if (type === 'CHAT_GO_NEXT') {
      goNextPage().then(sendResponse).catch(error => {
        sendResponse({ ok: false, error: error.message || String(error), diagnostic: diagnostic() });
      });
      return true;
    }

    if (type === 'CHAT_COLLECT_CONVERSATION_CONTEXT') {
      collectConversationContext(message).then(sendResponse).catch(error => {
        sendResponse({ ok: false, error: error.message || String(error), diagnostic: diagnostic() });
      });
      return true;
    }

    if (type === 'CHAT_SEARCH_DIAGNOSTIC') {
      sendResponse({ ok: true, text: diagnostic() });
      return;
    }
  });
})();