(() => {
  const VERSION = '2.2.0';

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
    return Array.from(document.querySelectorAll(
      'button[data-tid="AUTOSUGGEST_ACTION_PEOPLECENTRICSEARCH"]'
    )).filter(isRendered);
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

  async function openPeopleCentricSearch(query) {
    const input = await ensureSearchInput();
    input.focus();
    setInputValue(input, '');
    await sleep(80);
    setInputValue(input, query);

    await waitFor(() => {
      const popup = document.querySelector('[data-tid="ms-searchux-popup"]');
      return popup && isRendered(popup) && personSearchActions().length ? popup : null;
    }, 12000);

    const detected = selfSuggestionName();
    const desired = (detected || query || '').toLocaleLowerCase();
    const actions = personSearchActions();

    let action = actions.find(el => {
      const label = normalize(el.getAttribute('aria-label')).toLocaleLowerCase();
      return desired && label.includes(desired);
    });

    if (!action && detected) {
      action = actions.find(el =>
        normalize(el.getAttribute('aria-label')).toLocaleLowerCase().includes(detected.toLocaleLowerCase())
      );
    }
    if (!action && actions.length === 1) action = actions[0];

    if (!action) {
      throw new Error('Не найден результат People для текущего пользователя. Укажи точнее имя или фамилию в Teams.');
    }

    const label = normalize(action.getAttribute('aria-label'));
    const author = label.replace(/^All results from\s+/i, '').trim() || detected || query;

    clickElement(action);

    await waitFor(() => {
      const content = document.querySelector('[data-tid="search-content"]');
      const filter = document.querySelector('button[data-tid="search-people-filter"]');
      return content && filter && isRendered(content) ? content : null;
    }, 18000);

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
    if (apply.disabled) {
      throw new Error('Teams не принял диапазон дат. Нужна диагностика Date filter.');
    }

    clickElement(apply);

    await waitFor(() => {
      const visiblePopup = findDatePopup();
      const button = Array.from(document.querySelectorAll('button[data-tid="search-date-filter"]'))
        .find(isRendered);
      const label = normalize(button?.getAttribute('aria-label'));
      return !visiblePopup && button && !/^Date filter$/i.test(label) ? button : null;
    }, 10000).catch(() => null);

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
    const conversationTitle = rowHeader?.querySelector('[title]')?.getAttribute('title') || '';
    const headerLines = lines(rowHeader?.innerText || '');
    const conversation = normalize(conversationTitle || headerLines.find(x => !/^Send feedback$/i.test(x)) || '');

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
    const requested = normalize(request?.authorQuery);
    const autoDetected = detectSelfName();
    const query = requested || autoDetected;

    if (!query) {
      throw new Error('Не удалось автоматически определить твое имя в Teams. Укажи имя или фамилию в поле расширения.');
    }

    const author = await openPeopleCentricSearch(query);
    const dateRange = await applyDateRange(request.startDate, request.endDate);
    await openMessagesResults();

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

  function diagnostic() {
    const dateButton = Array.from(document.querySelectorAll('button[data-tid="search-date-filter"]')).find(isRendered);
    const peopleButton = Array.from(document.querySelectorAll('button[data-tid="search-people-filter"]')).find(isRendered);
    const next = paginationNext();
    const items = currentMessageRows();

    const info = {
      version: VERSION,
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
      sample: items.slice(0, 10)
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

    if (type === 'CHAT_SEARCH_DIAGNOSTIC') {
      sendResponse({ ok: true, text: diagnostic() });
      return;
    }
  });
})();