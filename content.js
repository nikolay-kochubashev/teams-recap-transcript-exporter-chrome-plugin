(() => {
  if (window.__teamsRecapTranscriptExporterLoaded) return;
  window.__teamsRecapTranscriptExporterLoaded = true;

  const VERSION = '1.7.0';
  const state = {
    status: 'idle',
    message: 'Готово к работе.',
    progress: 0,
    items: 0,
    chars: 0,
    title: 'teams-transcript',
    text: '',
    error: '',
    startedAt: 0,
    finishedAt: 0
  };

  let cancelRequested = false;
  let runningPromise = null;
  let lastRunDebug = {};

  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

  const normalize = value => (value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  const normalizeLine = value => (value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .trim();

  const sanitizeFileName = value => (value || 'teams-transcript')
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120) || 'teams-transcript';

  const publicState = () => ({ ...state });
  const update = patch => Object.assign(state, patch);

  function isRendered(el) {
    if (!el || !(el instanceof Element) || !el.isConnected) return false;
    const style = getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden' || Number(style.opacity) === 0) return false;
    const r = el.getBoundingClientRect();
    return r.width > 20 && r.height > 20;
  }

  function isInViewport(el) {
    if (!isRendered(el)) return false;
    const r = el.getBoundingClientRect();
    return r.bottom > 0 && r.right > 0 && r.top < innerHeight && r.left < innerWidth;
  }

  function isScrollable(el) {
    if (!isRendered(el)) return false;
    return el.scrollHeight > el.clientHeight + 40;
  }

  const clockTimeRegex = /(?:^|\s)(?:\d{1,2}:)?\d{1,2}:\d{2}(?:\s|$)/g;
  const exactClockRegex = /^(?:\d{1,2}:)?\d{1,2}:\d{2}$/;
  const spokenTimeRegex = /\b\d+\s+(?:hours?|minutes?|seconds?)\b/i;
  const transcriptWord = /transcript|transcription|транскрип/i;

  function nearestScrollableAncestor(el) {
    let node = el?.parentElement || null;
    for (let depth = 0; node && depth < 16; depth++, node = node.parentElement) {
      if (isScrollable(node)) return node;
    }
    return null;
  }

  function getTranscriptHeadingRects() {
    const result = [];
    const selectors = 'h1,h2,h3,h4,h5,h6,[role="heading"],[role="tab"],button,[aria-label]';
    for (const el of document.querySelectorAll(selectors)) {
      const text = normalizeLine(el.innerText || el.getAttribute('aria-label') || '');
      if (!/^transcript(?:ion)?$/i.test(text)) continue;
      const r = el.getBoundingClientRect();
      if (r.width <= 0 || r.height <= 0) continue;
      result.push({ x: r.x, y: r.y, left: r.left, right: r.right, top: r.top, bottom: r.bottom, width: r.width, height: r.height });
    }
    return result;
  }

  function transcriptEntryVotes() {
    const votes = new Map();
    const addVote = (scroller, amount, why) => {
      if (!scroller) return;
      const current = votes.get(scroller) || { count: 0, reasons: new Set() };
      current.count += amount;
      current.reasons.add(why);
      votes.set(scroller, current);
    };

    // The nearest scrollable ancestor of a visible transcript timestamp is the
    // strongest layout-independent signal. It works both for the right-side panel
    // and for the responsive layout where Transcript moves below the video.
    for (const el of document.querySelectorAll('body *')) {
      if (!isRendered(el)) continue;
      const text = normalizeLine(el.innerText || el.textContent || '');
      if (!text || text.length > 240) continue;

      if (exactClockRegex.test(text)) {
        addVote(nearestScrollableAncestor(el), 3, 'clock');
        continue;
      }

      if (spokenTimeRegex.test(text) && /[A-Za-zА-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі]/u.test(text)) {
        addVote(nearestScrollableAncestor(el), 2, 'spoken-time');
      }
    }
    return votes;
  }

  function horizontalOverlap(a, b) {
    const overlap = Math.max(0, Math.min(a.right, b.right) - Math.max(a.left, b.left));
    return overlap / Math.max(1, Math.min(a.width, b.width));
  }

  function findTranscriptScroller() {
    const candidates = [];
    const votes = transcriptEntryVotes();
    const headingRects = getTranscriptHeadingRects();

    for (const el of document.querySelectorAll('body *')) {
      if (!isScrollable(el)) continue;
      const r = el.getBoundingClientRect();
      if (r.width < 160 || r.height < 80) continue;

      const text = normalize(el.innerText || el.textContent || '');
      if (text.length < 20) continue;

      let score = 0;
      const reasons = [];
      const sample = text.slice(0, 24000);
      const times = (sample.match(clockTimeRegex) || []).length;
      const idClass = `${el.id || ''} ${typeof el.className === 'string' ? el.className : ''} ${el.getAttribute('aria-label') || ''}`;
      const ancestorText = normalize(el.parentElement?.innerText || '').slice(0, 3200);
      const vote = votes.get(el);
      const hasAiWarning = /AI-generated content may be incorrect/i.test(sample);
      const hasTranscriptA11y = /Transcript\. Use arrow keys to navigate between transcript entries/i.test(sample);
      const hasTranscriptText = transcriptWord.test(sample.slice(0, 3000));
      const hasStrongTranscriptSignal = hasAiWarning || hasTranscriptA11y || (hasTranscriptText && times >= 1);

      if (vote?.count) {
        score += vote.count * 65;
        reasons.push(`entry-votes:${vote.count}`);
      }
      if (transcriptWord.test(idClass)) { score += 150; reasons.push('transcript-id/class'); }
      if (hasTranscriptA11y) {
        score += 260; reasons.push('transcript-a11y');
      }
      if (hasAiWarning) {
        score += 150; reasons.push('ai-warning');
      }
      if (hasTranscriptText) { score += 90; reasons.push('transcript-text'); }
      if (transcriptWord.test(ancestorText)) { score += 55; reasons.push('transcript-parent'); }
      if (times >= 1) { score += Math.min(160, times * 14); reasons.push(`times:${times}`); }

      // Semantic proximity to the Transcript tab/heading. Geometry is only used
      // for proximity, not for deciding whether Transcript must be on the right.
      for (const h of headingRects) {
        const dy = r.top - h.bottom;
        if (dy >= -80 && dy <= 900 && horizontalOverlap(r, h) >= 0.18) {
          const bonus = dy <= 300 ? 170 : 100;
          score += bonus;
          reasons.push('near-transcript-heading');
          break;
        }
      }

      // A responsive main-content scroller can legitimately contain the video
      // and the Transcript section at the same time. Penalize it only mildly
      // when transcript semantics are already present.
      if (el.querySelector('video')) {
        score -= hasStrongTranscriptSignal ? 35 : 260;
        reasons.push('contains-video');
      }
      if (/\bRecord\b[\s\S]{0,200}\bUpload\b[\s\S]{0,200}\bFavorite\b/i.test(sample.slice(0, 2500))) {
        score -= hasStrongTranscriptSignal ? 25 : 140;
        reasons.push('page-shell');
      }
      if (!isInViewport(el)) {
        score -= 10;
        reasons.push('offscreen');
      }

      if (el.scrollHeight > el.clientHeight * 1.25) score += 20;
      if (el.scrollHeight > el.clientHeight * 2.0) score += 15;

      candidates.push({
        el, score, times, votes: vote?.count || 0, reasons,
        strong: hasStrongTranscriptSignal,
        rect: { x: r.x, y: r.y, width: r.width, height: r.height },
        scroll: { clientHeight: el.clientHeight, scrollHeight: el.scrollHeight, scrollTop: el.scrollTop },
        overflowY: getComputedStyle(el).overflowY,
        idClass
      });
    }

    candidates.sort((a, b) =>
      Number(b.strong) - Number(a.strong) ||
      b.score - a.score ||
      b.votes - a.votes ||
      b.times - a.times
    );
    return candidates;
  }
  const durationOnlyRegex = /^(?:(\d+)\s+hours?\s*)?(?:(\d+)\s+minutes?\s*)?(?:(\d+)\s+seconds?)$/i;
  const clockOnlyRegex = /^(?:\d{1,2}:)?\d{1,2}:\d{2}$/;
  const initialsRegex = /^[A-ZА-ЯЁӘҒҚҢӨҰҮҺІ]{1,3}$/u;

  function parseDurationText(value) {
    const line = normalizeLine(value).toLowerCase();
    const match = line.match(durationOnlyRegex);
    if (!match) return null;
    const hours = Number(match[1] || 0);
    const minutes = Number(match[2] || 0);
    const seconds = Number(match[3] || 0);
    if (!hours && !minutes && !seconds && !/0\s+seconds?/.test(line)) return null;
    return hours * 3600 + minutes * 60 + seconds;
  }

  function secondsToClock(totalSeconds) {
    totalSeconds = Math.max(0, Number(totalSeconds) || 0);
    const h = Math.floor(totalSeconds / 3600);
    const m = Math.floor((totalSeconds % 3600) / 60);
    const s = totalSeconds % 60;
    return h > 0 ? `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}` : `${m}:${String(s).padStart(2, '0')}`;
  }

  function parseClock(value) {
    const parts = String(value || '').trim().split(':').map(Number);
    if (parts.some(Number.isNaN)) return null;
    if (parts.length === 2) return parts[0] * 60 + parts[1];
    if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2];
    return null;
  }

  function getMediaDurationSeconds() {
    const video = document.querySelector('video');
    if (video && Number.isFinite(video.duration) && video.duration > 1) return Math.round(video.duration);
    const body = normalize(document.body.innerText || '');
    const matches = [...body.matchAll(/(?:^|\s)((?:\d{1,2}:)?\d{1,2}:\d{2})\s*\/\s*((?:\d{1,2}:)?\d{1,2}:\d{2})(?:\s|$)/g)];
    for (const m of matches) {
      const total = parseClock(m[2]);
      if (total && total > 10) return total;
    }
    return null;
  }

  function parseCombinedHeader(line) {
    const normalized = normalizeLine(line);
    if (!normalized) return null;
    const match = normalized.match(/^(.+?)\s+((?:(?:\d+)\s+hours?\s*)?(?:(?:\d+)\s+minutes?\s*)?(?:(?:\d+)\s+seconds?))$/i);
    if (!match) return null;

    const speaker = normalizeLine(match[1]);
    const durationText = normalizeLine(match[2]);
    const seconds = parseDurationText(durationText);
    if (seconds === null) return null;
    if (speaker.length < 2 || speaker.length > 180) return null;
    if (!/[A-Za-zА-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі]/u.test(speaker)) return null;
    if (/^(transcript|download|search|speakers?)$/i.test(speaker)) return null;
    return { speaker, seconds, durationText };
  }

  function isNoiseLine(line) {
    const value = normalizeLine(line);
    if (!value) return true;
    if (/^AI-generated content may be incorrect\.?$/i.test(value)) return true;
    if (/^Is this transcript useful\??$/i.test(value)) return true;
    if (/^(Transcript|Transcription|Search|Download|Speakers)$/i.test(value)) return true;
    if (/^Transcript\. Use arrow keys/i.test(value)) return true;
    if (/^You don't have permission to download the transcript/i.test(value)) return true;
    if (clockOnlyRegex.test(value)) return true;
    if (parseDurationText(value) !== null) return true;
    if (initialsRegex.test(value)) return true;
    return false;
  }

  function parseTranscriptText(rawText) {
    const lines = (rawText || '').split(/\r?\n/).map(normalizeLine).filter(Boolean);
    const headers = [];
    for (let i = 0; i < lines.length; i++) {
      const parsed = parseCombinedHeader(lines[i]);
      if (parsed) headers.push({ index: i, ...parsed });
    }

    const entries = [];
    if (headers.length < 1) return { entries, headerCount: headers.length, rawLines: lines };

    for (let h = 0; h < headers.length; h++) {
      const current = headers[h];
      const next = headers[h + 1] || null;
      const end = next ? next.index : lines.length;
      let segment = lines.slice(current.index + 1, end);

      if (next) {
        while (segment.length) {
          const tail = normalizeLine(segment[segment.length - 1]);
          if (!tail || tail === next.speaker || isNoiseLine(tail)) segment.pop();
          else break;
        }
      }

      const textLines = [];
      for (const line of segment) {
        if (!line || line === current.speaker || isNoiseLine(line) || parseCombinedHeader(line)) continue;
        textLines.push(line);
      }
      const text = normalize(textLines.join('\n'));
      if (text) entries.push({ speaker: current.speaker, seconds: current.seconds, text });
    }
    return { entries, headerCount: headers.length, rawLines: lines };
  }

  function mergeEntries(store, newEntries, seqRef) {
    let added = 0;
    for (const entry of newEntries) {
      const speakerKey = entry.speaker.toLowerCase().replace(/\s+/g, ' ').trim();
      const textKey = entry.text.toLowerCase().replace(/\s+/g, ' ').trim();
      const base = `${entry.seconds}|${speakerKey}`;

      let replaced = false;
      for (const [key, existing] of store) {
        if (existing.base !== base) continue;
        const oldText = existing.entry.text.toLowerCase().replace(/\s+/g, ' ').trim();
        if (oldText === textKey) { replaced = true; break; }
        if (textKey.includes(oldText) && textKey.length > oldText.length) {
          store.delete(key);
          const newKey = `${base}|${textKey}`;
          store.set(newKey, { base, seq: existing.seq, entry });
          replaced = true;
          added++;
          break;
        }
        if (oldText.includes(textKey)) { replaced = true; break; }
      }
      if (replaced) continue;

      const key = `${base}|${textKey}`;
      if (!store.has(key)) {
        store.set(key, { base, seq: seqRef.value++, entry });
        added++;
      }
    }
    return added;
  }

  function formatEntries(store) {
    const rows = Array.from(store.values())
      .sort((a, b) => a.entry.seconds - b.entry.seconds || a.seq - b.seq)
      .map(x => x.entry);
    const text = rows.map(e => `${secondsToClock(e.seconds)}\n${e.speaker}\n${e.text}`).join('\n\n').trim();
    return { rows, text };
  }

  function hashText(text) {
    let h = 2166136261;
    for (let i = 0; i < text.length; i++) {
      h ^= text.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    return (h >>> 0).toString(16);
  }

  async function extractTranscript() {
    cancelRequested = false;
    update({
      status: 'running', message: 'Ищу панель Transcript...', progress: 0,
      items: 0, chars: 0, text: '', error: '', startedAt: Date.now(), finishedAt: 0
    });

    const run = {
      version: VERSION,
      iterations: 0,
      bottomAttempts: 0,
      backwardSnaps: 0,
      maxObservedScrollHeight: 0,
      maxTargetTop: 0,
      maxSeenSeconds: 0,
      mediaDurationSeconds: null,
      snapshots: 0,
      parsedHeaders: 0,
      events: []
    };
    lastRunDebug = run;

    try {
      let candidates = [];
      let best = null;
      const discoveryStarted = Date.now();
      while (Date.now() - discoveryStarted < 8000) {
        candidates = findTranscriptScroller();
        best = candidates[0] || null;
        if (best && (best.strong || best.score >= 45)) break;
        await sleep(250);
      }
      if (!best || (!best.strong && best.score < 45)) {
        run.events.push({
          event: 'transcript-scroller-not-found',
          viewport: { width: innerWidth, height: innerHeight },
          candidates: candidates.slice(0, 8).map(x => ({
            score: x.score,
            strong: x.strong,
            times: x.times,
            votes: x.votes,
            reasons: x.reasons,
            rect: x.rect,
            scroll: x.scroll,
            overflowY: x.overflowY,
            idClass: x.idClass
          }))
        });
        throw new Error('Не удалось уверенно определить область прокрутки Transcript. Запусти диагностику.');
      }

      let scroller = best.el;
      run.events.push({
        event: 'transcript-scroller-selected',
        viewport: { width: innerWidth, height: innerHeight },
        score: best.score,
        strong: best.strong,
        times: best.times,
        votes: best.votes,
        reasons: best.reasons,
        rect: best.rect,
        scroll: best.scroll,
        overflowY: best.overflowY,
        idClass: best.idClass
      });
      const title = sanitizeFileName(normalize(document.querySelector('h1')?.innerText || document.title || 'teams-transcript'));
      const mediaDuration = getMediaDurationSeconds();
      run.mediaDurationSeconds = mediaDuration;

      const store = new Map();
      const seqRef = { value: 0 };
      const rawSnapshots = new Map();

      function capture(label) {
        const raw = normalize(scroller.innerText || scroller.textContent || '');
        if (raw) {
          const hash = hashText(raw);
          if (!rawSnapshots.has(hash)) rawSnapshots.set(hash, raw);
        }
        const parsed = parseTranscriptText(raw);
        run.snapshots++;
        run.parsedHeaders += parsed.headerCount;
        const added = mergeEntries(store, parsed.entries, seqRef);
        for (const e of parsed.entries) run.maxSeenSeconds = Math.max(run.maxSeenSeconds, e.seconds);
        const formatted = formatEntries(store);
        update({
          items: formatted.rows.length,
          chars: formatted.text.length,
          text: formatted.text,
          message: `${label}. Собрано реплик: ${formatted.rows.length}.`
        });
        return { added, count: formatted.rows.length, rawLength: raw.length };
      }

      update({ title, message: 'Панель найдена. Перехожу к началу и выполняю один проход вперед...' });
      scroller.scrollTop = 0;
      scroller.dispatchEvent(new Event('scroll', { bubbles: true }));
      await sleep(500);
      capture('Начало транскрипции');

      let logicalTop = 0;
      let peakProgress = 0;
      let stableBottomChecks = 0;
      let previousCount = store.size;
      let lastHeight = scroller.scrollHeight;

      // One forward scan only. logicalTop never decreases, even if the virtualized
      // SharePoint list snaps its physical scrollTop back to the beginning.
      for (let iteration = 0; iteration < 240; iteration++) {
        run.iterations = iteration + 1;
        if (cancelRequested) {
          update({ status: 'cancelled', message: 'Сбор остановлен пользователем.', finishedAt: Date.now() });
          return;
        }

        if (!scroller.isConnected || scroller.clientHeight < 40) {
          const refreshed = findTranscriptScroller()[0];
          if (!refreshed || (!refreshed.strong && refreshed.score < 45)) {
            throw new Error('Область Transcript была перестроена страницей и не найдена повторно. Запусти диагностику.');
          }
          scroller = refreshed.el;
          run.events.push({ i: iteration + 1, event: 'scroller-reselected', score: refreshed.score, votes: refreshed.votes || 0, reasons: refreshed.reasons || [] });
        }

        const clientHeight = Math.max(1, scroller.clientHeight);
        const scrollHeight = Math.max(clientHeight, scroller.scrollHeight);
        const maxTop = Math.max(0, scrollHeight - clientHeight);
        run.maxObservedScrollHeight = Math.max(run.maxObservedScrollHeight, scrollHeight);

        // Monotonic cursor. Never derive the next position from a backwards snap.
        logicalTop = Math.min(Math.max(logicalTop, 0), maxTop);
        const step = Math.max(140, Math.floor(clientHeight * 0.48));
        const remaining = maxTop - logicalTop;

        // Before the exact bottom, capture an overlapping near-bottom viewport.
        let targetTop;
        let finalAttempt = false;
        if (remaining <= step) {
          const preBottom = Math.max(logicalTop, maxTop - Math.max(60, Math.floor(clientHeight * 0.18)));
          if (preBottom > logicalTop + 5) {
            targetTop = preBottom;
          } else {
            targetTop = maxTop;
            finalAttempt = true;
          }
        } else {
          targetTop = logicalTop + step;
        }

        targetTop = Math.max(logicalTop, Math.min(maxTop, targetTop));
        run.maxTargetTop = Math.max(run.maxTargetTop, targetTop);

        const beforeHeight = scrollHeight;
        const beforeCount = store.size;
        const beforeActual = scroller.scrollTop;

        scroller.scrollTop = targetTop;
        scroller.dispatchEvent(new Event('scroll', { bubbles: true }));
        await sleep(finalAttempt ? 420 : 240);

        const cap = capture(finalAttempt ? 'Проверяю конец транскрипции' : 'Читаю следующую порцию');
        const actualTop = Math.max(0, scroller.scrollTop);
        const afterHeight = Math.max(scroller.clientHeight, scroller.scrollHeight);
        const afterMaxTop = Math.max(0, afterHeight - scroller.clientHeight);
        const backwards = actualTop + Math.max(40, clientHeight * 0.2) < Math.min(targetTop, afterMaxTop);
        if (backwards) run.backwardSnaps++;

        // logicalTop follows our requested target, not the physical scroll position.
        logicalTop = Math.max(logicalTop, targetTop);

        const denom = Math.max(1, maxTop);
        const progress = Math.max(0, Math.min(100, Math.floor((logicalTop / denom) * 100)));
        peakProgress = Math.max(peakProgress, progress);
        update({ progress: peakProgress });

        const heightGrew = afterHeight > beforeHeight + Math.max(40, Math.floor(clientHeight * 0.08));
        const newItems = store.size - beforeCount;

        run.events.push({
          i: iteration + 1,
          targetTop: Math.round(targetTop),
          actualTop: Math.round(actualTop),
          beforeActual: Math.round(beforeActual),
          beforeHeight: Math.round(beforeHeight),
          afterHeight: Math.round(afterHeight),
          maxTop: Math.round(maxTop),
          added: cap.added,
          total: store.size,
          finalAttempt,
          backwards,
          heightGrew
        });
        if (run.events.length > 60) run.events.shift();

        // Strong completion signal: transcript timestamps are already near the
        // recording duration and the last move produced nothing new.
        const nearMediaEnd = mediaDuration && run.maxSeenSeconds >= Math.max(0, mediaDuration - 75);
        if (nearMediaEnd && newItems === 0 && peakProgress >= 92) {
          stableBottomChecks++;
        }

        if (finalAttempt) {
          run.bottomAttempts++;
          await sleep(550);
          const settle = capture('Финализирую последний lazy-loaded фрагмент');
          const settledHeight = Math.max(scroller.clientHeight, scroller.scrollHeight);
          const settledGrew = settledHeight > afterHeight + Math.max(40, Math.floor(clientHeight * 0.08));
          const settledNew = store.size - beforeCount;

          if (settledGrew && run.bottomAttempts < 6) {
            // Continue only into newly appended space. Never go back to 0.
            logicalTop = Math.min(targetTop, Math.max(0, settledHeight - scroller.clientHeight));
            lastHeight = settledHeight;
            stableBottomChecks = 0;
            continue;
          }

          // We deliberately attempted the physical end and the list did not grow.
          // A snap back to the top is therefore a completion signal, not a new pass.
          if (!settledGrew) break;
        }

        // If SharePoint snaps backwards late in the scan, do not follow it.
        // Make the next iteration continue from the monotonic logical cursor.
        if (backwards && peakProgress >= 85 && !heightGrew) {
          stableBottomChecks++;
          if (stableBottomChecks >= 2) {
            // One exact end attempt on the next loop, but never a second pass.
            logicalTop = Math.max(logicalTop, maxTop - Math.max(30, Math.floor(clientHeight * 0.10)));
          }
        } else if (newItems > 0 || heightGrew) {
          stableBottomChecks = 0;
        }

        // Safety stop for a virtualized control that changes geometry but no longer
        // yields new transcript entries. This cannot restart from the beginning.
        if (peakProgress >= 98 && store.size === previousCount && Math.abs(afterHeight - lastHeight) < 40) {
          stableBottomChecks++;
          if (stableBottomChecks >= 3) break;
        }

        previousCount = store.size;
        lastHeight = afterHeight;
      }

      capture('Сбор завершен');
      let { rows, text } = formatEntries(store);

      if (!text || rows.length < 3) {
        // Parser fallback: return unique lazy-loaded snapshots rather than looping.
        const chunks = Array.from(rawSnapshots.values()).filter(Boolean);
        text = chunks.join('\n\n---\n\n').trim();
        rows = [];
      }
      if (!text || text.length < 20) throw new Error('Панель найдена, но текст транскрипции извлечь не удалось. Запусти диагностику.');

      update({
        status: 'done',
        message: rows.length
          ? `Готово. Один проход завершен. Собрано реплик: ${rows.length}.`
          : 'Готово. Один проход завершен, возвращены собранные DOM-фрагменты.',
        progress: 100,
        items: rows.length || rawSnapshots.size,
        chars: text.length,
        text,
        error: '',
        finishedAt: Date.now()
      });
    } catch (e) {
      update({
        status: 'error',
        message: `Ошибка: ${e?.message || e}`,
        error: e?.message || String(e),
        finishedAt: Date.now()
      });
    } finally {
      runningPromise = null;
    }
  }

  function probeTranscript() {
    const candidates = findTranscriptScroller();
    const best = candidates[0] || null;
    const bodyText = normalize(document.body?.innerText || document.body?.textContent || '');
    const transcriptVisible = /AI-generated content may be incorrect|Transcript\. Use arrow keys|\bTranscript\b/i.test(bodyText);

    return {
      ok: true,
      url: location.href,
      title: document.title,
      isTop: window === window.top,
      viewport: { width: innerWidth, height: innerHeight },
      transcriptVisible,
      bodyChars: bodyText.length,
      best: best ? {
        score: best.score,
        strong: !!best.strong,
        times: best.times || 0,
        votes: best.votes || 0,
        reasons: best.reasons || [],
        rect: best.rect || null,
        scroll: best.scroll || {
          clientHeight: best.el?.clientHeight || 0,
          scrollHeight: best.el?.scrollHeight || 0,
          scrollTop: best.el?.scrollTop || 0
        },
        overflowY: best.overflowY || getComputedStyle(best.el).overflowY,
        idClass: best.idClass || ''
      } : null
    };
  }

  function buildDiagnostic() {
    const candidates = findTranscriptScroller();
    const visibleText = normalize(document.body.innerText || '').slice(0, 12000);
    const lines = [
      'Teams Recap Transcript Exporter DOM diagnostic',
      `Version: ${VERSION}`,
      `URL: ${location.href}`,
      `Title: ${document.title}`,
      `Frame: ${window === window.top ? 'top' : 'child'}`,
      `Viewport: ${innerWidth}x${innerHeight}`,
      `State: ${JSON.stringify({ status: state.status, progress: state.progress, items: state.items, chars: state.chars })}`,
      `LastRun: ${JSON.stringify(lastRunDebug)}`,
      '',
      'Top scrollable candidates:'
    ];

    candidates.slice(0, 25).forEach((c, i) => {
      const preview = normalize(c.el.innerText || '').slice(0, 1200).replace(/\n/g, ' | ');
      lines.push(`[${i + 1}] score=${c.score} votes=${c.votes || 0} times=${c.times} reasons=${(c.reasons || []).join(',')} rect=${JSON.stringify(c.rect)} scrollTop=${c.el.scrollTop} clientHeight=${c.el.clientHeight} scrollHeight=${c.el.scrollHeight} id/class=${c.idClass}`);
      lines.push(`    ${preview}`);
    });

    lines.push('', 'Visible page text sample:', visibleText);
    return lines.join('\n');
  }

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    const type = message?.type;

    if (type === 'PROBE_TRANSCRIPT') {
      sendResponse(probeTranscript());
      return;
    }
    if (type === 'GET_STATE') {
      sendResponse({ ok: true, state: publicState() });
      return;
    }
    if (type === 'START_EXTRACT') {
      if (!runningPromise && state.status !== 'running') runningPromise = extractTranscript();
      sendResponse({ ok: true, state: publicState() });
      return;
    }
    if (type === 'STOP_EXTRACT') {
      cancelRequested = true;
      sendResponse({ ok: true, state: publicState() });
      return;
    }
    if (type === 'CLEAR_RESULT') {
      if (state.status !== 'running') {
        lastRunDebug = {};
        update({
          status: 'idle', message: 'Готово к работе.', progress: 0, items: 0,
          chars: 0, text: '', error: '', title: 'teams-transcript', startedAt: 0, finishedAt: 0
        });
      }
      sendResponse({ ok: true, state: publicState() });
      return;
    }
    if (type === 'GET_DEBUG') {
      sendResponse({ ok: true, text: buildDiagnostic() });
      return;
    }
  });
})();
