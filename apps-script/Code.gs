/**
 * Живые справки — скрипт для таблицы «список на РАЗНОС».
 *
 * Что делает:
 *  1. Раз в минуту находит новые треки на вкладке «На разнос» и готовит для них карточки:
 *     метаданные со страницы по ссылке + факты через OpenRouter (модель с веб-поиском).
 *     Карточки хранятся на вкладке «Справки (авто)» — её можно смотреть и править руками.
 *  2. Отдаёт пульту текущий трек («В процессе») и его карточку через Web App.
 *
 * Ключ OpenRouter хранится в свойствах скрипта (OPENROUTER_KEY), в коде его НЕТ.
 */

const CFG = {
  SHEET: 'На разнос',            // вкладка с очередью
  CARDS_SHEET: 'Справки (авто)', // сюда пишутся готовые карточки
  FIRST_ROW: 9,                  // первая строка с треками
  COL: { link: 1, who: 2, type: 3, price: 4, rating: 5, from: 6, genre: 8, extra: 10 },
  STATUS_COLS: [1, 2, 3, 4, 5],  // в каких колонках искать цвет отметки
  CURRENT_LABEL: 'В процессе',   // текст ячейки-легенды с цветом «в процессе»
  NEXT_LABEL: 'След трек',       // текст ячейки-легенды с цветом «следующий»
  FALLBACK_CURRENT: '#ffff00',
  FALLBACK_NEXT: '#00ff00',
  SKIP_TEXT: ['ЧТО-ТО', 'ЧТО ТО', ''],
  MODEL: 'google/gemini-2.5-flash:online', // можно сменить в свойствах скрипта: MODEL
  PER_RUN: 3,                    // сколько треков обрабатывать за один запуск триггера
};

// ───────────────────────── меню и установка ─────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Живые справки')
    .addItem('1. Первоначальная настройка', 'setup')
    .addItem('Сгенерировать для выделенной строки', 'generateSelected')
    .addItem('Обработать новые треки сейчас', 'processPending')
    .addToUi();
}

function setup() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('OPENROUTER_KEY')) {
    const r = ui.prompt('Ключ OpenRouter', 'Вставь API-ключ OpenRouter (sk-or-...). Он сохранится в свойствах скрипта и не будет виден в коде.', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    props.setProperty('OPENROUTER_KEY', r.getResponseText().trim());
  }
  if (!props.getProperty('GEN_TOKEN')) props.setProperty('GEN_TOKEN', Utilities.getUuid().slice(0, 8));
  cardsSheet_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'processPending').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processPending').timeBased().everyMinutes(1).create();
  ui.alert('Готово',
    'Триггер включён: раз в минуту новые треки получают карточки.\n\n' +
    'Токен для пульта (поле «Токен генерации»): ' + props.getProperty('GEN_TOKEN') + '\n\n' +
    'Дальше: Развернуть → Новое развёртывание → Веб-приложение (доступ: «Все») и вставь ссылку в пульт.',
    ui.ButtonSet.OK);
}

// ───────────────────────── Web App (для пульта) ─────────────────────────
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    let out;
    if (p.action === 'list') out = { rows: listRows_().map(r => ({ row: r.row, title: r.title, who: r.who, status: r.status, hasCard: !!getCard_(r.key) })) };
    else if (p.action === 'row') out = { track: withCard_(findRow_(Number(p.row))) };
    else if (p.action === 'generate') {
      if (!p.token || p.token !== PropertiesService.getScriptProperties().getProperty('GEN_TOKEN')) throw new Error('неверный токен');
      const r = findRow_(Number(p.row)); if (!r) throw new Error('строка не найдена');
      out = { card: generateFor_(r, true) };
    } else {
      const cache = CacheService.getScriptCache(); const hit = cache.get('state');
      if (hit) return json_(JSON.parse(hit));
      const rows = listRows_();
      out = { current: withCard_(rows.find(r => r.status === 'current')), next: withCard_(rows.find(r => r.status === 'next')), ok: true };
      cache.put('state', JSON.stringify(out), 3);
    }
    out.ok = true;
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function withCard_(r) { if (!r) return null; return Object.assign({}, r, { card: getCard_(r.key) }); }

// ───────────────────────── чтение очереди ─────────────────────────
function legendColor_(sh, label, fallback) {
  const cell = sh.createTextFinder(label).matchEntireCell(false).findNext();
  return (cell ? cell.getBackground() : fallback).toLowerCase();
}
function listRows_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET);
  if (!sh) throw new Error('нет вкладки «' + CFG.SHEET + '»');
  const last = sh.getLastRow(); if (last < CFG.FIRST_ROW) return [];
  const n = last - CFG.FIRST_ROW + 1;
  const width = Math.max(CFG.COL.extra, Math.max.apply(null, CFG.STATUS_COLS));
  const rng = sh.getRange(CFG.FIRST_ROW, 1, n, width);
  const vals = rng.getDisplayValues(), bgs = rng.getBackgrounds();
  const rich = sh.getRange(CFG.FIRST_ROW, CFG.COL.link, n, 1).getRichTextValues();
  const cCur = legendColor_(sh, CFG.CURRENT_LABEL, CFG.FALLBACK_CURRENT);
  const cNext = legendColor_(sh, CFG.NEXT_LABEL, CFG.FALLBACK_NEXT);
  const rows = [];
  for (let i = 0; i < n; i++) {
    const v = vals[i], text = String(v[CFG.COL.link - 1] || '').trim();
    if (CFG.SKIP_TEXT.indexOf(text.toUpperCase()) >= 0) continue;
    const link = linkOf_(rich[i][0], text);
    const colors = CFG.STATUS_COLS.map(c => String(bgs[i][c - 1]).toLowerCase());
    const status = colors.indexOf(cCur) >= 0 ? 'current' : colors.indexOf(cNext) >= 0 ? 'next' : '';
    rows.push({
      row: CFG.FIRST_ROW + i, title: text, link: link, key: link || text,
      who: v[CFG.COL.who - 1], type: v[CFG.COL.type - 1], genre: v[CFG.COL.genre - 1], extra: v[CFG.COL.extra - 1],
      status: status,
    });
  }
  return rows;
}
function findRow_(row) { return listRows_().filter(r => r.row === row)[0] || null; }
function linkOf_(rt, text) {
  if (rt) {
    const direct = rt.getLinkUrl(); if (direct) return direct;
    const runs = rt.getRuns(); for (let i = 0; i < runs.length; i++) { const u = runs[i].getLinkUrl(); if (u) return u; }
  }
  const m = String(text).match(/https?:\/\/\S+/); return m ? m[0] : '';
}

// ───────────────────────── хранилище карточек ─────────────────────────
function cardsSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.CARDS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CFG.CARDS_SHEET);
    sh.appendRow(['ключ (ссылка)', 'строка', 'статус', 'карточка (JSON)', 'обновлено']);
    sh.setFrozenRows(1); sh.setColumnWidth(1, 320); sh.setColumnWidth(4, 600);
  }
  return sh;
}
function cardIndex_() {
  const sh = cardsSheet_(); const last = sh.getLastRow();
  const map = {}; if (last < 2) return map;
  sh.getRange(2, 1, last - 1, 4).getValues().forEach((v, i) => { map[v[0]] = { rowIdx: i + 2, status: v[2], json: v[3] }; });
  return map;
}
function getCard_(key) {
  const c = cardIndex_()[key]; if (!c || !c.json) return c ? { status: c.status } : null;
  try { const o = JSON.parse(c.json); o.status = c.status; return o; } catch (e) { return { status: 'error', error: 'битый JSON' }; }
}
function saveCard_(key, row, status, obj) {
  const sh = cardsSheet_(); const idx = cardIndex_()[key];
  const data = [key, row, status, obj ? JSON.stringify(obj) : '', new Date()];
  if (idx) sh.getRange(idx.rowIdx, 1, 1, 5).setValues([data]); else sh.appendRow(data);
}

// ───────────────────────── генерация ─────────────────────────
function processPending() {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(1000)) return;
  try {
    const idx = cardIndex_(); let done = 0;
    const rows = listRows_();
    // сначала — текущий и следующий, потом остальные сверху вниз
    rows.sort((a, b) => (b.status ? 1 : 0) - (a.status ? 1 : 0));
    for (const r of rows) {
      if (done >= CFG.PER_RUN) break;
      const c = idx[r.key];
      if (c && (c.status === 'ready' || c.status === 'unknown')) continue;
      if (c && c.status === 'error' && !r.status) continue; // ошибки не повторяем бесконечно
      generateFor_(r, false); done++;
    }
  } finally { lock.releaseLock(); }
}
function generateSelected() {
  const sh = SpreadsheetApp.getActiveSheet(); const row = sh.getActiveRange().getRow();
  const r = findRow_(row); if (!r) { SpreadsheetApp.getUi().alert('Выдели строку с треком на вкладке «' + CFG.SHEET + '»'); return; }
  const card = generateFor_(r, true);
  SpreadsheetApp.getUi().alert(card.status === 'ready' ? 'Готово: ' + (card.facts || []).length + ' факт(а)' : 'Статус: ' + card.status + (card.error ? '\n' + card.error : ''));
}

function generateFor_(r, force) {
  saveCard_(r.key, r.row, 'pending', null);
  let card;
  try {
    const meta = pageMeta_(r.link);
    const ai = askAI_(r, meta);
    card = {
      known: ai.known !== false,
      meta: {
        title: ai.title || meta.title || r.title, artist: ai.artist || meta.artist || '',
        album: ai.album || '', year: ai.year ? String(ai.year) : '', cover: meta.image || '',
      },
      artistInfo: ai.artistInfo || '',
      facts: (ai.facts || []).filter(f => f && f.text).slice(0, 4),
    };
    card.status = card.known && card.facts.length ? 'ready' : 'unknown';
  } catch (e) {
    card = { status: 'error', error: String(e.message || e) };
  }
  saveCard_(r.key, r.row, card.status, card);
  CacheService.getScriptCache().remove('state');
  return card;
}

// Метаданные со страницы: og-теги (Spotify, YouTube, Яндекс Музыка, SoundCloud…)
function pageMeta_(url) {
  const meta = { title: '', artist: '', image: '', description: '' };
  if (!url || !/^https?:/.test(url)) return meta;
  try {
    if (/youtu\.?be/.test(url)) {
      const o = JSON.parse(UrlFetchApp.fetch('https://www.youtube.com/oembed?format=json&url=' + encodeURIComponent(url), { muteHttpExceptions: true }).getContentText());
      meta.title = o.title || ''; meta.artist = o.author_name || ''; meta.image = o.thumbnail_url || '';
      return meta;
    }
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'ru,en' } });
    if (res.getResponseCode() >= 400) return meta;
    const html = res.getContentText().slice(0, 300000);
    const og = (p) => { const m = html.match(new RegExp('<meta[^>]+property=["\']og:' + p + '["\'][^>]+content=["\']([^"\']*)', 'i')) || html.match(new RegExp('<meta[^>]+content=["\']([^"\']*)["\'][^>]+property=["\']og:' + p, 'i')); return m ? decode_(m[1]) : ''; };
    meta.title = og('title'); meta.image = og('image'); meta.description = og('description');
    // Spotify: «Meaningful Stone · A Call from My Dream · Song · 2020»
    if (/spotify\.com/.test(url) && meta.description) meta.artist = meta.description.split('·')[0].trim();
  } catch (e) {}
  return meta;
}
function decode_(s) { return s.replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;|&#x27;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>'); }

function askAI_(r, meta) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('OPENROUTER_KEY'); if (!key) throw new Error('нет OPENROUTER_KEY в свойствах скрипта');
  const model = props.getProperty('MODEL') || CFG.MODEL;
  const info = [
    'Ссылка: ' + (r.link || '—'),
    'Текст в таблице: ' + r.title,
    meta.title ? 'Название со страницы: ' + meta.title : '',
    meta.artist ? 'Исполнитель со страницы: ' + meta.artist : '',
    meta.description ? 'Описание страницы: ' + meta.description : '',
    r.genre ? 'Жанр (по таблице): ' + r.genre : '',
  ].filter(Boolean).join('\n');
  const prompt =
    'Ты готовишь короткие карточки-справки для музыкального стрима, где композитор разбирает треки зрителей.\n' +
    'Найди в интернете информацию о треке ниже и верни ТОЛЬКО JSON без пояснений:\n' +
    '{"known": true|false, "title": "", "artist": "", "album": "", "year": "", ' +
    '"artistInfo": "1–2 предложения об артисте, до 200 символов", ' +
    '"facts": [{"title": "заголовок до 40 символов", "text": "факт до 170 символов", "source": "URL источника"}]}\n\n' +
    'Правила:\n' +
    '- Пиши по-русски, живо и просто.\n' +
    '- Только факты, подтверждённые найденными источниками; у каждого факта — URL источника из результатов поиска.\n' +
    '- 2–4 факта: история создания, смысл текста, награды и чарты, сэмплы, коллаборации, необычные детали записи.\n' +
    '- Не пересказывай очевидное (название, длительность).\n' +
    '- Если трек найти не удалось или это малоизвестный авторский трек без информации в сети — верни "known": false и пустой facts. НИЧЕГО не выдумывай.\n\n' +
    info;
  const res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + key, 'HTTP-Referer': 'https://laitsberg.github.io', 'X-Title': 'Live Spravki' },
    payload: JSON.stringify({ model: model, temperature: 0.2, messages: [{ role: 'user', content: prompt }] }),
  });
  const code = res.getResponseCode(), body = res.getContentText();
  if (code >= 400) throw new Error('OpenRouter ' + code + ': ' + body.slice(0, 200));
  const txt = JSON.parse(body).choices[0].message.content || '';
  const m = txt.match(/\{[\s\S]*\}/); if (!m) throw new Error('модель не вернула JSON');
  return JSON.parse(m[0]);
}
