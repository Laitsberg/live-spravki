/**
 * Живые справки — отдельный скрипт (живёт в ТВОЁМ Google-аккаунте, а не в таблице композитора).
 *
 *  • Читает таблицу «список на РАЗНОС» (нужен только доступ на просмотр) — саму таблицу НЕ меняет.
 *  • Раз в минуту находит новые треки на вкладке «На разнос» и готовит карточки:
 *    метаданные по ссылке + факты через OpenRouter (модель с веб-поиском).
 *    Карточки лежат в твоей отдельной таблице «Живые справки — карточки» (её можно править руками).
 *  • Отдаёт пульту текущий трек («В процессе») и его карточку через Web App.
 *
 * Ключ OpenRouter — в «Настройки проекта → Свойства скрипта» (OPENROUTER_KEY). В коде его нет,
 * и так как проект твой, никто кроме тебя его не видит.
 */

const CFG = {
  SOURCE_ID: '1yEUr29llt9L1zWavC4lxkyhd6UeIwrQuGsPon9fXt4A', // «список на РАЗНОС»
  SHEET: 'На разнос',
  FIRST_ROW: 3,                  // с какой строки искать треки (служебные строки отсеиваются сами)
  COL: { link: 1, who: 2, type: 3, price: 4, rating: 5, from: 6, tags: 7, genre: 8, feat: 9, extra: 10 },
  STATUS_COLS: [1, 2, 3, 4, 5],  // в каких колонках искать цвет отметки
  CURRENT_LABEL: 'В процессе',   // ячейка-легенда с цветом «в процессе»
  NEXT_LABEL: 'След трек',       // ячейка-легенда с цветом «следующий»
  FALLBACK_CURRENT: '#ffff00',
  FALLBACK_NEXT: '#00ff00',
  PLACEHOLDERS: ['ЧТО-ТО', 'ЧТО ТО', 'ЧТОТО'],
  MODEL: 'google/gemini-2.5-flash:online', // можно переопределить свойством скрипта MODEL
  PER_RUN: 3,                    // сколько треков обрабатывать за один запуск триггера
};

// ───────────────────────── установка ─────────────────────────
/** Запусти ОДИН раз из редактора (▶ setup). Перед этим добавь свойство OPENROUTER_KEY. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('OPENROUTER_KEY')) throw new Error('Сначала добавь свойство скрипта OPENROUTER_KEY (Настройки проекта → Свойства скрипта)');
  if (!props.getProperty('GEN_TOKEN')) props.setProperty('GEN_TOKEN', Utilities.getUuid().slice(0, 8));
  const cards = cardsBook_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'processPending').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processPending').timeBased().everyMinutes(1).create();
  const rows = listRows_();
  Logger.log('Готово. Треков в очереди: ' + rows.length +
    '. Сейчас «В процессе»: ' + ((rows.filter(r => r.status === 'current')[0] || {}).title || 'нет') +
    '\nТаблица карточек: ' + cards.getUrl() +
    '\nТОКЕН ДЛЯ ПУЛЬТА: ' + props.getProperty('GEN_TOKEN'));
}

/** Проверка без триггера: покажет, как скрипт видит очередь. */
function checkQueue() {
  listRows_().slice(0, 40).forEach(r => Logger.log([r.row, r.status || '-', r.title, r.link || '(нет ссылки)', r.who, r.type].join(' | ')));
}

// ───────────────────────── Web App (для пульта) ─────────────────────────
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    let out;
    if (p.action === 'terms') {
      out = { terms: readTerms_() };
    } else if (p.action === 'list') {
      const idx = cardIndex_();
      out = { rows: listRows_().map(r => ({ row: r.row, title: r.title, who: r.who, status: r.status, hasCard: !!(idx[r.key] && idx[r.key].json) })) };
    } else if (p.action === 'row') {
      out = { track: withCard_(findRow_(Number(p.row))) };
    } else if (p.action === 'generate') {
      if (!p.token || p.token !== PropertiesService.getScriptProperties().getProperty('GEN_TOKEN')) throw new Error('неверный токен');
      const r = findRow_(Number(p.row)); if (!r) throw new Error('строка не найдена');
      out = { card: generateFor_(r) };
    } else {
      const cache = CacheService.getScriptCache(); const hit = cache.get('state');
      if (hit) return json_(JSON.parse(hit));
      const rows = listRows_();
      out = { current: withCard_(rows.filter(r => r.status === 'current')[0]), next: withCard_(rows.filter(r => r.status === 'next')[0]), ok: true };
      cache.put('state', JSON.stringify(out), 4);
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
function sourceSheet_() {
  const sh = SpreadsheetApp.openById(CFG.SOURCE_ID).getSheetByName(CFG.SHEET);
  if (!sh) throw new Error('нет вкладки «' + CFG.SHEET + '»');
  return sh;
}
function legendColor_(sh, label, fallback) {
  const cell = sh.createTextFinder(label).matchEntireCell(true).findNext() || sh.createTextFinder(label).findNext();
  return (cell ? cell.getBackground() : fallback).toLowerCase();
}
function listRows_() {
  const sh = sourceSheet_();
  const last = sh.getLastRow(); if (last < CFG.FIRST_ROW) return [];
  const n = last - CFG.FIRST_ROW + 1;
  const width = Math.max(CFG.COL.extra, Math.max.apply(null, CFG.STATUS_COLS));
  const rng = sh.getRange(CFG.FIRST_ROW, 1, n, width);
  const vals = rng.getDisplayValues(), bgs = rng.getBackgrounds(), notes = rng.getNotes();
  const rich = sh.getRange(CFG.FIRST_ROW, CFG.COL.link, n, 1).getRichTextValues();
  const chips = chipLinks_(CFG.FIRST_ROW, last);
  const cCur = legendColor_(sh, CFG.CURRENT_LABEL, CFG.FALLBACK_CURRENT);
  const cNext = legendColor_(sh, CFG.NEXT_LABEL, CFG.FALLBACK_NEXT);
  const rows = [];
  for (let i = 0; i < n; i++) {
    const v = vals[i];
    const text = String(v[CFG.COL.link - 1] || '').trim(), who = String(v[CFG.COL.who - 1] || '').trim(), type = String(v[CFG.COL.type - 1] || '').trim();
    // строка трека: есть «что», есть «тип», и это не объединённая служебная строка (там текст повторяется во всех колонках)
    if (!text || !type || type === text || type.length > 40 || type === 'Тип') continue;
    if (CFG.PLACEHOLDERS.indexOf(text.toUpperCase()) >= 0) continue;
    const row = CFG.FIRST_ROW + i;
    const link = linkOf_(rich[i][0], text) || chips[row] || '';
    const colors = CFG.STATUS_COLS.map(c => String(bgs[i][c - 1]).toLowerCase());
    const status = colors.indexOf(cCur) >= 0 ? 'current' : colors.indexOf(cNext) >= 0 ? 'next' : '';
    rows.push({
      row: row, title: text, link: link, key: link || text, who: who, type: type,
      genre: v[CFG.COL.genre - 1], extra: extraOf_(v[CFG.COL.extra - 1], notes[i]), status: status,
    });
  }
  return rows;
}
// «Доп. инфа»: если там отсылка вида «Комментарий^», берём настоящий текст из примечания ячейки
// (сначала из самой «Доп. инфа», потом из любой другой ячейки строки).
function extraOf_(cell, rowNotes) {
  cell = String(cell || '').trim();
  const own = String(rowNotes[CFG.COL.extra - 1] || '').trim();
  const any = rowNotes.map(x => String(x || '').trim()).filter(Boolean);
  const note = own || any[0] || '';
  if (/^коммент/i.test(cell)) return note;         // «Комментарий^» → текст примечания
  if (cell && note && cell.indexOf(note) < 0) return cell + ' — ' + note;
  return cell || note;
}
function findRow_(row) { return listRows_().filter(r => r.row === row)[0] || null; }
function linkOf_(rt, text) {
  if (rt) {
    const direct = rt.getLinkUrl(); if (direct) return direct;
    const runs = rt.getRuns(); for (let i = 0; i < runs.length; i++) { const u = runs[i].getLinkUrl(); if (u) return u; }
  }
  const m = String(text).match(/https?:\/\/\S+/); return m ? m[0] : '';
}
// Ссылки из «умных чипов» (YouTube-плашки с иконкой). Нужен сервис «Google Sheets API» (+ Сервисы).
// Если сервис не подключён — просто пропускаем, трек найдётся по названию.
function chipLinks_(first, last) {
  const out = {};
  try {
    if (typeof Sheets === 'undefined') return out;
    const res = Sheets.Spreadsheets.get(CFG.SOURCE_ID, {
      ranges: ["'" + CFG.SHEET + "'!A" + first + ':A' + last],
      fields: 'sheets.data.rowData.values(hyperlink,chipRuns)',
    });
    const rd = (((res.sheets || [])[0] || {}).data || [])[0].rowData || [];
    rd.forEach((r, i) => {
      const c = (r.values || [])[0] || {};
      let u = c.hyperlink || '';
      (c.chipRuns || []).forEach(run => { const rl = run.chip && run.chip.richLinkProperties; if (!u && rl && rl.uri) u = rl.uri; });
      if (u) out[first + i] = u;
    });
  } catch (e) { /* сервис не подключён или нет прав — не страшно */ }
  return out;
}

// ───────────────────────── хранилище карточек (твоя таблица) ─────────────────────────
function cardsBook_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('CARDS_ID'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Живые справки — карточки');
    const sh = ss.getSheets()[0]; sh.setName('Карточки');
    sh.appendRow(['ключ (ссылка или название)', 'строка', 'статус', 'карточка (JSON)', 'обновлено']);
    sh.setFrozenRows(1); sh.setColumnWidth(1, 320); sh.setColumnWidth(4, 700);
    props.setProperty('CARDS_ID', ss.getId());
  }
  return ss;
}
function cardsSheet_() { return cardsBook_().getSheets()[0]; }
function cardIndex_() {
  const sh = cardsSheet_(); const last = sh.getLastRow();
  const map = {}; if (last < 2) return map;
  sh.getRange(2, 1, last - 1, 4).getValues().forEach((v, i) => { map[v[0]] = { rowIdx: i + 2, status: v[2], json: v[3] }; });
  return map;
}
function getCard_(key) {
  const c = cardIndex_()[key]; if (!c) return null;
  if (!c.json) return { status: c.status };
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
    rows.sort((a, b) => (b.status ? 1 : 0) - (a.status ? 1 : 0)); // сначала текущий и следующий
    for (let i = 0; i < rows.length && done < CFG.PER_RUN; i++) {
      const r = rows[i], c = idx[r.key];
      if (c && (c.status === 'ready' || c.status === 'unknown')) continue;
      if (c && c.status === 'error' && !r.status) continue; // ошибки не повторяем бесконечно
      generateFor_(r); done++;
    }
  } finally { lock.releaseLock(); }
}

/** Ручная проверка из редактора: поменяй номер строки и запусти. */
function testRow() { Logger.log(JSON.stringify(generateFor_(findRow_(10)), null, 2)); }

function generateFor_(r) {
  if (!r) throw new Error('строка не найдена');
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

// Метаданные со страницы: YouTube через oEmbed, остальное — og-теги (Spotify, Яндекс Музыка, SoundCloud…)
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
    const og = (p) => {
      const m = html.match(new RegExp('<meta[^>]+property=["\']og:' + p + '["\'][^>]+content=["\']([^"\']*)', 'i')) ||
                html.match(new RegExp('<meta[^>]+content=["\']([^"\']*)["\'][^>]+property=["\']og:' + p, 'i'));
      return m ? decode_(m[1]) : '';
    };
    meta.title = og('title'); meta.image = og('image'); meta.description = og('description');
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
    'Как записано в таблице: ' + r.title,
    meta.title ? 'Название со страницы: ' + meta.title : '',
    meta.artist ? 'Исполнитель/канал со страницы: ' + meta.artist : '',
    meta.description ? 'Описание страницы: ' + meta.description : '',
    r.genre ? 'Жанр (по таблице): ' + r.genre : '',
    r.extra ? 'Комментарий заказчика (может подсказать, что это за трек): ' + r.extra : '',
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
    '- 2–4 факта: история создания, смысл текста, награды и чарты, сэмплы, коллаборации, откуда трек (игра, аниме, фильм), необычные детали записи.\n' +
    '- Не пересказывай очевидное (название, длительность).\n' +
    '- Если трек найти не удалось или это малоизвестный авторский трек без информации в сети — верни "known": false и пустой facts. НИЧЕГО не выдумывай.\n\n' +
    info;
  const res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + key, 'HTTP-Referer': 'https://laitsberg.github.io/live-spravki/', 'X-Title': 'Live Spravki' },
    payload: JSON.stringify({ model: model, temperature: 0.2, messages: [{ role: 'user', content: prompt }] }),
  });
  const code = res.getResponseCode(), body = res.getContentText();
  if (code >= 400) throw new Error('OpenRouter ' + code + ': ' + body.slice(0, 200));
  const txt = JSON.parse(body).choices[0].message.content || '';
  const m = txt.match(/\{[\s\S]*\}/); if (!m) throw new Error('модель не вернула JSON');
  return JSON.parse(m[0]);
}


// ───────────────────────── словарь терминов (редактируется из пульта) ─────────────────────────
// Хранится на вкладке «Термины» твоей таблицы карточек. Колонки: id | надпись | заголовок | пояснение | слова | вкл | обновлено
function termsSheet_() {
  const ss = cardsBook_(); let sh = ss.getSheetByName('Термины');
  if (!sh) {
    sh = ss.insertSheet('Термины');
    sh.appendRow(['id', 'надпись (kicker)', 'заголовок', 'пояснение', 'слова (через запятую; «!» в конце — только точное слово)', 'вкл', 'обновлено']);
    sh.setFrozenRows(1); sh.setColumnWidth(4, 420); sh.setColumnWidth(5, 320);
  }
  return sh;
}
function readTerms_() {
  const sh = termsSheet_(); const last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 6).getValues().filter(v => v[0]).map(v => ({
    id: String(v[0]), kicker: String(v[1]), title: String(v[2]), text: String(v[3]),
    match: String(v[4]).split(',').map(x => x.trim()).filter(Boolean), enabled: v[5] !== false && String(v[5]).toLowerCase() !== 'false',
  }));
}
function termRow_(t) { return [t.id, t.kicker || '', t.title || '', t.text || '', (t.match || []).join(', '), t.enabled !== false, new Date()]; }

// Запись идёт POST-запросом из пульта (text/plain, чтобы браузер не делал лишних проверок). Нужен токен.
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!body.token || body.token !== PropertiesService.getScriptProperties().getProperty('GEN_TOKEN')) throw new Error('неверный токен');
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      const sh = termsSheet_(); const last = sh.getLastRow();
      const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(v => String(v[0])) : [];
      if (body.action === 'saveTerm') {
        const t = body.term || {}; if (!t.id) t.id = 't' + Date.now().toString(36);
        if (!t.kicker || !(t.match || []).length) throw new Error('нужны надпись и хотя бы одно слово');
        const i = ids.indexOf(String(t.id));
        if (i >= 0) sh.getRange(i + 2, 1, 1, 7).setValues([termRow_(t)]); else sh.appendRow(termRow_(t));
      } else if (body.action === 'deleteTerm') {
        const i = ids.indexOf(String(body.id)); if (i >= 0) sh.deleteRow(i + 2);
      } else if (body.action === 'seedTerms') {
        if (ids.length) throw new Error('словарь уже есть');
        const rows = (body.terms || []).map(termRow_); if (rows.length) sh.getRange(2, 1, rows.length, 7).setValues(rows);
      } else throw new Error('неизвестное действие');
    } finally { lock.releaseLock(); }
    return json_({ ok: true, terms: readTerms_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}
