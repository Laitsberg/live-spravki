/**
 * Живые справки — ЗАГРУЗЧИК.
 *
 * Вставь этот файл в Apps Script ОДИН раз вместо всего кода — и больше код руками обновлять не надо.
 * Настоящий код скрипт сам берёт из GitHub (Laitsberg/live-spravki → apps-script/Code.gs)
 * и обновляется примерно через 5 минут после каждого изменения в репозитории.
 * Новую версию развёртывания тоже делать не нужно.
 */
const CODE_URL = 'https://raw.githubusercontent.com/Laitsberg/live-spravki/main/apps-script/Code.gs';
const EXPORTS = ['doGet', 'doPost', 'setup', 'checkQueue', 'processPending', 'testRow'];

function lib_() {
  if (globalThis.__LS) return globalThis.__LS;
  const cache = CacheService.getScriptCache();
  let code = cache.get('ls_code');
  if (!code) {
    try {
      const r = UrlFetchApp.fetch(CODE_URL + '?t=' + Date.now(), { muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) throw new Error('GitHub ответил ' + r.getResponseCode());
      code = r.getContentText();
      cache.put('ls_code', code, 300);          // свежесть — 5 минут
      cache.put('ls_code_backup', code, 21600); // запасная копия на 6 часов, если GitHub недоступен
    } catch (e) {
      code = cache.get('ls_code_backup');
      if (!code) throw new Error('Не удалось загрузить код с GitHub: ' + e.message);
    }
  }
  globalThis.__LS = new Function(code + '\nreturn {' + EXPORTS.map(n => n + ': ' + n).join(', ') + '};')();
  return globalThis.__LS;
}

function doGet(e) { return lib_().doGet(e); }
function doPost(e) { return lib_().doPost(e); }
function setup() { return lib_().setup(); }
function checkQueue() { return lib_().checkQueue(); }
function processPending() { return lib_().processPending(); }
function testRow() { return lib_().testRow(); }

/** Сбросить кэш, если нужно подтянуть обновление с GitHub прямо сейчас. */
function refreshCode() { CacheService.getScriptCache().removeAll(['ls_code', 'ls_code_backup']); Logger.log('Кэш сброшен, код загрузится заново.'); }

// Эта функция никогда не вызывается. Она нужна, чтобы Google понял, какие разрешения нужны скрипту
// (код с GitHub он не видит).
function permissions_() {
  SpreadsheetApp.openById(''); SpreadsheetApp.create(''); UrlFetchApp.fetch('');
  ScriptApp.newTrigger(''); PropertiesService.getScriptProperties(); LockService.getScriptLock();
  CacheService.getScriptCache(); ContentService.createTextOutput(''); Utilities.getUuid(); Logger.log('');
}
