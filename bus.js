// Канал связи пульт → оверлей.
// 1) BroadcastChannel — если пульт и оверлей открыты в одном браузере (локальный тест).
// 2) ntfy.sh — бесплатный публичный pub/sub, работает между разными компьютерами.
//    Имя «комнаты» (topic) — длинная случайная строка, её знают только пульт и оверлей.
window.Bus = (function () {
  const NTFY = 'https://ntfy.sh/';
  const seen = new Set();
  let bc = null;
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }

  function publish(topic, msg) {
    msg = Object.assign({ id: uid(), ts: Date.now() }, msg);
    try { (bc || (bc = new BroadcastChannel('live-spravki-' + topic))).postMessage(msg); } catch (e) {}
    if (topic) {
      return fetch(NTFY + encodeURIComponent(topic), { method: 'POST', body: JSON.stringify(msg) })
        .then(r => r.ok).catch(() => false);
    }
    return Promise.resolve(true);
  }

  function subscribe(topic, onMsg, onStatus) {
    const handle = (m) => {
      if (!m || !m.id || seen.has(m.id)) return;
      seen.add(m.id);
      if (m.ts && Date.now() - m.ts > 60000) return; // старьё не показываем
      onMsg(m);
    };
    try {
      const ch = new BroadcastChannel('live-spravki-' + topic);
      ch.onmessage = (e) => handle(e.data);
    } catch (e) {}
    if (!topic) return;
    let es, retry = 1000;
    const connect = () => {
      es = new EventSource(NTFY + encodeURIComponent(topic) + '/sse');
      const onData = (e) => {
        try {
          const d = JSON.parse(e.data);
          if (d.event && d.event !== 'message') return;
          handle(JSON.parse(d.message));
        } catch (err) {}
      };
      es.onmessage = onData;
      es.addEventListener('message', onData);
      es.onopen = () => { retry = 1000; onStatus && onStatus('online'); };
      es.onerror = () => {
        onStatus && onStatus('offline');
        es.close(); setTimeout(connect, retry); retry = Math.min(retry * 2, 15000);
      };
    };
    connect();
  }

  function newTopic() {
    const a = new Uint8Array(12); crypto.getRandomValues(a);
    return 'spravki-' + Array.from(a, b => b.toString(16).padStart(2, '0')).join('');
  }
  return { publish, subscribe, newTopic };
})();
