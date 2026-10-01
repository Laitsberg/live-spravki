// Оформление карточек: значения по умолчанию, готовые темы, шрифты.
// Общий файл для пульта (окно «Оформление») и оверлея (применяет стиль).
window.CardStyle = (function () {
  // Шрифты Google Fonts с кириллицей. Значение — какие начертания запрашивать.
  const FONTS = {
    'Montserrat': 'wght@400..900', 'Unbounded': 'wght@400..900', 'Inter': 'wght@400..900',
    'Manrope': 'wght@400..800', 'Onest': 'wght@400..900', 'Golos Text': 'wght@400..900',
    'Geologica': 'wght@400..900', 'Rubik': 'wght@400..900', 'Nunito': 'wght@400..900',
    'Jost': 'wght@400..900', 'Exo 2': 'wght@400..900', 'Raleway': 'wght@400..900',
    'Oswald': 'wght@400..700', 'Roboto Condensed': 'wght@400..900', 'Tektur': 'wght@400..900',
    'PT Sans': 'wght@400;700', 'PT Serif': 'wght@400;700', 'Playfair Display': 'wght@400..900',
    'Lora': 'wght@400..700', 'Comfortaa': 'wght@400..700', 'Caveat': 'wght@400..700',
    'JetBrains Mono': 'wght@400..800', 'Russo One': '', 'Rubik Mono One': '', 'Lobster': '', 'Press Start 2P': '',
  };

  const DEFAULT = {
    v: 0,                     // время сохранения: оверлей применяет только более свежий стиль
    pos: 'tl',
    term: '#6fe3ff', fact: '#ffce5c', track: '#ff5a6e', custom: '#b18cff',
    bg: '#0a101e', bgAlpha: 92, gloss: true,
    text: '#ffffff', muted: 74, border: 14, shadow: 55,
    fontHead: 'Unbounded', fontBody: 'Montserrat',
    title: 30, body: 22, kicker: 15, upper: true,
    width: 640, radius: 26, scale: 100, margin: 48,
    stripe: true, badge: true, timer: true, anim: 'slide',
  };
  // Темы не трогают положение и размер карточки на экране.
  const PLACEMENT = ['pos', 'scale', 'margin', 'width'];

  const PRESETS = [
    { id: 'night', name: 'Ночь', s: {} },
    { id: 'neon', name: 'Неон', s: { bg: '#0d0221', bgAlpha: 88, term: '#00f0ff', fact: '#f9f871', track: '#ff2e97', custom: '#a86bff',
      border: 32, radius: 14, fontHead: 'Tektur', fontBody: 'Onest', anim: 'pop' } },
    { id: 'light', name: 'Светлая', s: { bg: '#ffffff', bgAlpha: 96, text: '#141a2b', muted: 74, term: '#0a7bd1', fact: '#c27800', track: '#e0314b', custom: '#7a4ddb',
      border: 10, shadow: 30, gloss: false, fontHead: 'Manrope', fontBody: 'Manrope' } },
    { id: 'minimal', name: 'Минимал', s: { bg: '#000000', bgAlpha: 72, gloss: false, border: 0, shadow: 20, radius: 10, badge: false, timer: false,
      fontHead: 'Inter', fontBody: 'Inter', title: 28, body: 20, kicker: 13, anim: 'fade' } },
    { id: 'vinyl', name: 'Винил', s: { bg: '#24150d', bgAlpha: 94, text: '#fff4e6', term: '#ffb36b', fact: '#ffd66b', track: '#ff6b4a', custom: '#e6a1ff',
      border: 18, radius: 18, fontHead: 'Playfair Display', fontBody: 'Lora', upper: false, anim: 'rise' } },
    { id: 'arcade', name: 'Аркада', s: { bg: '#0b1020', bgAlpha: 95, term: '#7dff5a', fact: '#ffe156', track: '#ff4f7b', custom: '#5ad8ff',
      border: 40, radius: 0, shadow: 0, gloss: false, fontHead: 'Press Start 2P', fontBody: 'JetBrains Mono', title: 22, body: 19, kicker: 12, anim: 'pop' } },
  ];

  const clamp = (x, a, b, d) => { x = Number(x); return isFinite(x) ? Math.max(a, Math.min(b, x)) : d; };
  const isHex = (c) => /^#[0-9a-f]{6}$/i.test(String(c || ''));
  const ANIMS = ['slide', 'rise', 'pop', 'fade', 'none'];
  const LIMITS = { bgAlpha: [0, 100], muted: [20, 100], border: [0, 100], shadow: [0, 100], title: [12, 72], body: [10, 48], kicker: [8, 32],
    width: [360, 1200], radius: [0, 60], scale: [30, 200], margin: [0, 400] };

  // Привести любой (в т.ч. старый или чужой) объект к полному и безопасному стилю.
  function normalize(s) {
    s = Object.assign({}, DEFAULT, s || {});
    const out = { v: Number(s.v) || 0 };
    out.pos = ['tl', 'tr', 'bl', 'br'].includes(s.pos) ? s.pos : DEFAULT.pos;
    for (const k of ['term', 'fact', 'track', 'custom', 'bg', 'text']) out[k] = isHex(s[k]) ? s[k].toLowerCase() : DEFAULT[k];
    for (const k in LIMITS) out[k] = Math.round(clamp(s[k], LIMITS[k][0], LIMITS[k][1], DEFAULT[k]));
    for (const k of ['gloss', 'upper', 'stripe', 'badge', 'timer']) out[k] = s[k] !== false;
    for (const k of ['fontHead', 'fontBody']) out[k] = s[k] in FONTS ? s[k] : DEFAULT[k];
    out.anim = ANIMS.includes(s.anim) ? s.anim : DEFAULT.anim;
    return out;
  }
  function withPreset(cur, id) {
    const p = PRESETS.find(x => x.id === id) || PRESETS[0];
    const keep = {}; PLACEMENT.forEach(k => keep[k] = cur[k]);
    return normalize(Object.assign({}, DEFAULT, p.s, keep));
  }
  function fontUrl(names) {
    const fam = [...new Set(names)].filter(n => n in FONTS)
      .map(n => 'family=' + n.replace(/ /g, '+') + (FONTS[n] ? ':' + FONTS[n] : ''));
    return 'https://fonts.googleapis.com/css2?' + fam.join('&') + '&display=swap';
  }
  function rgba(hex, a) {
    const n = parseInt(String(hex).slice(1), 16);
    return `rgba(${n >> 16 & 255}, ${n >> 8 & 255}, ${n & 255}, ${Math.round(a * 1000) / 1000})`;
  }
  return { FONTS, DEFAULT, PRESETS, PLACEMENT, normalize, withPreset, fontUrl, rgba };
})();
