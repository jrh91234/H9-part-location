// H9 Part Location — ค้นหา/สแกน ว่ารหัส Part อยู่ Location ไหน (ข้อมูลจาก js/data.js)
(function () {
'use strict';
var APP_VERSION = '20261126-0733'; // ตั้งโดย tools/bump-version.py

var PARTS = window.PARTS || [], LEVELS = window.LEVELS || {};
var LV_LIST = Object.keys(LEVELS).map(Number).sort(function (a, b) { return b - a; }); // ชั้นบนสุดก่อน (เหมือนแผนผัง)
var MAX_COL = Math.max.apply(null, LV_LIST.map(function (l) { return LEVELS[l]; }).concat([1]));
var LINES = [4, 5, 1];

// ---------- ดัชนี ----------
var byLoc = {}, byCode = {}, bySpec = {};
PARTS.forEach(function (p, i) {
  p.i = i;
  p.short = p.loc ? p.loc.replace(/^H9-/, '') : '';
  p.hay = [p.loc, p.short, p.code, p.spec, p.name.toUpperCase(), p.grp.toUpperCase(), p.loc ? '' : 'รอจัดพื้นที่'].join(' ');
  if (p.loc) byLoc[p.loc] = p;
  (byCode[p.code] = byCode[p.code] || []).push(p);
  if (p.spec) (bySpec[p.spec] = bySpec[p.spec] || []).push(p);
});
function locOrder(a, b) {
  if (!a.loc !== !b.loc) return a.loc ? -1 : 1;
  return a.lv - b.lv || a.col - b.col || a.i - b.i;
}
// รับได้ทั้ง H9-L3-05, L3-05, L3-5, l3 5, H9L305
function normLoc(t) {
  var m = String(t || '').toUpperCase().trim().match(/^(?:H9[-\s]?)?L(\d)[-\s]?(\d{1,2})$/);
  if (!m) return '';
  return 'H9-L' + (+m[1]) + '-' + ('0' + (+m[2])).slice(-2);
}

// ---------- สถานะ ----------
var S = {
  tab: 'search', q: '', line: 'all', grp: 'all',
  scan: null,           // { text, parts } ผลจากการสแกนล่าสุด (แสดงแทนผลค้นหาจนกว่าจะพิมพ์ใหม่)
  mapLv: LV_LIST[0] || 1, mark: ''
};
var GROUPS = [];
PARTS.forEach(function (p) { if (p.grp && GROUPS.indexOf(p.grp) < 0) GROUPS.push(p.grp); });
GROUPS.sort();

// ---------- เครื่องมือ ----------
function $(id) { return document.getElementById(id); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function store(k, v) {
  try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { return null; }
}
var toastT;
function toast(msg, ms) {
  var t = $('toast'); t.textContent = msg; t.classList.remove('hidden');
  clearTimeout(toastT); toastT = setTimeout(function () { t.classList.add('hidden'); }, ms || 1800);
}
function openModal(html) {
  var m = $('modal');
  m.innerHTML = '<div class="sheet" id="sheet">' + html + '</div>';
  m.classList.remove('hidden');
  m.onclick = function (e) { if (e.target === m) closeModal(); };
}
var onModalClose = null;
function closeModal() {
  $('modal').classList.add('hidden'); $('modal').innerHTML = '';
  if (onModalClose) { var f = onModalClose; onModalClose = null; f(); }
}
var actx;
function beep(ok) {
  try {
    actx = actx || new (window.AudioContext || window.webkitAudioContext)();
    (ok ? [[1400, 0]] : [[300, 0], [300, 0.18]]).forEach(function (b) {
      var o = actx.createOscillator(), g = actx.createGain();
      o.frequency.value = b[0]; o.connect(g); g.connect(actx.destination);
      var t = actx.currentTime + b[1];
      g.gain.setValueAtTime(0.25, t); g.gain.exponentialRampToValueAtTime(0.001, t + 0.13);
      o.start(t); o.stop(t + 0.14);
    });
  } catch (e) { /* ไม่มีเสียงก็ไม่เป็นไร */ }
  if (navigator.vibrate) navigator.vibrate(ok ? 60 : [80, 60, 80]);
}
function copy(text) {
  var done = function () { toast('คัดลอก ' + text); };
  if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(text).then(done, fallback); else fallback();
  function fallback() {
    var t = document.createElement('textarea'); t.value = text; document.body.appendChild(t); t.select();
    try { document.execCommand('copy'); done(); } catch (e) { toast('คัดลอกไม่ได้'); }
    document.body.removeChild(t);
  }
}
function lnCls(p) { return LINES.indexOf(p.line) >= 0 ? 'ln' + p.line : ''; }
function hl(text, toks) {
  var s = esc(text);
  toks.forEach(function (t) {
    if (t.length < 2) return;
    var re = new RegExp('(' + esc(t).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'ig');
    s = s.replace(re, '\u0001$1\u0002');
  });
  return s.replace(/\u0001/g, '<mark>').replace(/\u0002/g, '</mark>');
}

// ---------- ค้นหา ----------
function tokens(q) { return String(q || '').toUpperCase().trim().split(/\s+/).filter(Boolean); }
function search(q) {
  var toks = tokens(q), nl = normLoc(q);
  var list = PARTS.filter(function (p) {
    if (S.line !== 'all' && p.line !== S.line) return false;
    if (S.grp !== 'all' && p.grp !== S.grp) return false;
    if (nl) return p.loc === nl;
    for (var i = 0; i < toks.length; i++) if (p.hay.indexOf(toks[i]) < 0) return false;
    return true;
  });
  if (!toks.length || nl) return list.sort(locOrder);
  var Q = toks.join(' ');
  function rank(p) {
    if (p.code === Q || p.spec === Q || p.short === Q || p.loc === Q) return 0;
    if (p.code.indexOf(Q) === 0 || p.spec.indexOf(Q) === 0 || p.short.indexOf(Q) === 0) return 1;
    if (p.code.indexOf(Q) >= 0 || p.spec.indexOf(Q) >= 0) return 2;
    return 3;
  }
  return list.sort(function (a, b) { return rank(a) - rank(b) || locOrder(a, b); });
}

// ข้อความจากเครื่องสแกน/กล้อง → หา Part ที่ตรง
// 1) ทั้งข้อความเป็น Location / Code / Spec  2) แยกเป็นคำแล้วหาทีละคำ  3) หารหัสที่ซ่อนอยู่ในข้อความยาว (QR หลายช่อง)
function resolveScan(text) {
  var lq = String(text || '').match(/#q=([^&\s]+)/);
  if (lq) { try { text = decodeURIComponent(lq[1]); } catch (e) { text = lq[1]; } }
  var U = String(text || '').toUpperCase().replace(/[\r\n\t]+/g, ' ').trim();
  if (!U) return null;
  function by(t) {
    var l = normLoc(t); if (l && byLoc[l]) return { parts: [byLoc[l]], how: 'loc' };
    if (byCode[t]) return { parts: byCode[t], how: 'code' };
    if (bySpec[t]) return { parts: bySpec[t], how: 'spec' };
    return null;
  }
  var r = by(U); if (r) return r;
  var toks = U.split(/[^A-Z0-9\-\/]+/).filter(Boolean), found = [], how = '';
  function add(list, h) { list.forEach(function (p) { if (found.indexOf(p) < 0) found.push(p); }); how = how || h; }
  toks.forEach(function (t) { var x = by(t); if (x) add(x.parts, x.how); });
  if (!found.length) {
    Object.keys(byCode).forEach(function (k) { if (U.indexOf(k) >= 0) add(byCode[k], 'code'); });
    if (!found.length) Object.keys(bySpec).forEach(function (k) { if (k.length >= 6 && U.indexOf(k) >= 0) add(bySpec[k], 'spec'); });
  }
  return found.length ? { parts: found.sort(locOrder), how: how } : null;
}

function handleScan(text, src) {
  text = String(text || '').trim();
  if (!text) return;
  var r = resolveScan(text);
  pushRecent(text);
  if (!r) {
    beep(false);
    S.scan = { text: text, parts: [] };
    S.q = ''; if ($('q')) $('q').value = '';
    if (S.tab !== 'search') setTab('search'); else renderResults();
    return;
  }
  beep(true);
  S.scan = { text: text, parts: r.parts, how: r.how };
  S.q = ''; if ($('q')) $('q').value = '';
  if (S.tab !== 'search') setTab('search'); else renderResults();
  if (r.parts.length === 1) showDetail(r.parts[0]);
  if (src === 'wedge' && $('q')) $('q').select();
}

function pushRecent(t) {
  var a = recent().filter(function (x) { return x !== t; });
  a.unshift(t); store('h9pl_recent', JSON.stringify(a.slice(0, 8)));
}
function recent() { try { return JSON.parse(store('h9pl_recent') || '[]'); } catch (e) { return []; } }

// ---------- หน้าจอ ----------
function render() {
  $('app').innerHTML =
    '<div class="top"><div class="brand">📍 H9 Part Location <span class="ver">v' + esc(APP_VERSION) + '</span></div>' +
      '<button class="chip" id="appQr" aria-label="QR ลิงก์เข้าหน้าเว็บ">🔗 QR แอป</button></div>' +
    '<div class="tabs">' +
      '<button data-tab="search" class="' + (S.tab === 'search' ? 'active' : '') + '">🔎 ค้นหา</button>' +
      '<button data-tab="check" class="' + (S.tab === 'check' ? 'active' : '') + '">✅ ไล่เช็ค</button>' +
      '<button data-tab="map" class="' + (S.tab === 'map' ? 'active' : '') + '">🗂 ผัง</button>' +
    '</div><main id="main"></main>';
  Array.prototype.forEach.call(document.querySelectorAll('[data-tab]'), function (b) {
    b.onclick = function () { setTab(b.getAttribute('data-tab')); };
  });
  $('appQr').onclick = function () {
    showQr(APP_URL, 'H9 Part Location', 'สแกน QR นี้ด้วยกล้องมือถือเพื่อเปิดแอปค้นหา Part ↔ Location');
  };
  if (S.tab === 'search') renderSearch(); else if (S.tab === 'check') renderCheck(); else renderMap();
}
function setTab(t) { S.tab = t; render(); window.scrollTo(0, 0); }

function renderSearch() {
  var chips = '<button class="fchip' + (S.line === 'all' ? ' active' : '') + '" data-line="all">ทุก Line</button>' +
    LINES.map(function (l) { return '<button class="fchip l' + l + (S.line === l ? ' active' : '') + '" data-line="' + l + '">Line ' + l + '</button>'; }).join('') +
    '<button class="fchip' + (S.grp !== 'all' ? ' active' : '') + '" id="grpBtn">' + esc(S.grp === 'all' ? 'กลุ่ม Part ▾' : S.grp + ' ✕') + '</button>';
  $('main').innerHTML =
    '<div class="search"><div class="search-wrap">' +
      '<input id="q" type="text" enterkeyhint="search" autocomplete="off" autocapitalize="characters" spellcheck="false" ' +
        'placeholder="Location / รหัส / Spec / ชื่อ" value="' + esc(S.q) + '">' +
      '<button class="clear' + (S.q ? '' : ' hidden') + '" id="qClear" aria-label="ล้าง">✕</button></div>' +
      '<button class="scan-btn" id="camBtn">📷 สแกน</button></div>' +
    '<div class="filters">' + chips + '</div>' +
    '<div id="results"></div>';
  var q = $('q');
  q.oninput = function () { S.q = q.value; S.scan = null; $('qClear').classList.toggle('hidden', !q.value); renderResults(); };
  // เครื่องสแกนแบบคีย์บอร์ด (PDA / หัวสแกน Bluetooth) ส่ง Enter ตอนจบ → ถือเป็นการสแกน
  q.onkeydown = function (e) {
    if (e.key === 'Enter') { e.preventDefault(); var v = q.value; if (v.trim()) { handleScan(v, 'wedge'); } }
  };
  $('qClear').onclick = function () { S.q = ''; S.scan = null; q.value = ''; this.classList.add('hidden'); renderResults(); q.focus(); };
  $('camBtn').onclick = function () { openCameraScanner(function (t) { handleScan(t, 'cam'); }); };
  Array.prototype.forEach.call(document.querySelectorAll('[data-line]'), function (b) {
    b.onclick = function () { var v = b.getAttribute('data-line'); S.line = v === 'all' ? 'all' : +v; renderSearch(); };
  });
  $('grpBtn').onclick = function () {
    if (S.grp !== 'all') { S.grp = 'all'; renderSearch(); return; }
    openModal('<h3>เลือกกลุ่ม Part</h3><div class="recent">' + GROUPS.map(function (g, i) {
      return '<button data-g="' + i + '">' + esc(g) + '</button>';
    }).join('') + '</div><div class="actions"><button class="btn" id="mClose">ปิด</button></div>');
    $('mClose').onclick = closeModal;
    Array.prototype.forEach.call(document.querySelectorAll('[data-g]'), function (b) {
      b.onclick = function () { S.grp = GROUPS[+b.getAttribute('data-g')]; closeModal(); renderSearch(); };
    });
  };
  renderResults();
}

var LIMIT = 80;
function renderResults() {
  var box = $('results'); if (!box) return;
  var html = '', list, toks = tokens(S.q);
  if (S.scan) {
    var sc = S.scan;
    if (!sc.parts.length) {
      html += '<div class="msg err">❌ ไม่พบในรายการ: <b>' + esc(sc.text) + '</b><br><span class="small">ลองสแกนใหม่ หรือพิมพ์บางส่วนของรหัสเพื่อค้นหา</span></div>';
      list = [];
    } else {
      var what = sc.how === 'loc' ? 'Location' : sc.how === 'spec' ? 'Spec' : 'รหัส';
      html += '<div class="msg ok">✅ สแกน' + what + ' <b>' + esc(sc.text.length > 60 ? sc.text.slice(0, 60) + '…' : sc.text) + '</b> → พบ ' + sc.parts.length + ' รายการ</div>';
      list = sc.parts;
      toks = [];
    }
  } else {
    list = search(S.q);
    if (!toks.length && S.line === 'all' && S.grp === 'all') {
      var rc = recent();
      html += '<div class="msg idle">พิมพ์ Location (เช่น <b>L3-05</b>), รหัส Part, Spec หรือชื่อ — หรือกด <b>📷 สแกน</b> เพื่อสแกนป้าย QR/บาร์โค้ด' +
        '<br><span class="small">ใช้ PDA / หัวสแกนได้: แตะช่องค้นหาแล้วยิงสแกนได้เลย</span></div>';
      if (rc.length) html += '<div class="count"><span>สแกน/ค้นหาล่าสุด</span></div><div class="recent">' +
        rc.map(function (t, i) { return '<button data-rc="' + i + '">' + esc(t.length > 24 ? t.slice(0, 24) + '…' : t) + '</button>'; }).join('') + '</div>';
    }
  }
  if (list.length || !S.scan) {
    var waiting = list.filter(function (p) { return !p.loc; }).length;
    html += '<div class="count"><span>' + list.length + ' รายการ' + (waiting ? ' · รอจัดพื้นที่ ' + waiting : '') + '</span>' +
      (list.length > LIMIT ? '<span>แสดง ' + LIMIT + ' รายการแรก</span>' : '') + '</div>';
    html += list.slice(0, LIMIT).map(function (p) { return cardHtml(p, toks, S.scan && S.scan.parts.length); }).join('');
    if (!list.length) html += '<div class="msg warn">ไม่พบรายการที่ตรงกับ “' + esc(S.q) + '”</div>';
  }
  box.innerHTML = html;
  Array.prototype.forEach.call(box.querySelectorAll('[data-i]'), function (b) {
    b.onclick = function () { showDetail(PARTS[+b.getAttribute('data-i')]); };
  });
  var rcs = recent();
  Array.prototype.forEach.call(box.querySelectorAll('[data-rc]'), function (b) {
    b.onclick = function () { handleScan(rcs[+b.getAttribute('data-rc')], 'recent'); };
  });
}

function cardHtml(p, toks, hit) {
  var loc = p.loc
    ? '<div class="card-loc"><div class="sm">H9</div><div class="big">' + hl(p.short, toks) + '</div><div class="sm">ชั้น ' + p.lv + ' · ช่อง ' + p.col + '</div></div>'
    : '<div class="card-loc none"><div class="big">รอจัด<br>พื้นที่</div></div>';
  return '<button class="card ' + lnCls(p) + (hit ? ' hit' : '') + '" data-i="' + p.i + '">' + loc +
    '<div class="card-main"><div class="code">' + hl(p.code, toks) + '</div>' +
    '<div class="spec">' + hl(p.spec, toks) + '</div>' +
    '<div class="name">' + hl(p.name, toks) + '</div>' +
    '<div class="tags"><span class="tag ' + lnCls(p) + '">Line ' + esc(p.line) + '</span><span class="tag">' + esc(p.grp) + '</span>' +
      (p.st && p.st !== 'Approved' ? '<span class="tag warn">' + esc(p.st) + '</span>' : '') + '</div></div></button>';
}

function rackHtml(mark, clickable) {
  var rows = LV_LIST.map(function (lv) {
    var cells = '';
    for (var c = 1; c <= MAX_COL; c++) {
      var loc = 'H9-L' + lv + '-' + ('0' + c).slice(-2), p = byLoc[loc];
      var cls = c > LEVELS[lv] ? 'void' : p ? lnCls(p) : '';
      cells += '<div class="rc ' + cls + (loc === mark ? ' on' : '') + '"></div>';
    }
    var on = mark && mark.indexOf('H9-L' + lv + '-') === 0;
    return '<div class="rack-row"' + (clickable ? ' data-lv="' + lv + '"' : '') + '><div class="rack-lv' + (on ? ' on' : '') + '">L' + lv + '</div><div class="rack-cells">' + cells + '</div></div>';
  }).join('');
  return '<div class="rack">' + rows + '<div class="rack-axis"><div></div><div><span>01</span><span>10</span><span>20</span><span>30</span><span>' + MAX_COL + '</span></div></div></div>';
}

function showDetail(p) {
  var hero = p.loc
    ? '<div class="loc-hero ' + lnCls(p) + '"><div class="lbl">Location</div><div class="big">' + esc(p.loc) + '</div><div class="sub">ชั้น L' + p.lv + ' · ช่องที่ ' + p.col + '</div></div>' + rackHtml(p.loc)
    : '<div class="loc-hero none"><div class="lbl">Location</div><div class="big" style="font-size:30px">รอจัดพื้นที่</div><div class="sub">ยังไม่มีช่องจัดเก็บ</div></div>';
  openModal(hero +
    '<div class="kv">' +
      '<div class="k">Child Code</div><div class="v">' + esc(p.code) + '<button class="copy" data-copy="' + esc(p.code) + '">คัดลอก</button></div>' +
      '<div class="k">Specification</div><div class="v">' + esc(p.spec || '-') + (p.spec ? '<button class="copy" data-copy="' + esc(p.spec) + '">คัดลอก</button>' : '') + '</div>' +
      '<div class="k">ชื่อ</div><div class="v">' + esc(p.name) + '</div>' +
      '<div class="k">Line ผลิต</div><div class="v"><span class="tag ' + lnCls(p) + '">Line ' + esc(p.line) + '</span></div>' +
      '<div class="k">กลุ่ม Part</div><div class="v">' + esc(p.grp) + '</div>' +
      '<div class="k">Data Status</div><div class="v">' + esc(p.st) + '</div>' +
      (p.note ? '<div class="k">หมายเหตุ</div><div class="v" style="color:var(--amber)">⚠️ ' + esc(p.note) + '</div>' : '') +
    '</div>' +
    '<div class="actions">' + (p.loc ? '<button class="btn" id="dMap">🗂 ดูในผัง</button>' : '') +
      '<button class="btn primary" id="dScan">📷 สแกนต่อ</button></div>' +
    '<div class="actions"><button class="btn" id="dQr">🔗 QR ลิงก์</button><button class="btn" id="dClose">ปิด</button></div>');
  Array.prototype.forEach.call(document.querySelectorAll('[data-copy]'), function (b) {
    b.onclick = function () { copy(b.getAttribute('data-copy')); };
  });
  $('dClose').onclick = closeModal;
  $('dQr').onclick = function () {
    var key = p.loc || p.code;
    showQr(APP_URL + '#q=' + encodeURIComponent(key), p.loc || p.code, (p.loc ? p.code + ' · ' : '') + p.name,
      'สแกนแล้วเปิดหน้าเว็บที่ ' + (p.loc ? 'Location ' + p.loc : 'รหัส ' + p.code) + ' ทันที — พิมพ์แปะที่ช่องชั้นวางได้');
  };
  $('dScan').onclick = function () { closeModal(); openCameraScanner(function (t) { handleScan(t, 'cam'); }); };
  if ($('dMap')) $('dMap').onclick = function () { closeModal(); S.mapLv = p.lv; S.mark = p.loc; setTab('map'); scrollToSlot(); };
}

function renderMap() {
  var lv = S.mapLv, n = LEVELS[lv] || 0, slots = '';
  for (var c = 1; c <= n; c++) {
    var loc = 'H9-L' + lv + '-' + ('0' + c).slice(-2), p = byLoc[loc];
    slots += p
      ? '<button class="slot ' + lnCls(p) + (loc === S.mark ? ' on' : '') + '" data-i="' + p.i + '" id="s-' + loc + '"><div class="no">' + esc(p.short) + '</div><div class="c">' + esc(p.code) + '</div><div class="n">' + esc(p.name) + '</div></button>'
      : '<div class="slot empty"><div class="no">' + esc(loc.replace(/^H9-/, '')) + '</div><div class="n">ว่าง</div></div>';
  }
  var waiting = PARTS.filter(function (p) { return !p.loc; }).length;
  $('main').innerHTML =
    '<div class="legend"><span><i style="background:var(--l4)"></i>Line 4</span><span><i style="background:var(--l5)"></i>Line 5</span><span><i style="background:var(--l1)"></i>Line 1</span>' +
      '<span>· แตะแถวในผังเพื่อเลือกชั้น</span></div>' +
    rackHtml(S.mark, true) +
    '<div class="lv-pick">' + LV_LIST.map(function (l) { return '<button data-pl="' + l + '" class="' + (l === lv ? 'active' : '') + '">L' + l + '</button>'; }).join('') + '</div>' +
    '<div class="count"><span>ชั้น L' + lv + ' · ' + n + ' ช่อง</span>' + (waiting ? '<a href="#" id="waitLink">รอจัดพื้นที่ ' + waiting + ' รายการ ›</a>' : '') + '</div>' +
    '<div class="slots">' + slots + '</div>';
  if ($('waitLink')) $('waitLink').onclick = function (e) { e.preventDefault(); S.q = 'รอจัดพื้นที่'; S.scan = null; S.line = 'all'; S.grp = 'all'; setTab('search'); };
  function pick(l) { S.mapLv = l; renderMap(); }
  Array.prototype.forEach.call(document.querySelectorAll('[data-pl]'), function (b) { b.onclick = function () { pick(+b.getAttribute('data-pl')); }; });
  Array.prototype.forEach.call(document.querySelectorAll('[data-lv]'), function (b) { b.onclick = function () { pick(+b.getAttribute('data-lv')); }; });
  Array.prototype.forEach.call(document.querySelectorAll('.slot[data-i]'), function (b) { b.onclick = function () { showDetail(PARTS[+b.getAttribute('data-i')]); }; });
}
function scrollToSlot() {
  setTimeout(function () { var el = S.mark && $('s-' + S.mark); if (el) el.scrollIntoView({ block: 'center', behavior: 'smooth' }); }, 60);
}

// ======================= ไล่เช็ค Part ↔ Location =======================
// สแกนป้าย Location → สแกนป้าย Part ในช่องนั้น → ตรง ✅ / ไม่ตรง ❌ · ผลล่าสุดของแต่ละช่องเก็บในเครื่อง (localStorage)
var CK = (function () {
  try { var d = JSON.parse(store('h9pl_check') || 'null'); if (d && d.rec) return d; } catch (e) { /* ข้อมูลเสีย → เริ่มใหม่ */ }
  return { cur: '', rec: {}, last: null, by: '' };
})();
var ALL_LOCS = PARTS.filter(function (p) { return p.loc; }).sort(locOrder).map(function (p) { return p.loc; });
var CK_LBL = { ok: '✅ ตรง', ng: '❌ ไม่ตรง', unk: '❓ ไม่พบในรายการ', empty: '⬜ ช่องว่าง' };
function saveCheck() { store('h9pl_check', JSON.stringify(CK)); }
function locExists(l) { var m = l && l.match(/^H9-L(\d+)-(\d+)$/); return !!(m && LEVELS[+m[1]] && +m[2] <= LEVELS[+m[1]]); }
// หา Location จากข้อความที่สแกน (ป้าย Location / QR ลิงก์ #q= / คำใดคำหนึ่ง)
function scanLoc(text) {
  var t = String(text || ''), lq = t.match(/#q=([^&\s]+)/);
  if (lq) { try { t = decodeURIComponent(lq[1]); } catch (e) { t = lq[1]; } }
  var l = normLoc(t); if (locExists(l)) return l;
  var toks = t.toUpperCase().split(/[^A-Z0-9\-]+/);
  for (var i = 0; i < toks.length; i++) { l = normLoc(toks[i]); if (locExists(l)) return l; }
  return '';
}
function nextLoc(loc) {
  var i = ALL_LOCS.indexOf(loc);
  return i >= 0 && i < ALL_LOCS.length - 1 ? ALL_LOCS[i + 1] : '';
}
function ckStats() {
  var st = { ok: 0, ng: 0, unk: 0, empty: 0, done: 0 };
  ALL_LOCS.forEach(function (l) { var r = CK.rec[l]; if (r) { st[r.st]++; st.done++; } });
  return st;
}
function partTxt(p) { return p.code + ' · ' + p.name; }

// ผลการสแกน 1 ครั้ง → คืน class สำหรับไฟกระพริบในกล้อง (ok / err / loc)
function checkScan(text) {
  text = String(text || '').trim(); if (!text) return '';
  var loc = scanLoc(text);
  if (loc) {
    CK.cur = loc; CK.last = { type: 'loc', loc: loc }; saveCheck(); beep(true);
    return 'loc';
  }
  if (!CK.cur) { CK.last = { type: 'need', text: text }; saveCheck(); beep(false); return 'err'; }
  var r = resolveScan(text), exp = byLoc[CK.cur], st;
  if (!r) st = 'unk';
  else if (exp && r.parts.indexOf(exp) >= 0) st = 'ok';
  else st = 'ng';
  var rec = { st: st, text: text.slice(0, 120), at: Date.now(),
    got: r ? r.parts.map(function (p) { return p.i; }) : [] };
  CK.rec[CK.cur] = rec;
  CK.last = { type: 'res', loc: CK.cur };
  saveCheck(); beep(st === 'ok');
  return st === 'ok' ? 'ok' : 'err';
}
function markEmpty() {
  if (!CK.cur) return;
  CK.rec[CK.cur] = { st: 'empty', text: '', at: Date.now(), got: [] };
  CK.last = { type: 'res', loc: CK.cur }; saveCheck();
}

function ckCurHtml() {
  if (!CK.cur) return '<div class="ck-cur none"><div class="lbl">Location ที่กำลังเช็ค</div><div class="big">—</div>' +
    '<div class="sub">สแกนป้าย Location หรือพิมพ์ เช่น <b>L3-05</b></div></div>';
  var exp = byLoc[CK.cur], r = CK.rec[CK.cur];
  return '<div class="ck-cur ' + (exp ? lnCls(exp) : '') + '"><div class="lbl">Location ที่กำลังเช็ค' +
      (r ? ' · <span class="ck-chip ' + r.st + '">' + CK_LBL[r.st] + '</span>' : '') + '</div>' +
    '<div class="big">' + esc(CK.cur) + '</div>' +
    '<div class="sub">ควรเป็น: ' + (exp ? '<b>' + esc(exp.code) + '</b> · ' + esc(exp.name) : '<b>ไม่มี Part ในรายการ</b>') + '</div></div>';
}
function ckLastHtml() {
  var L = CK.last;
  if (!L) return '<div class="ck-res idle">① สแกนป้าย <b>Location</b> → ② สแกนป้าย <b>Part</b> ในช่องนั้น</div>';
  if (L.type === 'need') return '<div class="ck-res err"><div class="h">⚠️ ยังไม่ได้เลือก Location</div>สแกนป้าย Location ก่อน แล้วค่อยสแกน Part</div>';
  if (L.type === 'loc') return '<div class="ck-res loc"><div class="h">📍 ' + esc(L.loc) + '</div>สแกนป้าย Part ที่อยู่ในช่องนี้ได้เลย</div>';
  var r = CK.rec[L.loc], exp = byLoc[L.loc]; if (!r) return '';
  if (r.st === 'empty') return '<div class="ck-res warn"><div class="h">⬜ บันทึกว่าช่องว่าง</div>' + esc(L.loc) + (exp ? ' (ควรมี ' + esc(exp.code) + ')' : '') + '</div>';
  if (r.st === 'ok') return '<div class="ck-res ok"><div class="h">✅ ตรง Location</div>' + esc(L.loc) + ' · ' + esc(exp ? partTxt(exp) : '') + '</div>';
  var got = r.got.map(function (i) { return PARTS[i]; });
  return '<div class="ck-res err"><div class="h">' + (r.st === 'unk' ? '❓ ไม่พบ Part นี้ในรายการ' : '❌ ไม่ตรง Location') + '</div>' +
    (got.length ? got.map(function (p) {
      return 'สแกนได้ <b>' + esc(p.code) + '</b> · ' + esc(p.name) + '<br>→ ควรอยู่ที่ <b>' + esc(p.loc || 'รอจัดพื้นที่') + '</b>';
    }).join('<hr>') : 'สแกนได้: <b>' + esc(r.text) + '</b>') +
    '<div class="exp">ช่อง ' + esc(L.loc) + ' ควรเป็น: ' + (exp ? '<b>' + esc(exp.code) + '</b> · ' + esc(exp.name) : 'ไม่มี Part ในรายการ') + '</div></div>';
}
function ckCamStatus() {
  var nx = nextLoc(CK.cur), st = ckStats();
  return ckCurHtml() + ckLastHtml() +
    '<div class="center small muted">เช็คแล้ว ' + st.done + '/' + ALL_LOCS.length + ' · ✅ ' + st.ok + ' · ❌ ' + (st.ng + st.unk) + ' · ⬜ ' + st.empty + '</div>' +
    (CK.cur ? '<div class="actions"><button class="btn" id="ckEmptyC">⬜ ช่องว่าง</button>' +
      (nx ? '<button class="btn" id="ckNextC">➡ ' + esc(nx.replace(/^H9-/, '')) + '</button>' : '') + '</div>' : '');
}
function bindCamStatus() {
  if ($('ckEmptyC')) $('ckEmptyC').onclick = function () { markEmpty(); refreshCam(); };
  if ($('ckNextC')) $('ckNextC').onclick = function () { CK.cur = nextLoc(CK.cur); CK.last = { type: 'loc', loc: CK.cur }; saveCheck(); refreshCam(); };
}
function refreshCam() { if ($('camStatus')) { $('camStatus').innerHTML = ckCamStatus(); bindCamStatus(); } }

function ckRackHtml() {
  var rows = LV_LIST.map(function (lv) {
    var cells = '';
    for (var c = 1; c <= MAX_COL; c++) {
      var loc = 'H9-L' + lv + '-' + ('0' + c).slice(-2), r = CK.rec[loc];
      var cls = c > LEVELS[lv] ? 'void' : r ? 'ck-' + r.st : byLoc[loc] ? '' : 'void';
      cells += '<div class="rc ' + cls + (loc === CK.cur ? ' on' : '') + '"' + (cls === 'void' ? '' : ' data-ck="' + loc + '"') + '></div>';
    }
    return '<div class="rack-row"><div class="rack-lv">L' + lv + '</div><div class="rack-cells">' + cells + '</div></div>';
  }).join('');
  return '<div class="rack ck">' + rows + '<div class="rack-axis"><div></div><div><span>01</span><span>10</span><span>20</span><span>30</span><span>' + MAX_COL + '</span></div></div></div>';
}

function renderCheck() {
  var st = ckStats(), pct = ALL_LOCS.length ? Math.round(st.done * 100 / ALL_LOCS.length) : 0, nx = nextLoc(CK.cur);
  var f = S.ckf || 'ng';
  var list = f === 'todo'
    ? ALL_LOCS.filter(function (l) { return !CK.rec[l]; })
    : Object.keys(CK.rec).filter(function (l) { var r = CK.rec[l]; return f === 'all' || r.st === 'ng' || r.st === 'unk' || (f === 'ng' && r.st === 'empty'); })
        .sort(function (a, b) { return CK.rec[b].at - CK.rec[a].at; });
  $('main').innerHTML =
    '<div class="ck-prog"><div class="ck-bar"><i class="ok" style="width:' + (st.ok * 100 / (ALL_LOCS.length || 1)) + '%"></i>' +
      '<i class="ng" style="width:' + ((st.ng + st.unk) * 100 / (ALL_LOCS.length || 1)) + '%"></i>' +
      '<i class="empty" style="width:' + (st.empty * 100 / (ALL_LOCS.length || 1)) + '%"></i></div>' +
      '<div class="ck-nums"><span>เช็คแล้ว <b>' + st.done + '</b>/' + ALL_LOCS.length + ' (' + pct + '%)</span>' +
      '<span>✅ ' + st.ok + '</span><span>❌ ' + (st.ng + st.unk) + '</span><span>⬜ ' + st.empty + '</span></div></div>' +
    ckCurHtml() +
    '<div class="search ck-search"><div class="search-wrap"><input id="ckq" type="text" enterkeyhint="go" autocomplete="off" autocapitalize="characters" spellcheck="false" placeholder="ยิงสแกน / พิมพ์ แล้ว Enter"></div>' +
      '<button class="scan-btn" id="ckCam">📷 สแกน</button></div>' +
    ckLastHtml() +
    (CK.cur ? '<div class="actions ck-act"><button class="btn" id="ckEmpty">⬜ ช่องว่าง</button>' +
      (nx ? '<button class="btn primary" id="ckNext">➡ ช่องถัดไป ' + esc(nx.replace(/^H9-/, '')) + '</button>' : '') + '</div>' : '') +
    '<div class="legend"><span><i style="background:var(--green)"></i>ตรง</span><span><i style="background:var(--red)"></i>ไม่ตรง</span>' +
      '<span><i style="background:#f2b544"></i>ว่าง</span><span><i style="background:var(--line)"></i>ยังไม่เช็ค</span><span>· แตะช่องเพื่อเลือก</span></div>' +
    ckRackHtml() +
    '<div class="filters">' +
      '<button class="fchip' + (f === 'ng' ? ' active' : '') + '" data-ckf="ng">❌ ไม่ตรง / ว่าง</button>' +
      '<button class="fchip' + (f === 'all' ? ' active' : '') + '" data-ckf="all">ผลทั้งหมด ' + st.done + '</button>' +
      '<button class="fchip' + (f === 'todo' ? ' active' : '') + '" data-ckf="todo">ยังไม่เช็ค ' + (ALL_LOCS.length - st.done) + '</button></div>' +
    '<div id="ckList">' + (list.length ? list.slice(0, 150).map(ckRow).join('') : '<div class="msg idle">ไม่มีรายการ</div>') +
      (list.length > 150 ? '<div class="count">แสดง 150 รายการแรก จาก ' + list.length + '</div>' : '') + '</div>' +
    '<div class="actions"><button class="btn" id="ckCsv">⬇️ Export (Excel/CSV)</button><button class="btn" id="ckReset">🗑 เริ่มเช็คใหม่</button></div>';

  var q = $('ckq');
  q.onkeydown = function (e) {
    if (e.key !== 'Enter') return;
    e.preventDefault(); var v = q.value.trim(); if (!v) return;
    checkScan(v); renderCheck(); if ($('ckq')) $('ckq').focus();
  };
  $('ckCam').onclick = function () {
    openCameraScanner(null, { title: 'ไล่เช็ค (สแกนต่อเนื่อง)', status: ckCamStatus,
      onEach: function (t) { var c = checkScan(t); setTimeout(bindCamStatus, 0); return c; } });
    onModalClose = (function (stopCam) { return function () { stopCam && stopCam(); renderCheck(); }; })(onModalClose);
    bindCamStatus();
  };
  if ($('ckEmpty')) $('ckEmpty').onclick = function () { markEmpty(); renderCheck(); };
  if ($('ckNext')) $('ckNext').onclick = function () { CK.cur = nx; CK.last = { type: 'loc', loc: nx }; saveCheck(); renderCheck(); };
  Array.prototype.forEach.call(document.querySelectorAll('[data-ck]'), function (b) {
    b.onclick = function () { CK.cur = b.getAttribute('data-ck'); CK.last = { type: 'loc', loc: CK.cur }; saveCheck(); renderCheck(); window.scrollTo({ top: 0, behavior: 'smooth' }); };
  });
  Array.prototype.forEach.call(document.querySelectorAll('[data-ckf]'), function (b) {
    b.onclick = function () { S.ckf = b.getAttribute('data-ckf'); renderCheck(); };
  });
  Array.prototype.forEach.call(document.querySelectorAll('[data-ckrow]'), function (b) {
    b.onclick = function () { CK.cur = b.getAttribute('data-ckrow'); CK.last = CK.rec[CK.cur] ? { type: 'res', loc: CK.cur } : { type: 'loc', loc: CK.cur }; saveCheck(); renderCheck(); window.scrollTo({ top: 0, behavior: 'smooth' }); };
  });
  $('ckCsv').onclick = exportCheck;
  $('ckReset').onclick = function () {
    if (!confirm('ล้างผลการเช็คทั้งหมดในเครื่องนี้ แล้วเริ่มใหม่?\n(แนะนำให้ Export เก็บไว้ก่อน)')) return;
    CK = { cur: '', rec: {}, last: null, by: CK.by }; saveCheck(); renderCheck();
  };
}
function fmtTime(t) {
  var d = new Date(t), z = function (n) { return ('0' + n).slice(-2); };
  return z(d.getDate()) + '/' + z(d.getMonth() + 1) + ' ' + z(d.getHours()) + ':' + z(d.getMinutes());
}
function ckRow(loc) {
  var r = CK.rec[loc], exp = byLoc[loc];
  var got = r ? r.got.map(function (i) { return PARTS[i]; }) : [];
  return '<button class="ck-row ' + (r ? r.st : 'todo') + '" data-ckrow="' + esc(loc) + '">' +
    '<div class="l">' + esc(loc.replace(/^H9-/, '')) + '</div><div class="m">' +
    '<div>' + (r ? '<span class="ck-chip ' + r.st + '">' + CK_LBL[r.st] + '</span> <span class="muted small">' + fmtTime(r.at) + '</span>' : '<span class="muted">ยังไม่เช็ค</span>') + '</div>' +
    '<div class="small">ควรเป็น: <b>' + esc(exp ? exp.code : '-') + '</b> ' + esc(exp ? exp.name : '') + '</div>' +
    (r && (r.st === 'ng' || r.st === 'unk') ? '<div class="small bad">สแกนได้: ' + (got.length ? got.map(function (p) {
      return '<b>' + esc(p.code) + '</b> (ควรอยู่ ' + esc(p.loc || 'รอจัดพื้นที่') + ')';
    }).join(', ') : esc(r.text)) + '</div>' : '') +
    '</div></button>';
}
function exportCheck() {
  var head = ['Location', 'ผลการเช็ค', 'Part ที่ควรอยู่ (Code)', 'Spec', 'ชื่อ Part', 'Line', 'สแกนได้ (Code)', 'ชื่อที่สแกนได้', 'Part ที่สแกนได้ควรอยู่ที่', 'ข้อความที่สแกน', 'เวลา'];
  var rows = ALL_LOCS.concat(Object.keys(CK.rec).filter(function (l) { return ALL_LOCS.indexOf(l) < 0; })).map(function (loc) {
    var r = CK.rec[loc], exp = byLoc[loc], got = r ? r.got.map(function (i) { return PARTS[i]; }) : [];
    return [loc, r ? CK_LBL[r.st].replace(/^\S+\s/, '') : 'ยังไม่เช็ค', exp ? exp.code : '', exp ? exp.spec : '', exp ? exp.name : '', exp ? exp.line : '',
      got.map(function (p) { return p.code; }).join(' / '), got.map(function (p) { return p.name; }).join(' / '),
      got.map(function (p) { return p.loc || 'รอจัดพื้นที่'; }).join(' / '), r ? r.text : '', r ? new Date(r.at).toLocaleString('th-TH') : ''];
  });
  var csv = '﻿' + [head].concat(rows).map(function (row) {
    return row.map(function (v) { v = String(v == null ? '' : v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(',');
  }).join('\r\n');
  var d = new Date(), z = function (n) { return ('0' + n).slice(-2); };
  var a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  a.download = 'H9-check-' + d.getFullYear() + z(d.getMonth() + 1) + z(d.getDate()) + '-' + z(d.getHours()) + z(d.getMinutes()) + '.csv';
  document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  toast('Export ' + rows.length + ' แถว');
}

// ======================= QR ลิงก์เข้าหน้าเว็บ =======================
// APP_URL = ที่อยู่หน้าเว็บนี้ (ตัด ?v= และ #… ออก) · ลิงก์ Part: APP_URL#q=H9-L3-05
var APP_URL = location.href.replace(/[?#].*$/, '');
var QRGEN_URL = 'js/vendor/qrcode.js?v=1';
function makeQr(text) {
  var q = window.qrcode(0, 'M'); q.addData(text, 'Byte'); q.make(); return q;
}
// วาด QR + ข้อความใต้ภาพลง canvas (ใช้ทั้งแสดง บันทึกรูป และพิมพ์)
function qrCanvas(text, title, sub) {
  var q = makeQr(text), n = q.getModuleCount(), cell = Math.max(6, Math.floor(600 / (n + 8))), m = cell * 4;
  var size = n * cell + m * 2, lines = [title, sub].filter(Boolean), extra = lines.length ? 30 + lines.length * 44 : 0;
  var c = document.createElement('canvas'); c.width = size; c.height = size + extra;
  var g = c.getContext('2d');
  g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.fillStyle = '#000';
  for (var r = 0; r < n; r++) for (var k = 0; k < n; k++) if (q.isDark(r, k)) g.fillRect(m + k * cell, m + r * cell, cell, cell);
  g.textAlign = 'center';
  lines.forEach(function (t, i) {
    g.font = (i === 0 ? '700 40px' : '400 28px') + ' Sarabun, sans-serif';
    var s = String(t); while (s.length > 4 && g.measureText(s).width > size - 30) s = s.slice(0, -2);
    if (s !== String(t)) s += '…';
    g.fillText(s, size / 2, size + 10 + (i + 1) * 40);
  });
  return c;
}
function showQr(url, title, sub, hint) {
  if (hint === undefined) { hint = sub; sub = ''; }
  loadScriptOnce(QRGEN_URL).then(function () {
    var c = qrCanvas(url, title, sub), png = c.toDataURL('image/png');
    var fname = 'QR-' + String(title).replace(/[^\w-]+/g, '_') + '.png';
    openModal('<h3>🔗 QR ลิงก์ ' + esc(title) + '</h3>' +
      '<div class="qr-box"><img src="' + png + '" alt="QR"></div>' +
      '<p class="center small muted">' + esc(hint) + '</p>' +
      '<div class="qr-url">' + esc(url) + '</div>' +
      '<div class="actions"><button class="btn" id="qCopy">📋 คัดลอกลิงก์</button>' +
        (navigator.share ? '<button class="btn" id="qShare">📤 แชร์</button>' : '') + '</div>' +
      '<div class="actions"><a class="btn center" id="qSave" href="' + png + '" download="' + esc(fname) + '">💾 บันทึกรูป</a>' +
        '<button class="btn" id="qPrint">🖨 พิมพ์</button></div>' +
      '<div class="actions"><button class="btn" id="qClose">ปิด</button></div>');
    $('qClose').onclick = closeModal;
    $('qCopy').onclick = function () { copy(url); };
    if ($('qShare')) $('qShare').onclick = function () {
      navigator.share({ title: 'H9 Part Location', text: title, url: url }).catch(function () {});
    };
    $('qPrint').onclick = function () {
      var w = window.open('', '_blank');
      if (!w) { toast('เบราว์เซอร์บล็อกหน้าต่างใหม่ — ใช้ปุ่มบันทึกรูปแทน', 3000); return; }
      w.document.write('<!doctype html><meta charset="utf-8"><title>' + esc(title) + '</title>' +
        '<style>body{margin:0;display:flex;justify-content:center;align-items:flex-start;padding:20px}img{width:80mm;max-width:100%}</style>' +
        '<img src="' + png + '" onload="setTimeout(function(){print()},200)">');
      w.document.close();
    };
  }, function () { toast('โหลดตัวสร้าง QR ไม่ได้', 3000); });
}

// ======================= สแกนด้วยกล้อง =======================
// ลำดับตัวอ่าน: 1) ตัวอ่านในเบราว์เซอร์ (Android Chrome) → 2) zxing (WebAssembly, iPhone) → 3) jsQR · เสริมด้วยตัวกู้ QR มุมขาด (js/partialqr.js)
var ZXING_URL = 'https://cdn.jsdelivr.net/npm/barcode-detector@2.2.2/dist/es/pure.min.js';
var JSQR_URL = 'js/vendor/jsqr.js?v=2'; // jsQR ในเครื่อง (เปิดส่วนภายในให้ตัวกู้ QR ใช้)
var PARTIALQR_URL = 'js/partialqr.js?v=1';
var SCAN_FORMATS = ['qr_code', 'code_128', 'code_39', 'ean_13', 'data_matrix'];
var CROP_RATIO = 0.72;
var IS_IOS = /iPhone|iPad|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

function loadScriptOnce(src) {
  loadScriptOnce.c = loadScriptOnce.c || {};
  if (!loadScriptOnce.c[src]) loadScriptOnce.c[src] = new Promise(function (ok, bad) {
    var s = document.createElement('script'); s.src = src; s.onload = ok;
    s.onerror = function () { delete loadScriptOnce.c[src]; bad(new Error('load ' + src)); };
    document.head.appendChild(s);
  });
  return loadScriptOnce.c[src];
}
function firstValue(r) { return r && r.length ? r[0].rawValue : null; }
function makeNative() {
  if (!(window.BarcodeDetector && window.BarcodeDetector.getSupportedFormats)) return Promise.reject(new Error('no native'));
  return window.BarcodeDetector.getSupportedFormats().then(function (f) {
    if (f.indexOf('qr_code') < 0) throw new Error('no qr');
    var d = new window.BarcodeDetector({ formats: SCAN_FORMATS.filter(function (x) { return f.indexOf(x) >= 0; }) });
    return { src: IS_IOS ? 'full' : 'video', detect: function (s) { return d.detect(s).then(firstValue); } };
  });
}
function makeZxing() {
  return import(ZXING_URL).then(function (m) {
    var d = new m.BarcodeDetector({ formats: SCAN_FORMATS });
    return { src: 'crop', detect: function (s) { return d.detect(s).then(firstValue); } };
  });
}
function makeJsqr() {
  return loadScriptOnce(JSQR_URL).then(function () {
    return { src: 'crop', detect: function (s) {
      var g = s.getContext('2d', { willReadFrequently: true });
      var r = window.jsQR(g.getImageData(0, 0, s.width, s.height).data, s.width, s.height, { inversionAttempts: 'attemptBoth' });
      return Promise.resolve(r && r.data ? r.data : null);
    } };
  });
}
// ตัวกู้ QR มุมขาด/ไม่ชัด (เหมือน pda-location-scan): หาสี่เหลี่ยมมุมที่เหลือ 2 มุม → คำนวณมุมที่หาย → ให้ Error Correction ของ QR กู้ข้อมูล
// ใช้ CPU มากกว่า → เว้นระยะตามเวลาที่ใช้จริง (อย่างน้อย 250 ms) และย่อภาพไม่เกิน 480 px
function makePartial() {
  return loadScriptOnce(JSQR_URL).then(function () { return loadScriptOnce(PARTIALQR_URL); }).then(function () {
    if (!window.jsQR || !window.jsQR._i || !window.decodePartialQR) throw new Error('no partial');
    var last = 0, gap = 250, cv = document.createElement('canvas');
    return { src: 'crop', detect: function (src) {
      var now = Date.now(); if (now - last < gap) return Promise.resolve(null); last = now;
      var k = Math.min(1, 480 / Math.max(src.width, src.height)), w = Math.round(src.width * k), h = Math.round(src.height * k);
      cv.width = w; cv.height = h;
      var g = cv.getContext('2d', { willReadFrequently: true }); g.drawImage(src, 0, 0, w, h);
      var r = window.decodePartialQR(window.jsQR, g.getImageData(0, 0, w, h).data, w, h);
      gap = Math.max(250, (Date.now() - now) * 3); // เครื่องช้า → เว้นนานขึ้น กล้องไม่หน่วง
      return Promise.resolve(r && r.data ? r.data : null);
    } };
  });
}
// Android: ตัวอ่านในเบราว์เซอร์ + ตัวกู้ QR · iPhone/อื่นๆ: zxing (สำรอง jsQR) + ตัวอ่าน Apple ถ้ามี + ตัวกู้ QR — สลับใช้ทุกเฟรม
function getDetectors() {
  if (getDetectors.p) return getDetectors.p;
  getDetectors.p = makeNative().catch(function () { return null; }).then(function (nat) {
    var part = makePartial().catch(function () { return null; });
    if (nat && !IS_IOS) return part.then(function (p) { return [nat, p].filter(Boolean); });
    return Promise.all([makeZxing().catch(function () { return makeJsqr(); }).catch(function () { return null; }), part]).then(function (zz) {
      var list = [nat, zz[0], zz[1]].filter(Boolean);
      if (!list.length) throw new Error('โหลดตัวอ่าน QR ไม่ได้ (ต้องต่ออินเทอร์เน็ต)');
      return list;
    });
  });
  getDetectors.p.catch(function () { getDetectors.p = null; });
  return getDetectors.p;
}

var camCanvas = null, camCropSide = 0;
// สลับวิธีเตรียมภาพทุกเฟรม: 0 = กรอบกลางเต็มความละเอียด · 1 = ย่อครึ่ง (ภาพเบลอ) · 2 = ย่อครึ่ง + เร่งความต่างสี (แสงน้อย) · 3 = ทั้งภาพ
function grabVariant(src, v) {
  if (v === 3) return grabFrame(src, true);
  var c = grabFrame(src, false, v === 0 ? 1 : 0.5);
  if (c && v === 2) stretchContrast(c);
  return c;
}
function stretchContrast(c) {
  var g = c.getContext('2d', { willReadFrequently: true }), img = g.getImageData(0, 0, c.width, c.height), d = img.data;
  var hist = new Uint32Array(256), n = d.length / 4, i, y;
  for (i = 0; i < d.length; i += 4) { y = (d[i] * 77 + d[i + 1] * 150 + d[i + 2] * 29) >> 8; d[i] = y; hist[y]++; }
  var lo = 0, hi = 255, acc = 0, cut = n * 0.02;
  while (lo < 255 && (acc += hist[lo]) < cut) lo++;
  acc = 0; while (hi > 0 && (acc += hist[hi]) < cut) hi--;
  if (hi - lo < 8) return;
  var lut = new Uint8ClampedArray(256), k = 255 / (hi - lo);
  for (i = 0; i < 256; i++) lut[i] = (i - lo) * k;
  for (i = 0; i < d.length; i += 4) { y = lut[d[i]]; d[i] = d[i + 1] = d[i + 2] = y; }
  g.putImageData(img, 0, 0);
}
function grabFrame(src, full, scale) {
  var w = src.videoWidth || src.width, h = src.videoHeight || src.height;
  if (!w || !h) return null;
  var sx = 0, sy = 0, sw = w, sh = h, ow, oh;
  if (full) { var k = Math.min(1, 1280 / Math.max(w, h)); ow = Math.round(w * k); oh = Math.round(h * k); }
  else {
    var side = camCropSide || Math.round(Math.min(w, h) * CROP_RATIO);
    sx = Math.round((w - side) / 2); sy = Math.round((h - side) / 2); sw = sh = side;
    ow = oh = Math.round(Math.min(side, 1000) * (scale || 1));
  }
  camCanvas = camCanvas || document.createElement('canvas');
  camCanvas.width = ow; camCanvas.height = oh;
  camCanvas.getContext('2d', { willReadFrequently: true }).drawImage(src, sx, sy, sw, sh, 0, 0, ow, oh);
  return camCanvas;
}

// opts.onEach(text): โหมดต่อเนื่อง (ไม่ปิดกล้องหลังอ่านได้) · opts.status(): HTML แสดงใต้กล้อง
function openCameraScanner(onResult, opts) {
  opts = opts || {};
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    toast('เบราว์เซอร์นี้เปิดกล้องไม่ได้ (ต้องเปิดผ่าน https)', 3500); return;
  }
  openModal('<h3>📷 ' + esc(opts.title || 'สแกน QR / บาร์โค้ด') + '</h3>' +
    '<div class="cam-wrap' + (opts.onEach ? ' short' : '') + '" id="camWrap"><video id="camV" playsinline muted autoplay></video><div class="cam-frame" id="camFrame"></div>' +
      '<div class="cam-flash" id="camFlash"></div></div>' +
    (opts.status ? '<div id="camStatus">' + opts.status() + '</div>' : '') +
    '<div class="cam-tools" id="camTools"></div>' +
    '<p class="center small" id="camHint">กำลังเปิดกล้อง…</p>' +
    '<div class="actions"><button class="btn" id="camClose">ปิด</button></div>');
  var video = $('camV'), stream = null, done = false;
  function stop() { done = true; if (stream) stream.getTracks().forEach(function (t) { t.stop(); }); }
  onModalClose = stop;
  $('camClose').onclick = closeModal;

  var cam = navigator.mediaDevices.getUserMedia({
    video: { facingMode: { ideal: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1080 } }, audio: false
  });
  Promise.all([cam, getDetectors()]).then(function (res) {
    stream = res[0];
    var engines = res[1].map(function (e) { e.frames = 0; e.fails = 0; return e; });
    if (done) { stop(); return; }
    var track = stream.getVideoTracks()[0];
    video.srcObject = stream;
    return video.play().then(function () {
      sizeFrame(); setupCamTools(track);
      var startAt = Date.now(), n = 0;
      $('camHint').innerHTML = 'เล็งป้าย QR / บาร์โค้ด ให้อยู่ในกรอบ' + (IS_IOS ? ' · <b>ถือห่างประมาณ 20–30 ซม.</b>' : '');
      function next() {
        if (done) return;
        if (video.requestVideoFrameCallback) video.requestVideoFrameCallback(function () { tick(); });
        else requestAnimationFrame(tick);
      }
      function tick() {
        if (done) return;
        if (Date.now() - startAt > 5000 && $('camHint')) $('camHint').innerHTML = 'ยังอ่านไม่ได้: <b>ถือให้นิ่ง</b> · ถอยออกเล็กน้อยให้ภาพชัด' + ($('camTorch') ? ' · ที่มืดเปิด <b>🔦 ไฟฉาย</b>' : '');
        var e = engines[n++ % engines.length];
        var src = e.src === 'video' ? video : e.src === 'full' ? grabFrame(video, true) : grabVariant(video, e.frames % 4);
        if (!src) { next(); return; }
        e.frames++;
        e.detect(src).then(function (text) {
          e.fails = 0;
          if (done) return;
          if (text && opts.onEach) { onCont(text); return; }
          if (text) { closeModal(); onResult(text); return; }
          next();
        }, function () {
          e.fails++;
          if (e.fails >= 5 && engines.length > 1) engines.splice(engines.indexOf(e), 1);
          setTimeout(next, 50);
        });
      }
      // ต่อเนื่อง: ป้ายเดิมที่ยังอยู่หน้ากล้องไม่นับซ้ำ (ต้องหายไปอย่างน้อย 2.5 วินาที) · พักสั้นๆ หลังอ่านได้
      var lastText = '', lastAt = 0;
      function onCont(text) {
        var now = Date.now();
        if (text === lastText && now - lastAt < 2500) { lastAt = now; setTimeout(next, 120); return; }
        lastText = text; lastAt = now; startAt = now;
        var cls = opts.onEach(text);
        if ($('camStatus') && opts.status) $('camStatus').innerHTML = opts.status();
        var f = $('camFlash'); if (f) { f.className = 'cam-flash ' + (cls || 'ok'); void f.offsetWidth; f.classList.add('on'); }
        setTimeout(next, 800);
      }
      next();
    });
  }).catch(function (err) {
    cam.then(function (s) { s.getTracks().forEach(function (t) { t.stop(); }); }, function () {});
    stop();
    var name = err && err.name;
    var msg = name === 'NotAllowedError' ? 'ไม่ได้รับอนุญาตให้ใช้กล้อง — กดอนุญาตกล้องที่แถบที่อยู่ของเบราว์เซอร์แล้วลองใหม่'
      : name === 'NotFoundError' ? 'ไม่พบกล้องในเครื่องนี้'
      : 'เปิดกล้องไม่ได้: ' + ((err && (err.message || err.name)) || err);
    if ($('camHint')) $('camHint').textContent = msg; else toast(msg, 4000);
  });

  function sizeFrame() {
    var f = $('camFrame'), wrap = video.parentNode; if (!f || !wrap) return;
    var W = wrap.clientWidth, H = wrap.clientHeight, vw = video.videoWidth, vh = video.videoHeight;
    var side = Math.round(Math.min(W, H) * CROP_RATIO);
    f.style.width = f.style.height = side + 'px';
    if (vw && vh && W && H) camCropSide = Math.min(Math.round(side / Math.max(W / vw, H / vh)), vw, vh);
  }
  camCropSide = 0;
  video.addEventListener('loadedmetadata', sizeFrame);
  video.addEventListener('resize', sizeFrame);
}

// ไฟฉาย / ซูม / โฟกัสต่อเนื่อง — แสดงเฉพาะที่เครื่องรองรับ
function setupCamTools(track) {
  var box = $('camTools'); if (!box || !track || !track.getCapabilities) return;
  var caps = {};
  try { caps = track.getCapabilities() || {}; } catch (e) { return; }
  if (caps.focusMode && caps.focusMode.indexOf('continuous') >= 0) track.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }).catch(function () {});
  var html = '';
  if (caps.torch) html += '<button id="camTorch">🔦 ไฟฉาย</button>';
  if (caps.zoom && caps.zoom.max > caps.zoom.min) html += '<input type="range" id="camZoom" min="' + caps.zoom.min + '" max="' + Math.min(caps.zoom.max, 5) + '" step="' + (caps.zoom.step || 0.1) + '" value="' + caps.zoom.min + '" aria-label="ซูม">';
  box.innerHTML = html;
  var torch = false;
  if ($('camTorch')) $('camTorch').onclick = function () {
    torch = !torch; var b = this;
    track.applyConstraints({ advanced: [{ torch: torch }] }).then(function () { b.classList.toggle('active', torch); }, function () { torch = !torch; toast('เปิดไฟฉายไม่ได้'); });
  };
  if ($('camZoom')) $('camZoom').oninput = function () { track.applyConstraints({ advanced: [{ zoom: +this.value }] }).catch(function () {}); };
}

// ======================= อัปเดตอัตโนมัติ =======================
// เช็ค version.json แบบไม่ใช้ cache ถ้ามีเวอร์ชันใหม่ → โหลดหน้าใหม่ (ตอนที่ไม่ได้เปิดกล้อง/หน้าต่างค้างอยู่)
function checkUpdate() {
  if (APP_VERSION === 'dev' || !window.fetch) return;
  fetch('version.json?_=' + Date.now(), { cache: 'no-store' }).then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
    if (d && d.version && d.version !== APP_VERSION && $('modal').classList.contains('hidden')) {
      location.replace(location.pathname + '?v=' + d.version + location.hash);
    }
  }).catch(function () {});
}

// ---------- เริ่ม ----------
// ลิงก์ตรง: index.html#q=L3-05 หรือ #q=1005QT00456
var hq = decodeURIComponent((location.hash.match(/[#&]q=([^&]*)/) || [])[1] || '');
render();
if (hq) handleScan(hq, 'link');
window.addEventListener('hashchange', function () {
  var q = (location.hash.match(/[#&]q=([^&]*)/) || [])[1];
  if (q) { closeModal(); handleScan(decodeURIComponent(q), 'link'); }
});
document.addEventListener('visibilitychange', function () { if (!document.hidden) checkUpdate(); });
setInterval(checkUpdate, 5 * 60 * 1000);
checkUpdate();
})();
