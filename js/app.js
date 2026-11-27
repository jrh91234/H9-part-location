// H9 Part Location — ค้นหา/สแกน ว่ารหัส Part อยู่ Location ไหน (ข้อมูลจาก js/data.js)
(function () {
'use strict';
var APP_VERSION = '20261127-0138'; // ตั้งโดย tools/bump-version.py

var PARTS = window.PARTS || [], LEVELS = window.LEVELS || {};
var LV_LIST = Object.keys(LEVELS).map(Number).sort(function (a, b) { return b - a; }); // ชั้นบนสุดก่อน (เหมือนแผนผัง)
var MAX_COL = Math.max.apply(null, LV_LIST.map(function (l) { return LEVELS[l]; }).concat([1]));
var LINES = [4, 5, 1];

// ---------- ดัชนี ----------
// p.loc0 = ช่องตามไฟล์ Excel · p.loc = ช่องจริงตอนนี้ (หลังรวมการเปลี่ยนจากเมนูตั้งค่า)
var byLoc = {}, byCode = {}, bySpec = {};
PARTS.forEach(function (p, i) {
  p.i = i; p.loc0 = p.loc;
  (byCode[p.code] = byCode[p.code] || []).push(p);
  if (p.spec) (bySpec[p.spec] = bySpec[p.spec] || []).push(p);
});
// ASSIGN.rows = { loc: { code ('' = ช่องว่าง), by, at, note } } — ช่องที่ถูกเปลี่ยนทับข้อมูล Excel (จาก Google Sheet / เก็บในเครื่อง)
var ASSIGN = (function () {
  try { var d = JSON.parse(localStorage.getItem('h9pl_assign') || 'null'); if (d && d.rows) return d; } catch (e) { /* ไม่มี/เสีย */ }
  return { rows: {}, syncedAt: 0, local: false };
})();
// รวมข้อมูล: ช่องที่ถูกเปลี่ยน → ใช้ค่าที่เปลี่ยน · Part ที่ถูกย้ายไปช่องอื่น → ออกจากช่องเดิม · Part ที่ไม่มีช่อง → รอจัดพื้นที่
// ถ้า Part เดียวถูกกำหนดไว้หลายช่อง ใช้รายการล่าสุด
function applyAssignments() {
  var A = ASSIGN.rows || {}, moved = {};
  Object.keys(A).forEach(function (l) {
    var r = A[l]; if (!r.code || !byCode[r.code]) return;
    if (!moved[r.code] || String(moved[r.code].at) < String(r.at)) moved[r.code] = { loc: l, at: r.at };
  });
  byLoc = {};
  PARTS.forEach(function (p) {
    p.loc = moved[p.code] ? moved[p.code].loc : (p.loc0 && !A[p.loc0] ? p.loc0 : '');
  });
  PARTS.forEach(function (p) {
    var m = p.loc.match(/^H9-L(\d+)-(\d+)$/);
    p.lv = m ? +m[1] : 0; p.col = m ? +m[2] : 0;
    if (p.loc && byLoc[p.loc]) p.loc = ''; // กันซ้ำช่อง
    p.short = p.loc ? p.loc.replace(/^H9-/, '') : '';
    p.hay = [p.loc, p.short, p.code, p.spec, p.name.toUpperCase(), p.grp.toUpperCase(), p.loc ? '' : 'รอจัดพื้นที่'].join(' ');
    if (p.loc) byLoc[p.loc] = p;
  });
}
applyAssignments();
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
// ---------- ชื่อผู้ใช้ (บังคับกรอกก่อนใช้แอป · ใช้เป็นชื่อผู้เช็ค / ผู้เปลี่ยน Location) ----------
function userName() { return (store('h9pl_editor') || '').trim(); }
function validName(n) { n = String(n || '').trim(); return n.length >= 2 && !/^[0\-]+$/.test(n); }
function renderNameGate() {
  $('app').innerHTML =
    '<div class="top"><div class="brand">📍 H9 Part Location <span class="ver">v' + esc(APP_VERSION) + '</span></div></div>' +
    '<main><div class="set-card gate"><div class="gate-ic">👤</div><h3>กรอกชื่อผู้ใช้งานก่อนเริ่มใช้แอป</h3>' +
      '<p class="small muted">ชื่อนี้จะถูกบันทึกเป็นผู้ไล่เช็ค และผู้เปลี่ยน Part ↔ Location</p>' +
      '<input id="gateName" class="pin-in name-in" type="text" maxlength="40" autocomplete="name" placeholder="ชื่อ เช่น สมชาย">' +
      '<p class="small center" id="gateMsg" style="color:var(--red)"></p>' +
      '<button class="btn primary gate-btn" id="gateOk">เริ่มใช้งาน</button></div></main>';
  var inp = $('gateName'); setTimeout(function () { inp.focus(); }, 50);
  function ok() {
    var v = inp.value.trim();
    if (!validName(v)) { $('gateMsg').textContent = 'กรุณากรอกชื่อจริง อย่างน้อย 2 ตัวอักษร'; inp.focus(); return; }
    store('h9pl_editor', v); render();
  }
  $('gateOk').onclick = ok;
  inp.onkeydown = function (e) { if (e.key === 'Enter') { e.preventDefault(); ok(); } };
}
function openUserMenu() {
  openModal('<h3>👤 ผู้ใช้งาน</h3><div class="ed-cur"><div class="big" style="font-size:24px">' + esc(userName()) + '</div></div>' +
    '<label class="small muted">เปลี่ยนชื่อ</label><input id="unName" class="ed-note" type="text" maxlength="40" value="' + esc(userName()) + '">' +
    '<div class="actions"><button class="btn" id="unOut">🚪 ออก (เปลี่ยนคนใช้)</button><button class="btn primary" id="unSave">บันทึก</button></div>' +
    '<div class="actions"><button class="btn" id="unClose">ปิด</button></div>');
  $('unClose').onclick = closeModal;
  $('unSave').onclick = function () {
    var v = $('unName').value.trim();
    if (!validName(v)) { toast('กรุณากรอกชื่ออย่างน้อย 2 ตัวอักษร', 2500); return; }
    store('h9pl_editor', v); closeModal(); render(); toast('เปลี่ยนชื่อเป็น ' + v);
  };
  $('unOut').onclick = function () {
    if (!confirm('ออกจากชื่อ ' + userName() + '?\nคนถัดไปต้องกรอกชื่อของตัวเองก่อนใช้งาน')) return;
    store('h9pl_editor', ''); setAdminPin(''); closeModal(); S.tab = 'search'; render();
  };
}

function render() {
  if (!validName(userName())) { renderNameGate(); return; }
  $('app').innerHTML =
    '<div class="top"><div class="brand">📍 H9 Part Location <span class="ver">v' + esc(APP_VERSION) + '</span></div>' +
      '<span class="chip net hidden" id="netChip"></span>' +
      '<button class="chip user" id="appUser" aria-label="ผู้ใช้งาน">👤 ' + esc(userName()) + '</button>' +
      '<button class="chip" id="appQr" aria-label="QR ลิงก์เข้าหน้าเว็บ">🔗</button>' +
      '<button class="chip' + (S.tab === 'settings' ? ' on' : '') + '" id="appSet" aria-label="ตั้งค่า">⚙</button></div>' +
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
  $('appSet').onclick = function () { openSettings(); };
  $('appUser').onclick = openUserMenu;
  updateNet();
  if (S.tab === 'search') renderSearch(); else if (S.tab === 'check') renderCheck(); else if (S.tab === 'settings') renderSettings(); else renderMap();
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
    var emptyLoc = normLoc(S.q);
    if (!list.length) html += locExists(emptyLoc) && !byLoc[emptyLoc]
      ? '<div class="msg warn">⬜ ช่อง <b>' + esc(emptyLoc) + '</b> ว่าง (ไม่มี Part)</div>'
      : '<div class="msg warn">ไม่พบรายการที่ตรงกับ “' + esc(S.q) + '”</div>';
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
      (p.st && p.st !== 'Approved' ? '<span class="tag warn">' + esc(p.st) + '</span>' : '') +
      (p.loc !== p.loc0 ? '<span class="tag chg">✏️ เปลี่ยนจาก Excel' + (p.loc0 ? ' (เดิม ' + esc(p.loc0.replace(/^H9-/, '')) + ')' : '') + '</span>' : '') + '</div></div></button>';
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
      (p.loc !== p.loc0 ? '<div class="k">ตาม Excel</div><div class="v chgtxt">' + esc(p.loc0 || 'รอจัดพื้นที่') + ' (เปลี่ยนในเมนูตั้งค่า)</div>' : '') +
      '<div class="k">Line ผลิต</div><div class="v"><span class="tag ' + lnCls(p) + '">Line ' + esc(p.line) + '</span></div>' +
      '<div class="k">กลุ่ม Part</div><div class="v">' + esc(p.grp) + '</div>' +
      '<div class="k">Data Status</div><div class="v">' + esc(p.st) + '</div>' +
      (p.note ? '<div class="k">หมายเหตุ</div><div class="v" style="color:var(--amber)">⚠️ ' + esc(p.note) + '</div>' : '') +
    '</div>' +
    '<div class="actions">' + (p.loc ? '<button class="btn" id="dMap">🗂 ดูในผัง</button>' : '') +
      '<button class="btn primary" id="dScan">📷 สแกนต่อ</button></div>' +
    '<div class="actions"><button class="btn" id="dQr">🔗 QR ลิงก์</button><button class="btn" id="dEdit">✏️ เปลี่ยน Part</button></div>' +
    '<div class="actions"><button class="btn" id="dClose">ปิด</button></div>');
  Array.prototype.forEach.call(document.querySelectorAll('[data-copy]'), function (b) {
    b.onclick = function () { copy(b.getAttribute('data-copy')); };
  });
  $('dClose').onclick = closeModal;
  $('dEdit').onclick = function () { closeModal(); openSettings({ loc: p.loc || '', code: p.loc ? '' : p.code }); };
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
  $('main').innerHTML = summaryHtml() +
    '<div class="legend"><span><i style="background:var(--l4)"></i>Line 4</span><span><i style="background:var(--l5)"></i>Line 5</span><span><i style="background:var(--l1)"></i>Line 1</span>' +
      '<span>· แตะแถวในผังเพื่อเลือกชั้น</span></div>' +
    rackHtml(S.mark, true) +
    '<div class="lv-pick">' + LV_LIST.map(function (l) { return '<button data-pl="' + l + '" class="' + (l === lv ? 'active' : '') + '">L' + l + '</button>'; }).join('') + '</div>' +
    '<div class="count"><span>ชั้น L' + lv + ' · ' + n + ' ช่อง</span>' + (waiting ? '<a href="#" id="waitLink">รอจัดพื้นที่ ' + waiting + ' รายการ ›</a>' : '') + '</div>' +
    '<div class="slots">' + slots + '</div>';
  function showWaiting(e) { e.preventDefault(); S.q = 'รอจัดพื้นที่'; S.scan = null; S.line = 'all'; S.grp = 'all'; setTab('search'); }
  if ($('waitLink')) $('waitLink').onclick = showWaiting;
  if ($('waitTile')) $('waitTile').onclick = showWaiting;
  if ($('sumBox')) $('sumBox').ontoggle = function () { store('h9pl_sum', this.open ? '1' : '0'); };
  function pick(l) { S.mapLv = l; renderMap(); }
  Array.prototype.forEach.call(document.querySelectorAll('[data-pl]'), function (b) { b.onclick = function () { pick(+b.getAttribute('data-pl')); }; });
  Array.prototype.forEach.call(document.querySelectorAll('[data-lv]'), function (b) { b.onclick = function () { pick(+b.getAttribute('data-lv')); }; });
  Array.prototype.forEach.call(document.querySelectorAll('.slot[data-i]'), function (b) { b.onclick = function () { showDetail(PARTS[+b.getAttribute('data-i')]); }; });
}
// สรุป: จำนวนช่อง / Part ทั้งหมด และแยกตาม Line · ชั้น · กลุ่ม (คำนวณจาก js/data.js ทุกครั้ง)
function summaryHtml() {
  var slots = 0, used = 0, withLoc = 0, waiting = 0, byLine = {}, byGrp = {};
  LV_LIST.forEach(function (l) { slots += LEVELS[l]; });
  PARTS.forEach(function (p) {
    var k = p.line, x = byLine[k] = byLine[k] || [0, 0], g = byGrp[p.grp] = byGrp[p.grp] || [0, 0];
    x[0]++; g[0]++;
    if (p.loc) { withLoc++; x[1]++; g[1]++; } else waiting++;
  });
  used = Object.keys(byLoc).length;
  function tile(n, label, cls, id) {
    return '<div class="tile' + (cls ? ' ' + cls : '') + '"' + (id ? ' id="' + id + '" role="button"' : '') + '><b>' + n + '</b><span>' + label + '</span></div>';
  }
  function row(cells, head) { var t = head ? 'th' : 'td'; return '<tr>' + cells.map(function (c) { return '<' + t + '>' + c + '</' + t + '>'; }).join('') + '</tr>'; }
  var lines = Object.keys(byLine).sort(function (a, b) { return LINES.indexOf(+a) - LINES.indexOf(+b); });
  var lvRows = LV_LIST.map(function (l) {
    var u = 0; for (var c = 1; c <= LEVELS[l]; c++) if (byLoc['H9-L' + l + '-' + ('0' + c).slice(-2)]) u++;
    return row(['L' + l, LEVELS[l], u, LEVELS[l] - u]);
  });
  var open = store('h9pl_sum') === '1';
  return '<details class="sum" id="sumBox"' + (open ? ' open' : '') + '><summary>' +
      '<div class="tiles">' + tile(slots, 'ช่องทั้งหมด') + tile(used, 'ช่องที่มี Part') + tile(slots - used, 'ช่องว่าง', slots - used ? 'warn' : '') +
        tile(PARTS.length, 'Part ทั้งหมด') + tile(withLoc, 'มี Location') + tile(waiting, 'รอจัดพื้นที่ ›', waiting ? 'warn' : '', 'waitTile') + '</div>' +
      '<div class="sum-more">📊 ดูสรุปแยก Line / ชั้น / กลุ่ม</div></summary>' +
    '<h4>แยกตาม Line ผลิต</h4><table>' + row(['Line', 'Part', 'มี Location', 'รอจัดพื้นที่'], true) +
      lines.map(function (k) { return row(['<span class="tag ln' + esc(k) + '">Line ' + esc(k) + '</span>', byLine[k][0], byLine[k][1], byLine[k][0] - byLine[k][1]]); }).join('') +
      row(['<b>รวม</b>', '<b>' + PARTS.length + '</b>', '<b>' + withLoc + '</b>', '<b>' + waiting + '</b>']) + '</table>' +
    '<h4>แยกตามชั้น</h4><table>' + row(['ชั้น', 'ช่อง', 'มี Part', 'ว่าง'], true) + lvRows.join('') +
      row(['<b>รวม</b>', '<b>' + slots + '</b>', '<b>' + used + '</b>', '<b>' + (slots - used) + '</b>']) + '</table>' +
    '<h4>แยกตามกลุ่ม Part</h4><table>' + row(['กลุ่ม', 'Part', 'มี Location', 'รอจัดพื้นที่'], true) +
      Object.keys(byGrp).sort().map(function (g) { return row([esc(g), byGrp[g][0], byGrp[g][1], byGrp[g][0] - byGrp[g][1]]); }).join('') + '</table>' +
    '</details>';
}
function scrollToSlot() {
  setTimeout(function () { var el = S.mark && $('s-' + S.mark); if (el) el.scrollIntoView({ block: 'center', behavior: 'smooth' }); }, 60);
}

// ======================= ไล่เช็ค Part ↔ Location =======================
// ป้าย Location ไม่มี QR → สแกนป้าย Part → แอปบอกว่า "ควรอยู่ช่องไหน" → ดูช่องจริงแล้วยืนยัน
// แต่ละช่องมีป้าย 2 ด้าน (หน้า / หลัง) → ต้องยืนยันครบทั้ง 2 ด้าน ช่องนั้นถึงนับว่าเช็คเสร็จ
// เก็บในเครื่อง (localStorage): rec[loc] = { f: {st, at, found, text}, b: {...} }  (f = ด้านหน้า, b = ด้านหลัง)
var SIDES = [['f', 'ด้านหน้า'], ['b', 'ด้านหลัง']];
var CK = (function () {
  try { var d = JSON.parse(store('h9pl_check3') || 'null'); if (d && d.rec) return d; } catch (e) { /* ข้อมูลเสีย → เริ่มใหม่ */ }
  return { rec: {}, pend: null, last: null };
})();
var ALL_LOCS = [];
Object.keys(LEVELS).map(Number).sort(function (a, b) { return a - b; }).forEach(function (l) {
  for (var c = 1; c <= LEVELS[l]; c++) ALL_LOCS.push('H9-L' + l + '-' + ('0' + c).slice(-2));
});
var CK_LBL = { ok: '✅ ตรง', ng: '❌ ไม่ตรง', empty: '⬜ ช่องว่าง', half: '½ เช็คแล้ว 1 ด้าน' };
var SIDE_LBL = { ok: '✅ ตรง', ng: '❌ ไม่ตรง', empty: '⬜ ว่าง' };
function saveCheck() { store('h9pl_check3', JSON.stringify(CK)); }
function locExists(l) { var m = l && l.match(/^H9-L(\d+)-(\d+)$/); return !!(m && LEVELS[+m[1]] && +m[2] <= LEVELS[+m[1]]); }
function shortLoc(l) { return String(l || '').replace(/^H9-/, ''); }
// ผลรวมของช่อง: ไม่ตรง (ด้านใดด้านหนึ่ง) > ว่าง > ครบ 2 ด้านตรง = ตรง · มีด้านเดียว = half
function locStatus(loc) {
  var r = CK.rec[loc]; if (!r || (!r.f && !r.b)) return '';
  var s = [r.f && r.f.st, r.b && r.b.st];
  if (s.indexOf('ng') >= 0) return 'ng';
  if (!r.f || !r.b) return 'half';
  if (s.indexOf('empty') >= 0) return 'empty';
  return 'ok';
}
function lastAt(loc) { var r = CK.rec[loc] || {}; return Math.max(r.f ? r.f.at : 0, r.b ? r.b.at : 0); }
function ckStats() {
  var st = { ok: 0, ng: 0, empty: 0, half: 0, done: 0 };
  ALL_LOCS.forEach(function (l) {
    var s = locStatus(l); if (!s) return;
    st[s]++;
    var r = CK.rec[l]; if (r.f && r.b) st.done++;
  });
  return st;
}

// สแกน/พิมพ์ 1 ครั้ง → คืน class ไฟกระพริบในกล้อง (ok / err / loc)
function checkScan(text) {
  text = String(text || '').trim(); if (!text) return '';
  // พิมพ์ Location เอง (เช่น L3-05) → เปิดช่องนั้นให้ยืนยันด้วยตา
  var lq = text.match(/#q=([^&\s]+)/), raw = text;
  if (lq) { try { raw = decodeURIComponent(lq[1]); } catch (e) { raw = lq[1]; } }
  var l = normLoc(raw);
  if (locExists(l)) { openLoc(l); beep(true); return 'loc'; }
  var r = resolveScan(text);
  if (!r) { CK.pend = null; CK.last = { type: 'unk', text: text }; saveCheck(); beep(false); return 'err'; }
  var cand = r.parts.filter(function (p) { return p.loc; });
  if (!cand.length) {
    CK.pend = null; CK.last = { type: 'wait', text: text, parts: r.parts.map(function (p) { return p.i; }) }; saveCheck(); beep(false); return 'err';
  }
  CK.pend = { text: text.slice(0, 120), cand: cand.map(function (p) { return p.i; }), ng: '' };
  CK.last = null; saveCheck(); beep(true);
  return 'loc';
}
// เลือกช่องเอง (แตะผัง / รายการ / พิมพ์ Location) → ยืนยันด้วยตาโดยไม่ต้องสแกน
function openLoc(loc) {
  var p = byLoc[loc];
  CK.pend = p ? { text: '', cand: [p.i], ng: '', manual: true } : null;
  CK.last = p ? null : { type: 'noloc', loc: loc };
  saveCheck();
}
// ยืนยัน 1 ด้าน · ครบ 2 ด้านแล้วปิดการ์ด แสดงผลรวม
function confirmSide(loc, side, st, found) {
  var pend = CK.pend || {};
  var r = CK.rec[loc] = CK.rec[loc] || {};
  r[side] = { st: st, at: Date.now(), text: pend.text || '', found: found || '', by: userName() };
  if (CK.pend) CK.pend.ng = '';
  if (r.f && r.b) { CK.pend = null; CK.last = { type: 'res', loc: loc }; }
  saveCheck();
  beep(st === 'ok');
}

function sideRow(p, side, label, manual) {
  var r = (CK.rec[p.loc] || {})[side], P = CK.pend || {}, key = p.loc + '|' + side, L = esc(p.loc);
  if (P.ng === key) {
    return '<div class="ck-side ng-open"><div class="sn">' + label + '</div>' +
      '<div class="ck-ng"><div class="small">' + label + ': พบ Part นี้อยู่ที่ช่องไหน? (ไม่ทราบ กด "ไม่ระบุ")</div>' +
        '<div class="ck-ngrow"><input id="ckFound" type="text" autocapitalize="characters" autocomplete="off" placeholder="เช่น L3-07">' +
        '<button class="btn primary" data-ckact="saveng" data-loc="' + L + '" data-side="' + side + '">บันทึก</button></div>' +
        '<div class="actions"><button class="btn" data-ckact="saveng0" data-loc="' + L + '" data-side="' + side + '">ไม่ระบุ</button>' +
        '<button class="btn" data-ckact="cancelng">ยกเลิก</button></div></div></div>';
  }
  return '<div class="ck-side' + (r ? ' done ' + r.st : '') + '"><div class="sn">' + label +
      (r ? ' <span class="ck-chip ' + r.st + '">' + SIDE_LBL[r.st] + (r.st === 'ng' && r.found ? ' → ' + esc(shortLoc(r.found)) : '') + '</span>' : ' <span class="todo">ยังไม่ยืนยัน</span>') + '</div>' +
    '<div class="ck-btns2"><button class="btn ck-yes" data-ckact="ok" data-loc="' + L + '" data-side="' + side + '">✅ ตรง</button>' +
      '<button class="btn ck-no" data-ckact="ng" data-loc="' + L + '" data-side="' + side + '">❌ ไม่ตรง</button>' +
      (manual ? '<button class="btn ck-empty" data-ckact="empty" data-loc="' + L + '" data-side="' + side + '">⬜</button>' : '') + '</div></div>';
}
function ckPendHtml() {
  var P = CK.pend;
  if (!P) return ckLastHtml();
  return P.cand.map(function (i) {
    var p = PARTS[i];
    return '<div class="ck-pend ' + lnCls(p) + '">' +
      '<div class="pp">📦 <b>' + esc(p.code) + '</b> · ' + esc(p.name) + (p.spec ? ' <span class="muted">(' + esc(p.spec) + ')</span>' : '') + '</div>' +
      '<div class="lbl">' + (P.manual ? 'ช่องนี้ควรมี Part ข้างบน' : 'ควรอยู่ที่') + '</div>' +
      '<div class="big">' + esc(p.loc) + '</div>' +
      '<div class="sub">ชั้น L' + p.lv + ' · ช่องที่ ' + p.col + ' · Line ' + esc(p.line) + '</div>' +
      SIDES.map(function (s) { return sideRow(p, s[0], s[1], P.manual); }).join('') +
      '</div>';
  }).join('') + (P.cand.length > 1 ? '<div class="small muted center">Spec นี้มีหลายรหัส — ยืนยันเฉพาะช่องที่ตรงกับของจริง</div>' : '');
}
function ckLastHtml() {
  var L = CK.last;
  if (!L) return '<div class="ck-res idle"><div class="h">📦 สแกนป้าย Part</div>แอปจะบอกว่า Part ควรอยู่ช่องไหน → ดูช่องจริง แล้วยืนยัน<b> ด้านหน้า</b> และ<b> ด้านหลัง</b></div>';
  if (L.type === 'unk') return '<div class="ck-res err"><div class="h">❓ ไม่พบ Part นี้ในรายการ</div>สแกนได้: <b>' + esc(L.text) + '</b></div>';
  if (L.type === 'noloc') return '<div class="ck-res warn"><div class="h">' + esc(L.loc) + '</div>ช่องนี้ไม่มี Part ในรายการ</div>';
  if (L.type === 'wait') return '<div class="ck-res warn"><div class="h">⏳ Part นี้ยังรอจัดพื้นที่</div>' +
    L.parts.map(function (i) { return '<b>' + esc(PARTS[i].code) + '</b> · ' + esc(PARTS[i].name); }).join('<br>') + '<br>ยังไม่มี Location ในรายการ</div>';
  var r = CK.rec[L.loc], p = byLoc[L.loc], st = locStatus(L.loc); if (!r) return '';
  var who = p ? esc(p.code) + ' · ' + esc(p.name) : '';
  var sides = SIDES.map(function (s) {
    var x = r[s[0]]; return s[1] + ' ' + (x ? SIDE_LBL[x.st] + (x.st === 'ng' ? ' (พบที่ ' + esc(x.found || 'ไม่ระบุ') + ')' : '') : '-');
  }).join(' · ');
  var cls = st === 'ok' ? 'ok' : st === 'ng' ? 'err' : 'warn';
  return '<div class="ck-res ' + cls + '"><div class="h">' + CK_LBL[st] + ' ครบ 2 ด้าน</div>' + esc(L.loc) + ' · ' + who +
    '<div class="small">' + sides + '</div><div class="small">สแกน Part ถัดไปได้เลย</div></div>';
}
function ckSummaryHtml() {
  var st = ckStats();
  return '<div class="center small muted">ครบ 2 ด้าน ' + st.done + '/' + ALL_LOCS.length + ' · ✅ ' + st.ok + ' · ❌ ' + st.ng + ' · ⬜ ' + st.empty + ' · ½ ' + st.half + '</div>';
}
function ckCamStatus() { return ckPendHtml() + ckSummaryHtml(); }
function refreshCheck() {
  if ($('camStatus')) $('camStatus').innerHTML = ckCamStatus();
  else if (S.tab === 'check') renderCheck();
}
// ปุ่มในการ์ด (ทั้งในหน้าแท็บ และใต้กล้อง) — ใช้ตัวจับคลิกตัวเดียว
document.addEventListener('click', function (e) {
  var b = e.target.closest && e.target.closest('[data-ckact]'); if (!b) return;
  var act = b.getAttribute('data-ckact'), loc = b.getAttribute('data-loc'), side = b.getAttribute('data-side');
  if (act === 'ok') confirmSide(loc, side, 'ok');
  else if (act === 'empty') confirmSide(loc, side, 'empty');
  else if (act === 'ng') { CK.pend.ng = loc + '|' + side; saveCheck(); refreshCheck(); setTimeout(function () { if ($('ckFound')) $('ckFound').focus(); }, 30); return; }
  else if (act === 'cancelng') { CK.pend.ng = ''; saveCheck(); }
  else if (act === 'saveng' || act === 'saveng0') {
    var v = act === 'saveng' && $('ckFound') ? $('ckFound').value.trim() : '';
    var nl = normLoc(v);
    if (v && !locExists(nl)) { toast('รูปแบบช่องไม่ถูกต้อง เช่น L3-07', 2500); return; }
    confirmSide(loc, side, 'ng', nl);
  }
  refreshCheck();
});
document.addEventListener('keydown', function (e) {
  if (e.key === 'Enter' && e.target && e.target.id === 'ckFound') {
    e.preventDefault(); var b = document.querySelector('[data-ckact="saveng"]'); if (b) b.click();
  }
});

function ckRackHtml() {
  var cur = CK.pend && CK.pend.cand.length === 1 ? PARTS[CK.pend.cand[0]].loc : '';
  var rows = LV_LIST.map(function (lv) {
    var cells = '';
    for (var c = 1; c <= MAX_COL; c++) {
      var loc = 'H9-L' + lv + '-' + ('0' + c).slice(-2), s = locStatus(loc);
      var cls = c > LEVELS[lv] || !byLoc[loc] ? 'void' : s ? 'ck-' + s : '';
      cells += '<div class="rc ' + cls + (loc === cur ? ' on' : '') + '"' + (cls === 'void' ? '' : ' data-ck="' + loc + '"') + '></div>';
    }
    return '<div class="rack-row"><div class="rack-lv">L' + lv + '</div><div class="rack-cells">' + cells + '</div></div>';
  }).join('');
  return '<div class="rack ck">' + rows + '<div class="rack-axis"><div></div><div><span>01</span><span>10</span><span>20</span><span>30</span><span>' + MAX_COL + '</span></div></div></div>';
}

function renderCheck() {
  var st = ckStats(), N = ALL_LOCS.length || 1, pct = Math.round(st.done * 100 / N);
  var f = S.ckf || 'ng';
  var list = f === 'todo'
    ? ALL_LOCS.filter(function (l) { var r = CK.rec[l]; return !r || !r.f || !r.b; })
        .sort(function (a, b) { return (lastAt(b) ? 1 : 0) - (lastAt(a) ? 1 : 0); })
    : ALL_LOCS.filter(function (l) { var s = locStatus(l); return s && (f === 'all' || s === 'ng' || s === 'empty'); })
        .sort(function (a, b) { return lastAt(b) - lastAt(a); });
  $('main').innerHTML =
    '<div class="ck-prog"><div class="ck-bar"><i class="ok" style="width:' + (st.ok * 100 / N) + '%"></i>' +
      '<i class="ng" style="width:' + (st.ng * 100 / N) + '%"></i><i class="empty" style="width:' + (st.empty * 100 / N) + '%"></i>' +
      '<i class="half" style="width:' + (st.half * 100 / N) + '%"></i></div>' +
      '<div class="ck-nums"><span>ครบ 2 ด้าน <b>' + st.done + '</b>/' + ALL_LOCS.length + ' (' + pct + '%)</span>' +
      '<span>✅ ' + st.ok + '</span><span>❌ ' + st.ng + '</span><span>⬜ ' + st.empty + '</span><span>½ ' + st.half + '</span></div></div>' +
    '<div class="search ck-search"><div class="search-wrap"><input id="ckq" type="text" enterkeyhint="go" autocomplete="off" autocapitalize="characters" spellcheck="false" placeholder="สแกน Part / พิมพ์รหัส"></div>' +
      '<button class="scan-btn" id="ckCam">📷 สแกน</button></div>' +
    ckPendHtml() +
    '<div class="legend"><span><i style="background:var(--green)"></i>ตรง 2 ด้าน</span><span><i style="background:#9fd9b4"></i>เช็ค 1 ด้าน</span>' +
      '<span><i style="background:var(--red)"></i>ไม่ตรง</span><span><i style="background:#f2b544"></i>ว่าง</span><span><i style="background:var(--line)"></i>ยังไม่เช็ค</span></div>' +
    ckRackHtml() +
    '<div class="filters">' +
      '<button class="fchip' + (f === 'ng' ? ' active' : '') + '" data-ckf="ng">❌ ไม่ตรง / ว่าง ' + (st.ng + st.empty) + '</button>' +
      '<button class="fchip' + (f === 'todo' ? ' active' : '') + '" data-ckf="todo">ยังไม่ครบ ' + (ALL_LOCS.length - st.done) + '</button>' +
      '<button class="fchip' + (f === 'all' ? ' active' : '') + '" data-ckf="all">ผลทั้งหมด</button></div>' +
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
    openCameraScanner(null, { title: 'ไล่เช็ค — สแกนป้าย Part', status: ckCamStatus, onEach: checkScan });
    onModalClose = (function (stopCam) { return function () { if (stopCam) stopCam(); renderCheck(); }; })(onModalClose);
  };
  function pick(loc) { openLoc(loc); renderCheck(); window.scrollTo({ top: 0, behavior: 'smooth' }); }
  Array.prototype.forEach.call(document.querySelectorAll('[data-ck]'), function (b) { b.onclick = function () { pick(b.getAttribute('data-ck')); }; });
  Array.prototype.forEach.call(document.querySelectorAll('[data-ckrow]'), function (b) { b.onclick = function () { pick(b.getAttribute('data-ckrow')); }; });
  Array.prototype.forEach.call(document.querySelectorAll('[data-ckf]'), function (b) {
    b.onclick = function () { S.ckf = b.getAttribute('data-ckf'); renderCheck(); };
  });
  $('ckCsv').onclick = exportCheck;
  $('ckReset').onclick = function () {
    if (!confirm('ล้างผลการเช็คทั้งหมดในเครื่องนี้ แล้วเริ่มใหม่?\n(แนะนำให้ Export เก็บไว้ก่อน)')) return;
    CK = { rec: {}, pend: null, last: null }; saveCheck(); renderCheck();
  };
}
function fmtTime(t) {
  var d = new Date(t), z = function (n) { return ('0' + n).slice(-2); };
  return z(d.getDate()) + '/' + z(d.getMonth() + 1) + ' ' + z(d.getHours()) + ':' + z(d.getMinutes());
}
function ckRow(loc) {
  var r = CK.rec[loc] || {}, p = byLoc[loc], s = locStatus(loc), t = lastAt(loc);
  var sides = SIDES.map(function (sd) {
    var x = r[sd[0]];
    return '<span class="ck-chip ' + (x ? x.st : 'todo') + '">' + sd[1].replace('ด้าน', '') + ' ' + (x ? SIDE_LBL[x.st] : '—') +
      (x && x.st === 'ng' ? ' → ' + esc(shortLoc(x.found) || 'ไม่ระบุ') : '') + '</span>';
  }).join(' ');
  return '<button class="ck-row ' + (s || 'todo') + '" data-ckrow="' + esc(loc) + '">' +
    '<div class="l">' + esc(shortLoc(loc)) + '</div><div class="m">' +
    '<div>' + sides + (t ? ' <span class="muted small">' + fmtTime(t) + '</span>' : '') + '</div>' +
    '<div class="small"><b>' + esc(p ? p.code : '-') + '</b> ' + esc(p ? p.name : '') + '</div>' +
    '</div></button>';
}
function exportCheck() {
  var head = ['Location', 'ผลรวม', 'ด้านหน้า', 'ด้านหน้า: พบอยู่ที่', 'ด้านหลัง', 'ด้านหลัง: พบอยู่ที่',
    'Child Item Code', 'Specification', 'ชื่อ Part', 'Line', 'กลุ่ม Part', 'เวลาล่าสุด', 'ผู้เช็ค'];
  var rows = ALL_LOCS.map(function (loc) {
    var r = CK.rec[loc] || {}, p = byLoc[loc], s = locStatus(loc), t = lastAt(loc);
    function sd(x) { return x ? SIDE_LBL[x.st].replace(/^\S+\s/, '') : 'ยังไม่เช็ค'; }
    function fd(x) { return x && x.st === 'ng' ? (x.found || 'ไม่ระบุ') : ''; }
    return [loc, s ? CK_LBL[s].replace(/^\S+\s/, '') : 'ยังไม่เช็ค', sd(r.f), fd(r.f), sd(r.b), fd(r.b),
      p ? p.code : '', p ? p.spec : '', p ? p.name : '(ไม่มี Part)', p ? p.line : '', p ? p.grp : '', t ? new Date(t).toLocaleString('th-TH') : '',
      [r.f && r.f.by, r.b && r.b.by].filter(function (x, i, a) { return x && a.indexOf(x) === i; }).join(', ')];
  });
  var csv = '﻿' + [head].concat(rows).map(function (row) {
    return row.map(function (v) { v = String(v == null ? '' : v); return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }).join(',');
  }).join('\r\n');
  var d = new Date(), z = function (n) { return ('0' + n).slice(-2); };
  var a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  a.download = 'H9-check-' + d.getFullYear() + z(d.getMonth() + 1) + z(d.getDate()) + '-' + z(d.getHours()) + z(d.getMinutes()) + '.csv';
  document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); document.body.removeChild(a); }, 1000);
  toast('Export ' + rows.length + ' แถว');
}

// ======================= เชื่อมต่อ Apps Script (Google Sheet) =======================
// ส่ง POST แบบ text/plain (ไม่เกิด CORS preflight) → Apps Script ตอบ JSON { ok, data | error } (แบบเดียวกับ pda-location-scan)
function apiUrl() { return store('h9pl_api_url') || (window.H9_CONFIG && window.H9_CONFIG.API_URL) || ''; }
function api(action) {
  var args = Array.prototype.slice.call(arguments, 1), url = apiUrl();
  if (!url) return Promise.reject(new Error('ยังไม่ได้ตั้งค่า Google Sheet'));
  var ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
  var timer = ctrl && setTimeout(function () { ctrl.abort(); }, 30000);
  return fetch(url + (url.indexOf('?') >= 0 ? '&' : '?') + '_r=' + Date.now(), {
    method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action: action, args: args }), redirect: 'follow', cache: 'no-store',
    signal: ctrl ? ctrl.signal : undefined
  }).then(function (res) {
    return res.text().then(function (t) {
      try { return JSON.parse(t); } catch (e) { throw new Error('เซิร์ฟเวอร์ตอบกลับผิดรูปแบบ (HTTP ' + res.status + ')'); }
    });
  }, function () { var e = new Error('เชื่อมต่อเซิร์ฟเวอร์ไม่ได้ ตรวจสอบอินเทอร์เน็ต'); e.network = true; throw e; }).then(function (r) {
    if (timer) clearTimeout(timer);
    if (!r.ok) throw new Error(/auth|permission|权限/i.test(r.error || '') ? 'ระบบยังไม่ได้รับอนุญาต — ผู้ดูแลต้องเปิด Google Sheet → เมนู H9 Location → ตั้งค่าเริ่มต้น' : (r.error || 'เกิดข้อผิดพลาด'));
    return r.data;
  }, function (e) { if (timer) clearTimeout(timer); throw e; });
}
function rowsToMap(rows) {
  var m = {}; (rows || []).forEach(function (r) { if (r.loc) m[r.loc] = { code: r.code || '', by: r.by, at: r.at, note: r.note }; }); return m;
}
function saveAssign() { store('h9pl_assign', JSON.stringify(ASSIGN)); }
function setAssignRows(rows) {
  var pending = ASSIGN.pending || [];
  ASSIGN = { rows: rowsToMap(rows), syncedAt: Date.now(), local: false, pending: pending };
  pending.forEach(applyOpLocal);
  saveAssign(); applyAssignments();
}
// ---------- ออฟไลน์: การเปลี่ยนที่ยังส่งไม่ได้ เก็บเป็นคิวในเครื่อง แล้วส่งเมื่อกลับมามีเน็ต ----------
function nowLocal() { var d = new Date(), z = function (n) { return ('0' + n).slice(-2); }; return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate()) + 'T' + z(d.getHours()) + ':' + z(d.getMinutes()) + ':' + z(d.getSeconds()); }
function applyOpLocal(op) {
  if (op.type === 'revert') { delete ASSIGN.rows[op.loc]; return; }
  op.changes.forEach(function (c) { ASSIGN.rows[c.loc] = { code: c.code, by: op.by, at: op.at, note: op.note || '', pending: true }; });
}
function queueOp(op) {
  ASSIGN.pending = ASSIGN.pending || [];
  ASSIGN.pending.push(op); applyOpLocal(op);
  saveAssign(); applyAssignments(); updateNet();
}
function pendingCount() { return (ASSIGN.pending || []).length; }
var flushing = null;
function flushPending() {
  if (flushing) return flushing;
  if (!pendingCount() || !apiUrl()) return Promise.resolve(0);
  var sent = 0;
  function next() {
    var op = (ASSIGN.pending || [])[0];
    if (!op) return Promise.resolve(sent);
    var call = op.type === 'revert'
      ? api('revertAssignment', op.pin, op.by, op.loc, op.oldCode, op.newCode)
      : api('saveAssignments', op.pin, op.by, op.changes, op.note || '');
    return call.then(function (rows) {
      ASSIGN.pending.shift(); sent++; setAssignRows(rows); return next();
    }, function (e) {
      if (e.network) throw e;                       // ยังไม่มีเน็ต → เก็บไว้ส่งรอบหน้า
      ASSIGN.pending.shift(); saveAssign();         // เซิร์ฟเวอร์ปฏิเสธ (เช่น รหัสผิด) → ทิ้งรายการนั้น
      toast('ส่งรายการที่รอไม่สำเร็จ: ' + e.message, 4000);
      return next();
    });
  }
  flushing = next().then(function (n) {
    flushing = null; updateNet();
    if (n) { toast('ส่งการเปลี่ยนที่รออยู่ขึ้น Google Sheet แล้ว ' + n + ' รายการ', 3000); if ($('modal').classList.contains('hidden')) render(); }
    return n;
  }, function (e) { flushing = null; updateNet(); throw e; });
  return flushing;
}
// แถบสถานะเน็ตบนหัวแอป: 📴 ออฟไลน์ · ⏳ รอส่ง n
function updateNet() {
  var el = $('netChip'); if (!el) return;
  var off = navigator.onLine === false, n = pendingCount();
  el.textContent = off ? '📴 ออฟไลน์' + (n ? ' · รอส่ง ' + n : '') : n ? '⏳ รอส่ง ' + n : '';
  el.classList.toggle('hidden', !off && !n);
}
// ดึงการเปลี่ยนล่าสุดจาก Google Sheet (ตอนเปิดแอป / กลับมาที่หน้าเว็บ) → รวมข้อมูลใหม่แล้ววาดหน้าจอใหม่
var syncing = null, syncErr = '';
function syncAssignments(force) {
  if (!apiUrl()) return Promise.resolve(false);
  if (syncing) return syncing;
  if (!force && Date.now() - (ASSIGN.syncedAt || 0) < 60000) return Promise.resolve(false);
  syncing = flushPending().catch(function () { /* ยังส่งไม่ได้ → ดึงข้อมูลต่อ */ }).then(function () { return api('getAssignments'); }).then(function (rows) {
    var before = JSON.stringify(ASSIGN.rows);
    setAssignRows(rows);
    var changed = before !== JSON.stringify(ASSIGN.rows);
    if ((changed || force) && $('modal').classList.contains('hidden')) render();
    return changed;
  }).then(function (c) { syncing = null; syncErr = ''; return c; }, function (e) {
    syncing = null; syncErr = e.message || String(e);
    if (S.tab === 'settings' && $('modal').classList.contains('hidden')) renderSettings();
    throw e;
  });
  return syncing;
}

// ======================= ⚙ ตั้งค่า (ต้องใส่รหัส) =======================
var LOCAL_PIN = '1234'; // ใช้เฉพาะตอนยังไม่ได้ต่อ Google Sheet — ถ้าต่อแล้ว เซิร์ฟเวอร์เป็นคนตรวจ (Script Properties → ADMIN_PIN)
function pinHash(v) { var h = 5381; v = 'h9pl:' + v; for (var i = 0; i < v.length; i++) h = ((h << 5) + h + v.charCodeAt(i)) | 0; return String(h >>> 0); }
function adminPin() { try { return sessionStorage.getItem('h9pl_pin') || ''; } catch (e) { return ''; } }
function setAdminPin(v) { try { if (v) sessionStorage.setItem('h9pl_pin', v); else sessionStorage.removeItem('h9pl_pin'); } catch (e) { /* โหมดส่วนตัว */ } }
function openSettings(prefill) {
  function go() {
    S.ed = { loc: '', code: '', mode: 'swap', q: '', note: '' };
    if (prefill) { S.ed.loc = prefill.loc || ''; S.ed.code = prefill.code || ''; }
    setTab('settings');
  }
  if (adminPin()) { go(); return; }
  openModal('<h3>⚙ ตั้งค่า</h3><p class="small muted">ใส่รหัสเพื่อเข้าเมนูตั้งค่า</p>' +
    '<input id="pinIn" class="pin-in" type="password" inputmode="numeric" autocomplete="off" maxlength="12" placeholder="รหัส">' +
    '<p class="small center" id="pinMsg" style="color:var(--red)"></p>' +
    '<div class="actions"><button class="btn" id="pinCancel">ยกเลิก</button><button class="btn primary" id="pinOk">เข้า</button></div>');
  var inp = $('pinIn'); setTimeout(function () { inp.focus(); }, 50);
  $('pinCancel').onclick = closeModal;
  function check() {
    var v = inp.value.trim(); if (!v) return;
    $('pinOk').disabled = true; $('pinMsg').textContent = '';
    var p = apiUrl() ? api('checkPin', v).then(function (ok) {
      if (ok) store('h9pl_pinok', pinHash(v));
      return ok;
    }, function (e) {
      if (e.network && store('h9pl_pinok')) return pinHash(v) === store('h9pl_pinok'); // ออฟไลน์ → เทียบกับรหัสที่เคยผ่าน
      if (e.network) throw new Error('ไม่มีเน็ต — ต้องเข้าเมนูตั้งค่าตอนมีเน็ตอย่างน้อย 1 ครั้งก่อน');
      throw e;
    }) : Promise.resolve(v === LOCAL_PIN);
    function msg(t) { if ($('pinMsg')) $('pinMsg').textContent = t; if ($('pinOk')) $('pinOk').disabled = false; }
    p.then(function (ok) {
      if (!$('pinIn')) return; // ปิดหน้าต่างไปแล้ว
      if (!ok) { msg('รหัสไม่ถูกต้อง'); inp.select(); beep(false); return; }
      setAdminPin(v); closeModal(); go();
    }, function (e) { msg(e.message); });
  }
  $('pinOk').onclick = check;
  inp.onkeydown = function (e) { if (e.key === 'Enter') { e.preventDefault(); check(); } };
}

function editorBy() { return userName(); }
function partFind(q) {
  var r = resolveScan(q); if (r && r.parts.length) return r.parts;
  var toks = tokens(q); if (!toks.length) return [];
  return PARTS.filter(function (p) {
    for (var i = 0; i < toks.length; i++) if (p.hay.indexOf(toks[i]) < 0) return false;
    return true;
  }).sort(locOrder).slice(0, 8);
}
function partLine(p) {
  return '<b>' + esc(p.code) + '</b> · ' + esc(p.name) + ' <span class="muted small">(' + esc(p.spec) + ')</span>';
}
// รายการเปลี่ยนจากสถานะในตัวแก้ไข: ช่อง T ← Part N · Part เดิมของ T (O) → สลับไปช่องเดิมของ N หรือ รอจัดพื้นที่
function plannedChanges() {
  var E = S.ed, T = E.loc, O = byLoc[T] || null, N = E.code === '-' ? null : (E.code && byCode[E.code] ? byCode[E.code][0] : undefined);
  if (!T || N === undefined) return null;
  if (N && O && N === O) return [];
  var ch = [{ loc: T, code: N ? N.code : '', old: O ? O.code : '' }];
  if (N && N.loc && N.loc !== T) ch.push({ loc: N.loc, code: E.mode === 'swap' && O ? O.code : '', old: N.code });
  return ch;
}

function renderSettings() {
  if (!adminPin()) { setTab('search'); return; }
  var E = S.ed = S.ed || { loc: '', code: '', mode: 'swap', q: '', note: '' };
  var changed = Object.keys(ASSIGN.rows || {}).sort(), api_ = !!apiUrl();
  var T = E.loc, O = byLoc[T], N = E.code && E.code !== '-' && byCode[E.code] ? byCode[E.code][0] : null;
  var ch = plannedChanges();

  var edit = '<div class="set-card"><h4>✏️ เปลี่ยน Part ของ Location</h4>' +
    '<label class="small muted">① Location</label>' +
    '<div class="ck-ngrow"><input id="edLoc" type="text" autocapitalize="characters" autocomplete="off" placeholder="เช่น L3-05" value="' + esc(T ? T.replace(/^H9-/, '') : '') + '">' +
      '<button class="btn" id="edLocOk">เลือก</button></div>';
  if (!T) edit += '<div class="small muted ed-hint">พิมพ์ Location แล้ว ② ช่องค้นหา / 📷 สแกน Part จะขึ้นทันที</div>';
  if (T) {
    edit += '<div class="ed-cur"><div class="big">' + esc(T) + '</div>' +
      '<div class="small">ตอนนี้: ' + (O ? partLine(O) : '<b>ช่องว่าง</b>') + '</div>' +
      (byLoc[T] && O.loc0 !== T ? '<div class="small chgtxt">✏️ เปลี่ยนจาก Excel แล้ว</div>' : '') +
      (ASSIGN.rows[T] ? '' : (function () { var x = PARTS.filter(function (p) { return p.loc0 === T; })[0]; return x && x !== O ? '<div class="small muted">ตาม Excel: ' + esc(x.code) + '</div>' : ''; })()) + '</div>' +
      '<label class="small muted">② Part ใหม่ที่จะใส่ช่องนี้</label>' +
      '<div class="ck-ngrow"><input id="edQ" type="text" autocapitalize="characters" autocomplete="off" placeholder="รหัส / Spec / ชื่อ" value="' + esc(E.q) + '">' +
        '<button class="btn" id="edCam">📷</button></div>' +
      '<div id="edSug"></div>' +
      '<div class="actions"><button class="btn' + (E.code === '-' ? ' primary' : '') + '" id="edEmpty">⬜ ทำให้ช่องนี้ว่าง</button></div>';
    if (N || E.code === '-') {
      edit += '<div class="ed-new">' + (N ? 'Part ใหม่: ' + partLine(N) + '<div class="small">ตอนนี้อยู่ที่: <b>' + esc(N.loc || 'รอจัดพื้นที่') + '</b></div>' : '<b>ทำให้ช่อง ' + esc(T) + ' ว่าง</b>') + '</div>';
      if (N && O && N.loc && N !== O) {
        edit += '<label class="small muted">③ Part เดิมของช่องนี้ (' + esc(O.code) + ') ให้ไปที่</label>' +
          '<div class="seg"><button data-mode="swap" class="' + (E.mode === 'swap' ? 'active' : '') + '">🔁 สลับไป ' + esc(N.loc.replace(/^H9-/, '')) + '</button>' +
          '<button data-mode="wait" class="' + (E.mode === 'wait' ? 'active' : '') + '">⏳ รอจัดพื้นที่</button></div>';
      }
      if (ch && ch.length) {
        edit += '<div class="ed-plan"><div class="small muted">จะบันทึก:</div>' + ch.map(function (c) {
          return '<div>• <b>' + esc(c.loc.replace(/^H9-/, '')) + '</b>: ' + esc(c.old || 'ว่าง') + ' → <b>' + esc(c.code || 'ว่าง') + '</b></div>';
        }).join('') +
          (O && (!N || !N.loc || E.mode === 'wait') && N !== O ? '<div class="small muted">' + esc(O.code) + ' จะไป "รอจัดพื้นที่"</div>' : '') + '</div>' +
          '<input id="edNote" class="ed-note" type="text" maxlength="200" placeholder="หมายเหตุ (ไม่บังคับ)" value="' + esc(E.note) + '">' +
          '<div class="actions"><button class="btn" id="edCancel">ยกเลิก</button><button class="btn primary" id="edSave">💾 บันทึก</button></div>';
      } else if (ch) edit += '<div class="msg idle">Part นี้อยู่ช่องนี้อยู่แล้ว</div>';
    }
  }
  edit += '</div>';

  var list = '<div class="set-card"><h4>📋 ช่องที่เปลี่ยนจาก Excel (' + changed.length + ')</h4>' +
    (changed.length ? changed.map(function (l) {
      var r = ASSIGN.rows[l], x = PARTS.filter(function (p) { return p.loc0 === l; })[0];
      return '<div class="chg-row"><div class="l">' + esc(l.replace(/^H9-/, '')) + '</div><div class="m">' +
        '<div><span class="muted">Excel:</span> ' + esc(x ? x.code : 'ว่าง') + ' → <b>' + esc(r.code || 'ว่าง') + '</b></div>' +
        '<div class="small muted">' + (r.pending ? '<b style="color:var(--amber)">⏳ รอส่ง</b> · ' : '') + esc(r.by || '') + (r.at ? ' · ' + esc(String(r.at).replace('T', ' ').slice(0, 16)) : '') + (r.note ? ' · ' + esc(r.note) : '') + '</div></div>' +
        '<button class="btn sm" data-revert="' + esc(l) + '">↩ คืนค่า</button></div>';
    }).join('') : '<div class="small muted">ยังไม่มี — ทุกช่องตรงตามไฟล์ Excel</div>') + '</div>';

  $('main').innerHTML =
    '<div class="set-head"><button class="btn sm" id="setBack">← กลับ</button><b>⚙ ตั้งค่า</b><button class="btn sm" id="setLock">🔒 ออก</button></div>' +
    '<div class="set-card"><div class="set-status ' + (api_ ? 'on' : 'off') + '">' +
      (api_ ? (navigator.onLine === false ? '📴 ออฟไลน์ — บันทึกได้ จะส่งขึ้น Google Sheet ให้เองเมื่อมีเน็ต' : '🟢 เชื่อม Google Sheet แล้ว — ทุกเครื่องเห็นการเปลี่ยน') : '🟠 ยังไม่ได้ต่อ Google Sheet — การเปลี่ยนเก็บในเครื่องนี้เท่านั้น') + '</div>' +
      (pendingCount() ? '<div class="msg warn small">⏳ รอส่งขึ้น Google Sheet ' + pendingCount() + ' รายการ <button class="btn sm" id="setFlush">ส่งตอนนี้</button></div>' : '') +
      (api_ ? '<div class="small muted">ซิงค์ล่าสุด: ' + (ASSIGN.syncedAt ? fmtTime(ASSIGN.syncedAt) : '-') + ' <button class="btn sm" id="setSync">🔄 ซิงค์</button></div>' +
        (syncErr ? '<div class="msg err small">ดึงข้อมูลจาก Google Sheet ไม่สำเร็จ: ' + esc(syncErr) + '</div>' : '') : '') +
      '<div class="small">ผู้แก้ไข: <b>👤 ' + esc(editorBy()) + '</b> <span class="muted">(เปลี่ยนได้ที่ชื่อมุมขวาบน)</span></div></div>' +
    edit + list +
    (api_ ? '<div class="set-card"><h4>🕘 ประวัติการเปลี่ยน</h4><div id="logBox"><button class="btn" id="logLoad">ดูประวัติล่าสุด</button></div></div>' : '') +
    '<details class="set-card"><summary><b>ขั้นสูง</b></summary>' +
      '<label class="small muted">Apps Script URL (/exec) — ว่าง = ใช้ค่าใน js/config.js</label>' +
      '<input id="setApi" class="ed-note" type="url" placeholder="https://script.google.com/macros/s/…/exec" value="' + esc(store('h9pl_api_url') || '') + '">' +
      '<div class="actions"><button class="btn" id="setApiSave">บันทึก URL</button></div></details>';

  // ----- ผูกปุ่ม -----
  $('setBack').onclick = function () { setTab('search'); };
  $('setLock').onclick = function () { setAdminPin(''); setTab('search'); toast('ออกจากเมนูตั้งค่าแล้ว'); };
  if ($('setFlush')) $('setFlush').onclick = function () {
    this.disabled = true;
    flushPending().then(function () { renderSettings(); }, function (e) { toast(e.message, 3500); renderSettings(); });
  };
  if ($('setSync')) $('setSync').onclick = function () {
    this.disabled = true;
    syncAssignments(true).then(function () { toast('ซิงค์แล้ว'); renderSettings(); }, function (e) { toast(e.message, 3500); renderSettings(); });
  };
  function pickLoc() {
    var l = normLoc($('edLoc').value);
    if (!locExists(l)) { toast('ไม่มี Location นี้ในผัง (เช่น L3-05)', 2500); return; }
    E.loc = l; E.code = ''; E.q = ''; E.mode = 'swap'; renderSettings(); focusPart();
  }
  function focusPart() { setTimeout(function () { if ($('edQ')) { $('edQ').focus(); $('edQ').scrollIntoView({ block: 'center', behavior: 'smooth' }); } }, 30); }
  $('edLocOk').onclick = pickLoc;
  $('edLoc').oninput = function () {
    var l = normLoc(this.value);
    if (locExists(l) && l !== E.loc && /\d{2}$/.test(this.value.trim())) { E.loc = l; E.code = ''; E.q = ''; E.mode = 'swap'; renderSettings(); focusPart(); }
  };
  $('edLoc').onkeydown = function (e) { if (e.key === 'Enter') { e.preventDefault(); pickLoc(); } };
  if ($('edQ')) {
    var sug = function () {
      var q = $('edQ').value; E.q = q;
      var list = q.trim() ? partFind(q) : [];
      $('edSug').innerHTML = list.map(function (p) {
        return '<button class="sug" data-pc="' + esc(p.code) + '">' + partLine(p) + '<div class="small muted">อยู่ที่: ' + esc(p.loc || 'รอจัดพื้นที่') + '</div></button>';
      }).join('') + (q.trim() && !list.length ? '<div class="small muted">ไม่พบ Part</div>' : '');
      Array.prototype.forEach.call(document.querySelectorAll('[data-pc]'), function (b) {
        b.onclick = function () { E.code = b.getAttribute('data-pc'); E.q = ''; E.mode = 'swap'; renderSettings(); };
      });
    };
    $('edQ').oninput = sug;
    $('edQ').onkeydown = function (e) {
      if (e.key !== 'Enter') return; e.preventDefault();
      var list = partFind($('edQ').value); if (list.length === 1) { E.code = list[0].code; E.q = ''; renderSettings(); }
    };
    if (E.q) sug();
    $('edCam').onclick = function () {
      openCameraScanner(function (t) {
        var r = resolveScan(t);
        if (!r) { beep(false); toast('ไม่พบ Part: ' + t, 3000); return; }
        beep(true); E.code = r.parts[0].code; E.q = ''; renderSettings();
      }, { title: 'สแกนป้าย Part ใหม่' });
    };
    $('edEmpty').onclick = function () { E.code = E.code === '-' ? '' : '-'; renderSettings(); };
  }
  Array.prototype.forEach.call(document.querySelectorAll('[data-mode]'), function (b) {
    b.onclick = function () { E.mode = b.getAttribute('data-mode'); renderSettings(); };
  });
  if ($('edNote')) $('edNote').oninput = function () { E.note = this.value; };
  if ($('edCancel')) $('edCancel').onclick = function () { E.code = ''; E.q = ''; E.note = ''; renderSettings(); };
  if ($('edSave')) $('edSave').onclick = function () {
    var by = editorBy();
    var changes = plannedChanges(); if (!changes || !changes.length) return;
    var btn = this; btn.disabled = true; btn.textContent = 'กำลังบันทึก…';
    saveChanges(changes, by, E.note).then(function (res) {
      beep(true); toast((res && res.queued ? '📴 บันทึกในเครื่องแล้ว จะส่งขึ้น Google Sheet เมื่อมีเน็ต: ' : 'บันทึกแล้ว: ') + changes.map(function (c) { return c.loc.replace(/^H9-/, '') + ' → ' + (c.code || 'ว่าง'); }).join(' · '), 4000);
      E.code = ''; E.q = ''; E.note = ''; renderSettings();
    }, function (e) { beep(false); toast(e.message, 4000); btn.disabled = false; btn.textContent = '💾 บันทึก'; });
  };
  Array.prototype.forEach.call(document.querySelectorAll('[data-revert]'), function (b) {
    b.onclick = function () {
      var l = b.getAttribute('data-revert'), by = editorBy();
      if (!confirm('คืนค่าช่อง ' + l + ' กลับไปตามไฟล์ Excel?')) return;
      b.disabled = true;
      revertChange(l, by).then(function (res) { toast('คืนค่า ' + l + (res && res.queued ? ' (รอส่งเมื่อมีเน็ต)' : ' แล้ว')); renderSettings(); }, function (e) { toast(e.message, 4000); b.disabled = false; });
    };
  });
  if ($('logLoad')) $('logLoad').onclick = function () {
    $('logBox').innerHTML = '<div class="small muted">กำลังโหลด…</div>';
    api('getLog', 50).then(function (rows) {
      var A = { assign: 'ใส่', clear: 'ทำให้ว่าง', revert: 'คืนค่า Excel' };
      $('logBox').innerHTML = rows.length ? rows.map(function (r) {
        return '<div class="log-row"><b>' + esc(String(r.loc).replace(/^H9-/, '')) + '</b> ' + esc(A[r.action] || r.action) + ': ' + esc(r.old || 'ว่าง') + ' → <b>' + esc(r.code || (r.action === 'revert' ? 'ตาม Excel' : 'ว่าง')) + '</b>' +
          '<div class="small muted">' + esc(String(r.at).replace('T', ' ').slice(0, 16)) + ' · ' + esc(r.by) + (r.note ? ' · ' + esc(r.note) : '') + '</div></div>';
      }).join('') : '<div class="small muted">ยังไม่มีประวัติ</div>';
    }, function (e) { $('logBox').innerHTML = '<div class="small" style="color:var(--red)">' + esc(e.message) + '</div>'; });
  };
  $('setApiSave').onclick = function () {
    var v = $('setApi').value.trim();
    if (v && !/^https:\/\/script\.google\.com\/.+\/exec$/.test(v)) { toast('URL ต้องเป็น https://script.google.com/…/exec', 3000); return; }
    store('h9pl_api_url', v); ASSIGN.syncedAt = 0; setAdminPin('');
    toast('บันทึก URL แล้ว — กรุณาใส่รหัสใหม่', 3000); setTab('search');
    syncAssignments(true).catch(function (e) { toast(e.message, 3500); });
  };
}

// บันทึก: มี Google Sheet → ส่งเซิร์ฟเวอร์ (ตรวจรหัส + เก็บประวัติ) แล้วใช้ข้อมูลที่ตอบกลับ · ไม่มี → เก็บในเครื่อง
// ออนไลน์ก่อน: ส่งขึ้น Google Sheet ทันที · ส่งไม่ได้เพราะไม่มีเน็ต → เก็บในคิว (ใช้ผลในเครื่องทันที) แล้วส่งให้เองเมื่อกลับมามีเน็ต · คืนค่า { queued: true }
function sendAfterQueue(op) {
  queueOp(op);
  return flushPending().then(function () { return pendingCount() ? { queued: true } : {}; }, function () { return { queued: true }; });
}
function saveChanges(changes, by, note) {
  var op = { type: 'save', changes: changes, by: by, note: note || '', at: nowLocal(), pin: adminPin() };
  if (!apiUrl()) { applyOpLocal(op); ASSIGN.local = true; saveAssign(); applyAssignments(); return Promise.resolve({}); }
  if (pendingCount()) return sendAfterQueue(op); // มีรายการค้าง → ส่งตามลำดับ
  return api('saveAssignments', op.pin, by, changes, op.note).then(setAssignRows, function (e) {
    if (!e.network) throw e;
    queueOp(op); return { queued: true };
  });
}
function revertChange(loc, by) {
  var r = ASSIGN.rows[loc] || {}, x = PARTS.filter(function (p) { return p.loc0 === loc; })[0];
  var op = { type: 'revert', loc: loc, by: by, oldCode: r.code || '', newCode: x ? x.code : '', at: nowLocal(), pin: adminPin() };
  if (!apiUrl()) { applyOpLocal(op); saveAssign(); applyAssignments(); return Promise.resolve({}); }
  if (pendingCount()) return sendAfterQueue(op); // มีรายการค้าง → ส่งตามลำดับ
  return api('revertAssignment', op.pin, by, loc, op.oldCode, op.newCode).then(setAssignRows, function (e) {
    if (!e.network) throw e;
    queueOp(op); return { queued: true };
  });
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
document.addEventListener('visibilitychange', function () { if (!document.hidden) { checkUpdate(); syncAssignments().catch(function () {}); } });
syncAssignments(true).catch(function () { /* ออฟไลน์ → ใช้ข้อมูลที่เก็บในเครื่อง */ });
window.addEventListener('online', function () { updateNet(); syncAssignments(true).catch(function () {}); if (S.tab === 'settings' && $('modal').classList.contains('hidden')) renderSettings(); });
window.addEventListener('offline', function () { updateNet(); if (S.tab === 'settings' && $('modal').classList.contains('hidden')) renderSettings(); });
// Service Worker: เก็บไฟล์แอปไว้ในเครื่อง → เปิดแอปได้แม้ไม่มีเน็ต (ต้องเปิดผ่าน https หรือ localhost)
if ('serviceWorker' in navigator && /^(https:|http:\/\/localhost)/.test(location.href)) {
  window.addEventListener('load', function () {
    navigator.serviceWorker.register('sw.js').catch(function () { /* ไม่รองรับ → ใช้แบบออนไลน์ */ });
    // ตอนมีเน็ต: โหลดตัวอ่านบาร์โค้ด (zxing / jsQR / ตัวกู้ QR) เก็บไว้ล่วงหน้า → สแกนด้วยกล้องได้แม้ออฟไลน์
    setTimeout(function () {
      if (navigator.onLine === false) return;
      getDetectors().then(function (list) {
        var c = document.createElement('canvas'); c.width = c.height = 32;
        list.forEach(function (e) { try { e.detect(c).catch(function () {}); } catch (x) { /* ไม่เป็นไร */ } });
      }, function () {});
    }, 4000);
  });
}
setInterval(checkUpdate, 5 * 60 * 1000);
checkUpdate();
})();
