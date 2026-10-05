/* Кайрос · окно «Обновить данные»: выбор файлов 1С → пересчёт в браузере → сохранение в браузере (IndexedDB).
 * Запуск приложения идёт через KairosUI.boot(): сначала проверяем, нет ли загруженных ранее свежих данных. */
(function (root) {
'use strict';
var DB = 'kairos-data', STORE = 'kv', KEY = 'krs';
var UI = {}, emb = root.KRS_DATA, files = {}, result = null, busy = false;

function idb() {
  return new Promise(function (res, rej) {
    try {
      var rq = indexedDB.open(DB, 1);
      rq.onupgradeneeded = function () { rq.result.createObjectStore(STORE); };
      rq.onsuccess = function () { res(rq.result); };
      rq.onerror = function () { rej(rq.error); };
      rq.onblocked = function () { rej(new Error('blocked')); };
    } catch (e) { rej(e); }
  });
}
function idbGet() { return idb().then(function (db) { return new Promise(function (res, rej) { var r = db.transaction(STORE).objectStore(STORE).get(KEY); r.onsuccess = function () { res(r.result); }; r.onerror = function () { rej(r.error); }; }); }); }
function idbPut(v) { return idb().then(function (db) { return new Promise(function (res, rej) { var t = db.transaction(STORE, 'readwrite'); t.objectStore(STORE).put(v, KEY); t.oncomplete = res; t.onerror = function () { rej(t.error); }; t.onabort = function () { rej(t.error || new Error('abort')); }; }); }); }
function idbDel() { return idb().then(function (db) { return new Promise(function (res, rej) { var t = db.transaction(STORE, 'readwrite'); t.objectStore(STORE).delete(KEY); t.oncomplete = res; t.onerror = function () { rej(t.error); }; }); }); }

/* ───────── запуск приложения ───────── */
UI.boot = function () {
  var started = false;
  function start() {
    if (started) return; started = true;
    var s = document.getElementById('krs-app'); if (!s) return;
    var n = document.createElement('script'); n.textContent = s.textContent; document.body.appendChild(n);
  }
  var timer = setTimeout(start, 2500);
  idbGet().then(function (st) {
    if (st && st.data && st.data.tx && String(st.data.dataAsOf || '') >= String(emb.dataAsOf || '')) {
      var d = st.data;
      ['geoRef', 'ruMap', 'pop', 'plan'].forEach(function (k) { if (emb[k] !== undefined) d[k] = emb[k]; });     // справочники всегда из свежесобранного файла
      if (d.refs && emb.refs) { d.refs.price = emb.refs.price; d.refs.artPrice = emb.refs.artPrice; }
      root.KRS_DATA = d; root.__krsSource = { kind: 'uploaded', savedAt: st.savedAt };
    } else root.__krsSource = { kind: 'embedded' };
  }).catch(function () { root.__krsSource = { kind: 'embedded', idbFail: true }; }).then(function () { clearTimeout(timer); start(); });
};

/* ───────── оформление ───────── */
var CSS = '.krs-chip{margin-left:auto;display:inline-flex;align-items:center;gap:8px;background:#fff;border:1px solid var(--line-strong,#d8dbe0);border-radius:10px;padding:7px 12px;font:inherit;font-size:12.5px;font-weight:600;color:#2A2A2A;cursor:pointer}' +
  '.krs-chip:hover{border-color:#2A2A2A}.krs-chip i{width:8px;height:8px;border-radius:50%;background:#2e9e5b;display:inline-block}.krs-chip i.stale{background:#C7001F}.krs-chip small{font-weight:500;color:#727987;font-size:11.5px}' +
  '.ki-ov{position:fixed;inset:0;background:rgba(42,42,42,.55);z-index:99999;display:flex;align-items:flex-start;justify-content:center;overflow:auto;padding:5vh 16px}' +
  '.ki-box{background:#fff;border-radius:14px;width:min(720px,100%);box-shadow:0 20px 60px rgba(0,0,0,.35);font-family:Montserrat,system-ui,sans-serif;color:#2A2A2A}' +
  '.ki-h{padding:20px 24px 14px;border-bottom:1px solid #eceef1;display:flex;align-items:flex-start;gap:12px}.ki-h h2{margin:0;font-family:Unbounded,Montserrat,sans-serif;font-size:18px;font-weight:700}' +
  '.ki-h p{margin:4px 0 0;font-size:12.5px;color:#727987;line-height:1.45}.ki-x{margin-left:auto;border:0;background:none;font-size:22px;line-height:1;cursor:pointer;color:#727987}' +
  '.ki-b{padding:18px 24px}.ki-now{background:#f6f7f9;border-radius:10px;padding:10px 14px;font-size:12.5px;line-height:1.5;margin-bottom:16px}.ki-now b{font-weight:700}' +
  '.ki-drop{border:2px dashed #c9ced6;border-radius:12px;padding:16px;text-align:center;font-size:13px;color:#727987;margin-bottom:14px;transition:.15s}.ki-drop.over{border-color:#C7001F;background:#fff6f7}' +
  '.ki-row{display:grid;grid-template-columns:1fr auto;gap:4px 12px;align-items:center;padding:11px 0;border-top:1px solid #eceef1}.ki-row:first-of-type{border-top:0}' +
  '.ki-row b{font-size:13.5px}.ki-row em{font-style:normal;font-size:11px;font-weight:700;letter-spacing:.3px;text-transform:uppercase;margin-left:8px;color:#C7001F}.ki-row em.opt{color:#727987}' +
  '.ki-row small{grid-column:1;font-size:11.5px;color:#727987;line-height:1.4}.ki-file{grid-column:1;font-size:12.5px;color:#2e7d4f;font-weight:600;word-break:break-all}' +
  '.ki-btn{border:1px solid #2A2A2A;background:#fff;color:#2A2A2A;border-radius:9px;padding:8px 14px;font:inherit;font-size:12.5px;font-weight:600;cursor:pointer}.ki-btn:hover{background:#f1f2f4}' +
  '.ki-btn.pri{background:#2A2A2A;color:#fff}.ki-btn.pri:hover{background:#000}.ki-btn[disabled]{opacity:.4;cursor:default}.ki-btn.link{border:0;background:none;text-decoration:underline;padding:6px 0;color:#727987;font-weight:500}' +
  '.ki-f{padding:14px 24px 20px;border-top:1px solid #eceef1;display:flex;gap:10px;align-items:center;flex-wrap:wrap}.ki-f .sp{flex:1}' +
  '.ki-bar{height:8px;background:#eceef1;border-radius:6px;overflow:hidden;margin:10px 0 6px}.ki-bar i{display:block;height:100%;width:0;background:linear-gradient(90deg,#FDCF08,#C7001F);transition:width .2s}' +
  '.ki-kv{display:grid;grid-template-columns:repeat(2,1fr);gap:10px;margin:4px 0 14px}.ki-kv div{background:#f6f7f9;border-radius:10px;padding:10px 12px}.ki-kv span{display:block;font-size:11px;color:#727987;text-transform:uppercase;letter-spacing:.3px}.ki-kv b{font-size:16px}' +
  '.ki-w{margin:8px 0;padding:10px 12px;border-radius:10px;font-size:12.5px;line-height:1.45;background:#fff8e1;border:1px solid #f1dc8a}.ki-w.err{background:#fff0f1;border-color:#f0b4bb;color:#8a0016}.ki-ok{background:#effaf3;border:1px solid #bfe6cd}' +
  '@media(max-width:560px){.ki-kv{grid-template-columns:1fr}}';

function el(html) { var d = document.createElement('div'); d.innerHTML = html; return d.firstChild; }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (m) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[m]; }); }
function mln(v) { return (v / 1e6).toLocaleString('ru-RU', { maximumFractionDigits: 1 }) + ' млн ₽'; }
function ruDate(d) { return d ? new Date(d).toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : ''; }
function mb(n) { return (n / 1048576).toFixed(1) + ' МБ'; }

var KINDS = [
  { id: 'sales', title: 'Продажи по бизнес-регионам', req: true, re: /продаж/i,
    hint: '1С → отчёт «Продажи по бизнес регионам», период с 01.01.2024 по сегодня, подразделения Москва / HeadRock МСК / Владивосток / HeadRock Владивосток / СЕТИ, бренды ENKI, KRONbuild, HeadRock. Сохранить в .xlsx без изменений. Можно выбрать два файла: полную историю и свежий период — каждый год берётся из файла, где по нему больше данных.' },
  { id: 'stock', title: 'Остатки и доступность по сериям', req: false, re: /остатк|доступн/i, hint: 'Отчёт «Остатки и доступность по сериям», все склады. Если не выбирать — останутся прежние остатки.' },
  { id: 'expiry', title: 'Товары с окончанием срока годности', req: false, re: /годност|срок/i, hint: 'Отчёт «Товары на складах с окончанием срока годности». Если не выбирать — останутся прежние сроки.' }
];

function close() { var o = document.getElementById('ki-ov'); if (o) o.remove(); }
function current() { return root.KRS_DATA || {}; }
function nowText() {
  var D = current(), src = root.__krsSource || {}, ing = D.__ingest;
  if (src.kind === 'uploaded' && ing) {
    var f = ing.files || {};
    return 'Сейчас в дашборде: <b>данные на ' + esc(D.dataAsOfHuman) + '</b>. Загружено в этом браузере ' + esc(ruDate(src.savedAt)) +
      ' (продажи: ' + esc(f.sales || '—') + (f.stock ? '; остатки: ' + esc(f.stock) : '') + (f.expiry ? '; сроки: ' + esc(f.expiry) : '') + ').';
  }
  return 'Сейчас в дашборде: <b>данные на ' + esc(D.dataAsOfHuman || '—') + '</b> — они зашиты в этот файл при сборке. Загруженные здесь файлы сохранятся только в этом браузере.' + (src.idbFail ? ' <b style="color:#C7001F">Браузер не даёт сохранять данные (режим файла/приватное окно) — загрузка не сработает.</b>' : '');
}

function render() {
  var box = document.getElementById('ki-box'); if (!box) return;
  var ok = !!(files.sales && [].concat(files.sales).length);
  var rows = KINDS.map(function (k) {
    var f = files[k.id], fl = Array.isArray(f) ? f : (f ? [f] : []);
    return '<div class="ki-row"><div><b>' + k.title + '</b><em class="' + (k.req ? '' : 'opt') + '">' + (k.req ? 'обязательно' : 'по желанию') + '</em></div>' +
      '<div><button class="ki-btn" data-pick="' + k.id + '">' + (fl.length ? 'Заменить' : 'Выбрать файл') + '</button></div>' +
      (fl.length ? fl.map(function (x) { return '<div class="ki-file">✓ ' + esc(x.name) + ' · ' + mb(x.size) + '</div>'; }).join('') : '<small>' + k.hint + '</small>') + '</div>';
  }).join('');
  box.innerHTML = '<div class="ki-h"><div><h2>Обновить данные из 1С</h2><p>Выберите свежие выгрузки — дашборд пересчитается прямо в браузере, ничего никуда не отправляется.</p></div><button class="ki-x" data-x>×</button></div>' +
    '<div class="ki-b"><div class="ki-now">' + nowText() + '</div>' +
    '<div class="ki-drop" id="ki-drop">Перетащите сюда файлы (можно сразу все три) или выберите ниже</div>' + rows +
    '<div id="ki-out"></div></div>' +
    '<div class="ki-f"><button class="ki-btn link" data-act="exp">Сохранить текущие данные в файл</button><button class="ki-btn link" data-act="imp">Открыть файл данных</button>' +
    ((root.__krsSource || {}).kind === 'uploaded' ? '<button class="ki-btn link" data-act="reset">Вернуть данные из файла</button>' : '') +
    '<span class="sp"></span><button class="ki-btn" data-x>Закрыть</button><button class="ki-btn pri" id="ki-go" ' + (ok && !busy ? '' : 'disabled') + '>Загрузить и пересчитать</button></div>';
  bindBox(box);
}

function pickFile(accept, cb) {
  var multi = accept === 'multi', i = document.createElement('input');
  i.type = 'file'; i.accept = multi ? '.xlsx' : (accept || '.xlsx'); i.multiple = multi; i.style.cssText = 'position:fixed;left:-9999px;opacity:0';
  document.body.appendChild(i);                                         // в DOM, иначе браузер может «забыть» выбор
  i.onchange = function () { var l = [].slice.call(i.files || []); i.remove(); if (l.length) cb(l); };
  i.click();
}
function assign(list) {
  var unknown = [];
  list.forEach(function (f) {
    if (!/\.xlsx$/i.test(f.name)) { unknown.push(f.name + ' (нужен .xlsx)'); return; }
    var k = KINDS.filter(function (x) { return x.re.test(f.name); })[0];
    if (k && k.id === 'sales') { files.sales = (files.sales || []).filter(function (x) { return x.name !== f.name; }).concat([f]); }
    else if (k) files[k.id] = f; else unknown.push(f.name);
  });
  render();
  if (unknown.length) out('<div class="ki-w">Не удалось определить тип файла: ' + unknown.map(esc).join(', ') + '. Нажмите «Выбрать файл» в нужной строке.</div>');
}
function out(h) { var o = document.getElementById('ki-out'); if (o) o.innerHTML = h; }

function bindBox(box) {
  [].forEach.call(box.querySelectorAll('[data-x]'), function (b) { b.onclick = close; });
  [].forEach.call(box.querySelectorAll('[data-pick]'), function (b) {
    b.onclick = function () { var id = b.getAttribute('data-pick'); pickFile(id === 'sales' ? 'multi' : '.xlsx', function (l) { if (id === 'sales') files.sales = l; else files[id] = l[0]; render(); }); };
  });
  var dz = box.querySelector('#ki-drop');
  if (dz) {
    dz.onclick = function () { pickFile('multi', assign); };
    dz.ondragover = function (e) { e.preventDefault(); dz.classList.add('over'); };
    dz.ondragleave = function () { dz.classList.remove('over'); };
    dz.ondrop = function (e) { e.preventDefault(); dz.classList.remove('over'); assign([].slice.call(e.dataTransfer.files)); };
  }
  var go = box.querySelector('#ki-go'); if (go) go.onclick = run;
  [].forEach.call(box.querySelectorAll('[data-act]'), function (b) {
    b.onclick = function () {
      var a = b.getAttribute('data-act');
      if (a === 'exp') exportData(); else if (a === 'imp') pickFile('.kairos,.gz', importData);
      else if (a === 'reset') { if (confirm('Удалить загруженные в этом браузере данные и вернуться к данным из файла?')) idbDel().then(function () { location.reload(); }); }
    };
  });
}

function run() {
  if (busy || !(files.sales && [].concat(files.sales).length)) return; busy = true;
  var go = document.getElementById('ki-go'); if (go) go.disabled = true;
  out('<div id="ki-st" style="font-size:13px;font-weight:600">Начинаю…</div><div class="ki-bar"><i id="ki-pb"></i></div><small style="color:#727987">Файл продаж большой (≈18 МБ) — обычно это занимает 10–40 секунд. Не закрывайте вкладку.</small>');
  var base = current(), t0 = Date.now();
  root.KairosIngest.run(files, base, function (m, p) {
    var s = document.getElementById('ki-st'), b = document.getElementById('ki-pb');
    if (s) s.textContent = m; if (b) b.style.width = Math.round((p || 0) * 100) + '%';
  }).then(function (K) {
    busy = false; result = K; report(K, base, (Date.now() - t0) / 1000);
  }).catch(function (e) {
    busy = false; console.error(e);
    out('<div class="ki-w err"><b>Не получилось.</b> ' + esc(e && e.message || e) + '</div>'); var g = document.getElementById('ki-go'); if (g) g.disabled = false;
  });
}

function report(K, base, sec) {
  var ing = K.__ingest, inf = ing.info, ov = K.overview, ob = base.overview || {}, warn = ing.warns.slice();
  if (inf.newClients > 0) warn.push('Новых клиентов в выгрузке: ' + inf.newClients + (inf.noCrm ? ', из них без карточки CRM: ' + inf.noCrm : '') + '. Регион и город берутся из отчёта продаж, менеджер — по правилам округов (для Москвы — «Не назначен», пока клиента нет в CRM).');
  var delta = ov.sales - (ob.sales || 0), sameYear = (base.refDate || '').slice(0, 4) === (K.refDate || '').slice(0, 4);
  out('<div class="ki-w ki-ok"><b>Готово за ' + Math.round(sec) + ' с.</b> Данные на <b>' + esc(K.dataAsOfHuman) + '</b>' + (inf.periodTo ? ' (период в отчёте: ' + inf.periodFrom + ' – ' + inf.periodTo + ')' : '') + '.</div>' +
    '<div class="ki-kv"><div><span>Продажи ' + esc((K.refDate || '').slice(0, 4)) + ', нетто</span><b>' + mln(ov.sales) + '</b>' + (sameYear ? '<small style="color:#727987"> было ' + mln(ob.sales || 0) + ' (' + (delta >= 0 ? '+' : '−') + mln(Math.abs(delta)).replace(' ₽', '') + ')</small>' : '') + '</div>' +
    '<div><span>Клиентов с продажами / строк продаж</span><b>' + K.clients.length.toLocaleString('ru-RU') + ' / ' + K.tx.length.toLocaleString('ru-RU') + '</b></div></div>' +
    (inf.filter ? '<div style="font-size:11.5px;color:#727987;margin-bottom:8px">Отбор в файле: ' + esc(inf.filter) + '</div>' : '') +
    (ing.notes.length ? '<div style="font-size:12px;color:#727987;margin-bottom:6px">' + ing.notes.map(esc).join(' · ') + '</div>' : '') +
    warn.map(function (w) { return '<div class="ki-w">' + esc(w) + '</div>'; }).join('') +
    '<div style="margin-top:10px"><button class="ki-btn pri" id="ki-apply">Применить и открыть дашборд</button> <span style="font-size:12px;color:#727987">Данные сохранятся в этом браузере.</span></div>');
  var a = document.getElementById('ki-apply');
  if (a) a.onclick = function () {
    a.disabled = true; a.textContent = 'Сохраняю…';
    idbPut({ data: result, savedAt: Date.now() }).then(function () { location.reload(); }).catch(function (e) {
      a.disabled = false; a.textContent = 'Применить и открыть дашборд';
      out('<div class="ki-w err">Браузер не дал сохранить данные (' + esc(e && e.message || e) + '). Чаще всего это приватное окно или запрет на хранение для локальных файлов. Откройте дашборд в обычном окне Chrome/Edge.</div>');
    });
  };
}

function gz(str) { return new Response(new Blob([str]).stream().pipeThrough(new CompressionStream('gzip'))).blob(); }
function gunzip(blob) { return new Response(blob.stream().pipeThrough(new DecompressionStream('gzip'))).text(); }
function exportData() {
  var D = current();
  out('<div style="font-size:13px">Готовлю файл…</div>');
  gz(JSON.stringify(D)).then(function (b) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(b); a.download = 'kairos_data_' + (D.dataAsOf || 'now') + '.kairos'; document.body.appendChild(a); a.click(); setTimeout(function () { a.remove(); URL.revokeObjectURL(a.href); }, 2000);
    out('<div class="ki-w ki-ok">Файл данных сохранён (' + mb(b.size) + '). Коллега откроет его в своём дашборде кнопкой «Открыть файл данных».</div>');
  });
}
function importData(l) {
  var f = l[0]; out('<div style="font-size:13px">Читаю файл данных…</div>');
  gunzip(f).then(function (t) {
    var d = JSON.parse(t);
    if (!d || !d.tx || !d.overview || !d.clients) throw new Error('это не файл данных Кайроса');
    return idbPut({ data: d, savedAt: Date.now() });
  }).then(function () { location.reload(); }).catch(function (e) { out('<div class="ki-w err">Не получилось открыть: ' + esc(e && e.message || e) + '</div>'); });
}

UI.open = function () {
  close(); files = {}; result = null; busy = false;
  if (!document.getElementById('ki-css')) { var s = document.createElement('style'); s.id = 'ki-css'; s.textContent = CSS; document.head.appendChild(s); }
  var ov = el('<div class="ki-ov" id="ki-ov"><div class="ki-box" id="ki-box" role="dialog" aria-modal="true"></div></div>');
  ov.addEventListener('mousedown', function (e) { if (e.target === ov && !busy) close(); });
  document.body.appendChild(ov); render();
};
UI.chip = function () {
  var D = current(), asof = D.dataAsOf, diff = asof ? Math.round((new Date() - new Date(asof)) / 86400000) : null, stale = diff != null && diff > 3;
  return '<button class="krs-chip" onclick="krsOpenIngest()" title="Загрузить свежие выгрузки 1С"><i class="' + (stale ? 'stale' : '') + '"></i>Данные на ' + esc(D.dataAsOfHuman || '—') +
    '<small>' + (stale ? diff + ' дн. назад · ' : '') + 'обновить</small></button>';
};
if (!document.getElementById('ki-css')) { var st = document.createElement('style'); st.id = 'ki-css'; st.textContent = CSS; (document.head || document.documentElement).appendChild(st); }
root.krsOpenIngest = UI.open; root.KairosUI = UI;
})(window);
