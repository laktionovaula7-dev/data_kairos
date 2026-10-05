/* Кайрос · выбор периода: кнопка-плашка + всплывающее окно (пресеты, год, сетка месяцев, свои даты).
 * Работает поверх существующей модели window.PERIOD: ytd | yГГГГ | ГГГГ-ММ | custom {from,to}.
 * Применение идёт через те же обработчики (#periodSel / #pApply), поэтому пересчёт и перерисовка не меняются. */
(function (root) {
'use strict';
var MN = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
var st = { open: false, year: null, first: null };
var pad = function (n) { return (n < 10 ? '0' : '') + n; };
var esc = function (s) { return String(s).replace(/[&<>"]/g, function (m) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[m]; }); };
function D() { return root.KRS_DATA || {}; }
function ref() {
  var r = String(D().refDate || D().dataAsOf || '').replace(/-/g, '');
  if (r.length === 8) return { y: +r.slice(0, 4), m: +r.slice(4, 6), d: +r.slice(6) };
  var n = new Date(); return { y: n.getFullYear(), m: n.getMonth() + 1, d: n.getDate() };
}
function years() {
  var sum = {}, ym = D().ym || {}, r = ref();
  Object.keys(ym).forEach(function (k) { var y = +k.slice(0, 4); sum[y] = (sum[y] || 0) + ym[k]; });
  var ys = {}; Object.keys(sum).forEach(function (y) { if (sum[y] > 5e6) ys[y] = 1; });
  var out = Object.keys(ys).map(Number).filter(function (y) { return y >= r.y - 3 && y <= r.y; }).sort();
  if (out.indexOf(r.y) < 0) out.push(r.y);
  return out;
}
function lastDay(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
/* текущий период → [yearFrom, monthFrom, dayFrom, yearTo, monthTo, dayTo] */
function rangeOf(P) {
  var r = ref();
  P = P || { mode: 'ytd' };
  if (P.mode === 'custom' && P.from && P.to) { var a = P.from.split('-').map(Number), b = P.to.split('-').map(Number); return [a[0], a[1], a[2], b[0], b[1], b[2]]; }
  if (/^y\d{4}$/.test(P.mode)) { var y = +P.mode.slice(1); return [y, 1, 1, y, 12, 31]; }
  if (/^\d{4}-\d{2}$/.test(P.mode)) { var yy = +P.mode.slice(0, 4), mm = +P.mode.slice(5); return [yy, mm, 1, yy, mm, lastDay(yy, mm)]; }
  return [r.y, 1, 1, r.y, 12, 31];
}
function ruD(y, m, d) { return pad(d) + '.' + pad(m) + '.' + y; }
function label() {
  var P = root.PERIOD || { mode: 'ytd' }, r = rangeOf(P), r0 = ref();
  if (P.mode === 'ytd') return 'Весь ' + r0.y;
  if (/^y\d{4}$/.test(P.mode)) return 'Весь ' + r[0];
  if (/^\d{4}-\d{2}$/.test(P.mode)) return MN[r[1] - 1] + ' ' + r[0];
  var wholeMonths = r[2] === 1 && (r[5] === lastDay(r[3], r[4]) || (r[3] === r0.y && r[4] === r0.m && r[5] >= r0.d));
  if (wholeMonths) {
    if (r[0] === r[3]) return r[1] === 1 && r[4] === 12 ? 'Весь ' + r[0] : MN[r[1] - 1] + ' – ' + MN[r[4] - 1] + ' ' + r[0];
    return MN[r[1] - 1] + ' ' + r[0] + ' – ' + MN[r[4] - 1] + ' ' + r[3];
  }
  return ruD(r[0], r[1], r[2]) + ' – ' + ruD(r[3], r[4], r[5]);
}
function ymKey(y, m) { return y * 12 + (m - 1); }

function popHtml() {
  var r0 = ref(), ys = years(), P = root.PERIOD || { mode: 'ytd' }, cur = rangeOf(P);
  if (st.year == null) st.year = (cur[0] === cur[3] ? cur[0] : r0.y);
  var cLo = ymKey(cur[0], cur[1]), cHi = ymKey(cur[3], cur[4]), fk = st.first;
  var q0 = Math.floor((r0.m - 1) / 3) * 3 + 1;
  var pre = [['all', 'Весь год'], ['q', 'Квартал'], ['3m', 'Последние 3 мес.'], ['m', 'Текущий месяц']]
    .map(function (p) { return '<button class="kpp-p" data-kpp="pre:' + p[0] + '">' + p[1] + '</button>'; }).join('');
  var yrs = ys.map(function (y) { return '<button class="kpp-y' + (y === st.year ? ' on' : '') + '" data-kpp="year:' + y + '">' + y + '</button>'; }).join('');
  var grid = MN.map(function (n, i) {
    var m = i + 1, k = ymKey(st.year, m), off = (st.year > r0.y) || (st.year === r0.y && m > r0.m);
    var cls = 'kpp-m';
    if (off) cls += ' off';
    else if (fk != null) { if (k === fk) cls += ' on'; }
    else if (k >= cLo && k <= cHi) cls += (k === cLo || k === cHi) ? ' on' : ' in';
    return '<button class="' + cls + '" ' + (off ? 'disabled' : 'data-kpp="m:' + st.year + '-' + pad(m) + '"') + '>' + n + '</button>';
  }).join('');
  var hint = fk != null ? 'Выберите последний месяц диапазона или нажмите на тот же месяц ещё раз — будет один месяц.' : 'Клик по двум месяцам — диапазон. Месяцы без данных затемнены.';
  var f = cur, from = f[0] + '-' + pad(f[1]) + '-' + pad(f[2]), to = f[3] + '-' + pad(f[4]) + '-' + pad(Math.min(f[5], f[3] === r0.y && f[4] === r0.m ? r0.d : f[5]));
  return '<div class="kpp-pre">' + pre + '</div><div class="kpp-yrs">' + yrs + '</div><div class="kpp-grid">' + grid + '</div><div class="kpp-note">' + hint + '</div>'
    + '<div class="kpp-cus"><b>Свой период</b><label>с <input type="date" id="kppFrom" value="' + from + '"></label><label>по <input type="date" id="kppTo" value="' + to + '"></label><button class="kpp-ap" data-kpp="custom">Применить</button></div>'
    + '<div class="kpp-err" id="kppErr"></div>';
}
function refreshPops() {
  [].forEach.call(document.querySelectorAll('.kpp'), function (el) {
    el.classList.toggle('open', st.open);
    var p = el.querySelector('.kpp-pop'); if (p) p.innerHTML = st.open ? popHtml() : '';
  });
}
function apply(mode, from, to) {
  st.open = false; st.first = null;
  if (mode === 'custom') {
    var f = document.createElement('input'), t = document.createElement('input'), b = document.createElement('button');
    f.type = t.type = 'date'; f.id = 'pFrom'; t.id = 'pTo'; b.id = 'pApply'; f.value = from; t.value = to;
    [f, t, b].forEach(function (e) { e.style.display = 'none'; document.body.appendChild(e); });
    b.click(); [f, t, b].forEach(function (e) { e.remove(); });
  } else {
    var s = document.createElement('select'); s.id = 'periodSel'; s.style.display = 'none';
    s.innerHTML = '<option value="' + mode + '" selected>x</option>'; document.body.appendChild(s);
    s.dispatchEvent(new Event('change', { bubbles: true })); s.remove();
  }
  refreshPops();
}
function rangeApply(a, b) {   // a,b: [y,m]; a<=b
  var r0 = ref();
  if (a[0] === b[0] && a[1] === b[1]) return apply(a[0] + '-' + pad(a[1]), '', '');
  if (a[0] === b[0] && a[1] === 1 && b[1] === 12) return apply(a[0] === r0.y ? 'ytd' : 'y' + a[0], '', '');
  apply('custom', a[0] + '-' + pad(a[1]) + '-01', b[0] + '-' + pad(b[1]) + '-' + pad(lastDay(b[0], b[1])));
}
function shiftMonths(y, m, k) { var t = y * 12 + (m - 1) + k; return [Math.floor(t / 12), (t % 12) + 1]; }

document.addEventListener('click', function (e) {
  var el = e.target.closest ? e.target.closest('[data-kpp]') : null;
  if (!el) {
    if (st.open && !(e.target.closest && e.target.closest('.kpp'))) { st.open = false; st.first = null; refreshPops(); }
    return;
  }
  var a = el.getAttribute('data-kpp'), r0 = ref();
  if (a === 'toggle') { st.open = !st.open; st.first = null; if (st.open) st.year = null; refreshPops(); return; }
  if (a.indexOf('year:') === 0) { st.year = +a.slice(5); refreshPops(); return; }
  if (a.indexOf('m:') === 0) {
    var p = a.slice(2).split('-').map(Number), k = ymKey(p[0], p[1]);
    if (st.first == null) { st.first = k; refreshPops(); return; }
    var lo = Math.min(st.first, k), hi = Math.max(st.first, k);
    return rangeApply([Math.floor(lo / 12), (lo % 12) + 1], [Math.floor(hi / 12), (hi % 12) + 1]);
  }
  if (a.indexOf('pre:') === 0) {
    var t = a.slice(4), y = st.year || r0.y;
    if (t === 'all') return rangeApply([y, 1], [y, 12]);
    if (t === 'q') { var q0 = Math.floor((r0.m - 1) / 3) * 3 + 1; return rangeApply([r0.y, q0], [r0.y, q0 + 2]); }
    if (t === '3m') return rangeApply(shiftMonths(r0.y, r0.m, -2), [r0.y, r0.m]);
    return rangeApply([r0.y, r0.m], [r0.y, r0.m]);
  }
  if (a === 'custom') {
    var f = document.getElementById('kppFrom').value, to = document.getElementById('kppTo').value, er = document.getElementById('kppErr');
    if (!f || !to) { er.textContent = 'Укажите обе даты.'; return; }
    if (f > to) { er.textContent = 'Дата «с» позже даты «по».'; return; }
    return apply('custom', f, to);
  }
});
document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && st.open) { st.open = false; st.first = null; refreshPops(); } });

root.krsPeriodUI = function () {
  return '<div class="kpp' + (st.open ? ' open' : '') + '"><button type="button" class="kpp-pill" data-kpp="toggle">' + esc(label()) + ' <i>▾</i></button><div class="kpp-pop">' + (st.open ? popHtml() : '') + '</div></div>';
};

var css = '.kpp{position:relative;display:inline-block}.kpp-pill{display:inline-flex;align-items:center;gap:8px;border:1px solid #23242A;border-radius:999px;padding:9px 16px;font:700 14px Montserrat,system-ui,sans-serif;color:#23242A;background:#fff;cursor:pointer;white-space:nowrap}'
  + '.kpp-pill i{font-style:normal;color:#86888D;font-weight:500;font-size:12px}.kpp-pill:hover{background:#FBFAF7}'
  + '.kpp-pop{display:none;position:absolute;top:calc(100% + 8px);left:0;z-index:300;background:#fff;border:1px solid #E6E6E1;border-radius:16px;box-shadow:0 18px 50px rgba(0,0,0,.16);padding:16px;width:440px;max-width:92vw;font-family:Montserrat,system-ui,sans-serif;text-transform:none;letter-spacing:0;color:#23242A;font-weight:500}'
  + '.kpp.open .kpp-pop{display:block}.kpp-pre{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px}'
  + '.kpp-p{padding:7px 10px;white-space:nowrap;border-radius:999px;border:1px solid #E6E6E1;background:#fff;font:600 12px Montserrat,sans-serif;cursor:pointer;color:#23242A}.kpp-p:hover{background:#23242A;color:#fff;border-color:#23242A}'
  + '.kpp-yrs{display:flex;gap:6px;margin-bottom:12px}.kpp-y{padding:7px 16px;border-radius:999px;border:0;background:#F4F3EF;font:700 13px Montserrat,sans-serif;cursor:pointer;color:#23242A}.kpp-y.on{background:#23242A;color:#fff}'
  + '.kpp-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}.kpp-m{padding:10px 0;border-radius:10px;border:1px solid #EEEBE4;background:#FBFAF7;font:600 13px Montserrat,sans-serif;cursor:pointer;color:#23242A}'
  + '.kpp-m:hover:not(:disabled){border-color:#23242A}.kpp-m.in{background:#FFF1B8;border-color:#FFF1B8}.kpp-m.on{background:#FDCF08;border-color:#FDCF08}.kpp-m.off{color:#c3c3bd;cursor:default}'
  + '.kpp-note{font-size:11.5px;color:#6b6e72;margin:10px 0 12px;line-height:1.4}'
  + '.kpp-cus{display:flex;align-items:center;gap:8px;flex-wrap:wrap;border-top:1px solid #EEEBE4;padding-top:12px;font-size:12.5px}.kpp-cus b{width:100%;font-size:12px;color:#86888D;text-transform:uppercase;letter-spacing:.05em}'
  + '.kpp-cus label{display:flex;align-items:center;gap:6px}.kpp-cus input{border:1px solid #E6E6E1;border-radius:8px;padding:6px 8px;font:500 13px Montserrat,sans-serif;background:#FBFAF7}'
  + '.kpp-ap{margin-left:auto;border:0;border-radius:9px;background:#23242A;color:#fff;padding:8px 14px;font:700 12.5px Montserrat,sans-serif;cursor:pointer}.kpp-err{color:#C7001F;font-size:12px;margin-top:6px;min-height:0}';
var s = document.createElement('style'); s.id = 'kpp-css'; s.textContent = css; (document.head || document.documentElement).appendChild(s);
})(window);
