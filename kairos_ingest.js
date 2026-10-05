/* Кайрос · загрузка выгрузок 1С прямо в браузере.
 * Читает «Продажи по бизнес регионам», «Остатки и доступность по сериям» и (необязательно) отчёт по срокам годности
 * и пересчитывает модель данных так же, как rebuild.py (порт parse_sales / build_krs).
 * Работает и в браузере (File/Blob), и в Node (проверка совпадения с Python).
 * Справочники (прайс, план, CRM, гео) берутся из D.refs / D.plan / D.crm, которые зашиты в дашборд при сборке. */
(function (root) {
'use strict';
var KI = {};

/* ───────────── XLSX: zip + потоковый разбор листа ───────────── */
function u16(b, o) { return b[o] | (b[o + 1] << 8); }
function u32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }

async function readBytes(blob, a, b) { return new Uint8Array(await blob.slice(a, b).arrayBuffer()); }

async function zipIndex(blob) {
  var size = blob.size, tailLen = Math.min(size, 70000);
  var tail = await readBytes(blob, size - tailLen, size), e = -1;
  for (var i = tail.length - 22; i >= 0; i--) if (u32(tail, i) === 0x06054b50) { e = i; break; }
  if (e < 0) throw new Error('Это не файл .xlsx (не найден каталог архива)');
  var cnt = u16(tail, e + 10), cdSize = u32(tail, e + 12), cdOff = u32(tail, e + 16);
  var cd = await readBytes(blob, cdOff, cdOff + cdSize), p = 0, ents = {};
  for (var n = 0; n < cnt; n++) {
    if (u32(cd, p) !== 0x02014b50) break;
    var method = u16(cd, p + 10), csize = u32(cd, p + 20), usize = u32(cd, p + 24);
    var nl = u16(cd, p + 28), el = u16(cd, p + 30), cl = u16(cd, p + 32), lho = u32(cd, p + 42);
    var name = new TextDecoder().decode(cd.subarray(p + 46, p + 46 + nl));
    ents[name] = { method: method, csize: csize, usize: usize, lho: lho };
    p += 46 + nl + el + cl;
  }
  return ents;
}

async function entryStream(blob, ent) {
  var lh = await readBytes(blob, ent.lho, ent.lho + 30);
  var start = ent.lho + 30 + u16(lh, 26) + u16(lh, 28);
  var raw = blob.slice(start, start + ent.csize);
  if (ent.method === 0) return raw.stream();
  if (ent.method !== 8) throw new Error('Неподдерживаемое сжатие в xlsx');
  return raw.stream().pipeThrough(new DecompressionStream('deflate-raw'));
}

async function entryText(blob, ents, name) {
  var ent = ents[name]; if (!ent) return null;
  var st = await entryStream(blob, ent), rd = st.getReader(), dec = new TextDecoder('utf-8'), out = '';
  for (;;) { var r = await rd.read(); if (r.done) break; out += dec.decode(r.value, { stream: true }); }
  return out + dec.decode();
}

function unesc(s) {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, function (m, g) {
    if (g === 'amp') return '&'; if (g === 'lt') return '<'; if (g === 'gt') return '>';
    if (g === 'quot') return '"'; if (g === 'apos') return "'";
    return String.fromCodePoint(g[1] === 'x' ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10));
  });
}

function parseSST(xml) {
  var out = [], parts = xml.split('<si>'); parts.shift();
  for (var i = 0; i < parts.length; i++) {
    var s = parts[i].replace(/<rPh[\s\S]*?<\/rPh>/g, ''), t = '', re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, m;
    while ((m = re.exec(s))) t += m[1];
    out.push(unesc(t));
  }
  return out;
}

function colIdx(letters) { var n = 0; for (var i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64; return n - 1; }

/* Разбирает лист построчно. onRow(r, level, cells[]) — cells индексируются с 0 (A=0). */
async function readSheet(blob, opts, onRow) {
  opts = opts || {};
  var ents = await zipIndex(blob);
  var sstXml = await entryText(blob, ents, 'xl/sharedStrings.xml'), sst = sstXml ? parseSST(sstXml) : [];
  // какой файл листа: по имени, иначе первый
  var path = 'xl/worksheets/sheet1.xml';
  var wb = await entryText(blob, ents, 'xl/workbook.xml'), rels = await entryText(blob, ents, 'xl/_rels/workbook.xml.rels');
  if (wb && rels) {
    var sheets = [], re0 = /<sheet\s[^>]*?name="([^"]*)"[^>]*?r:id="([^"]*)"/g, mm;
    while ((mm = re0.exec(wb))) sheets.push([unesc(mm[1]), mm[2]]);
    var pick = sheets.filter(function (s) { return s[0] === opts.sheet; })[0] || sheets[0];
    if (pick) {
      var rr = new RegExp('<Relationship\\s[^>]*?Id="' + pick[1] + '"[^>]*>').exec(rels), tg = rr && /Target="([^"]*)"/.exec(rr[0]);
      if (tg) path = tg[1].charAt(0) === '/' ? tg[1].slice(1) : 'xl/' + tg[1];
    }
  }
  if (!ents[path]) throw new Error('В файле не найден лист данных');
  var st = await entryStream(blob, ents[path]), rd = st.getReader(), dec = new TextDecoder('utf-8');
  var buf = '', total = ents[path].usize, done = 0, cellRe = /<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, nrows = 0;
  function flush(final) {
    var pos = 0;
    for (;;) {
      var i = buf.indexOf('<row ', pos); if (i < 0) { pos = Math.max(pos, buf.length - 6); break; }
      var j = buf.indexOf('>', i); if (j < 0) { pos = i; break; }
      var attrs, body = '';
      if (buf.charCodeAt(j - 1) === 47) { attrs = buf.slice(i + 5, j - 1); pos = j + 1; }
      else {
        var k = buf.indexOf('</row>', j); if (k < 0) { pos = i; break; }
        attrs = buf.slice(i + 5, j); body = buf.slice(j + 1, k); pos = k + 6;
      }
      var rm = /\br="(\d+)"/.exec(attrs), lm = /outlineLevel="(\d+)"/.exec(attrs), cells = [];
      if (body) {
        cellRe.lastIndex = 0; var c;
        while ((c = cellRe.exec(body))) {
          var inner = c[3]; if (inner === undefined) continue;
          var ta = /\bt="(\w+)"/.exec(c[2]), t = ta ? ta[1] : 'n', val = null, vm;
          if (t === 'inlineStr') { var im = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/.exec(inner); val = im ? unesc(im[1]) : null; }
          else {
            vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
            if (vm) {
              if (t === 's') val = sst[+vm[1]]; else if (t === 'str') val = unesc(vm[1]);
              else if (t === 'b') val = vm[1] === '1'; else if (t === 'e') val = null; else val = parseFloat(vm[1]);
            }
          }
          if (val !== null && val !== undefined) cells[colIdx(c[1])] = val;
        }
      }
      nrows++;
      onRow(+rm[1], lm ? +lm[1] : 0, cells);
    }
    buf = buf.slice(pos);
  }
  for (;;) {
    var r = await rd.read(); if (r.done) break;
    buf += dec.decode(r.value, { stream: true }); done += r.value.length;
    flush(false);
    if (opts.progress) opts.progress(Math.min(0.99, done / (ents[path].csize || 1)));
    if (nrows % 2000 === 0) await new Promise(function (res) { setTimeout(res, 0); });
  }
  buf += dec.decode(); flush(true);
  return nrows;
}
KI.readSheet = readSheet;

/* ───────────── общие функции (порт rebuild.py) ───────────── */
function pyRound(x) { var f = Math.floor(x), d = x - f; if (d < 0.5) return f; if (d > 0.5) return f + 1; return (f % 2 === 0) ? f : f + 1; }
function round1(x) { return Number(x.toFixed(1)); }
function nrm(s) {
  s = String(s == null ? '' : s).toLowerCase().trim().replace(/ё/g, 'е');
  s = s.replace(/[«»"'`(),]/g, ' '); s = s.replace(/х/g, 'x'); s = s.replace(/\s+/g, ' ');
  return s.trim();
}
function cid(n) { return 'c_' + String(n).toLowerCase().replace(/[^a-zа-я0-9]/g, '').slice(0, 20); }
var MN = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
var POSM_KW = ['каталог', 'буклет', 'брошюр', 'плакат', 'воблер', 'листовк', 'ценник', 'наклейк', 'пакет', 'стенд', 'полиграф', 'посм', 'pos-', 'pos ', 'сумка', 'ролл-ап', 'ролап', 'штендер', 'флаер'];
function isPosm(name) { var n = (name || '').toLowerCase(); return POSM_KW.some(function (k) { return n.indexOf(k) >= 0; }); }
var DOC = ['заказ клиента', 'реализация', 'корректировка', 'возврат', 'поступление', 'перемещение', 'списание', 'оприходование', 'инвентаризация', 'отчет комиссионера'];
function isDoc(s) { s = s.toLowerCase(); return DOC.some(function (k) { return s.indexOf(k) === 0; }); }
var VLAD_OKRUGA = { 'Дальневосточный ФО': 1, 'Сибирский ФО (Восток)': 1, 'Сибирский ФО (Запад)': 1, 'Уральский ФО': 1 };
var SNG = { 'Беларусь': 1, 'Казахстан': 1, 'Кыргызстан': 1 };
function geoOf(stack) {
  var r0 = stack[0], channel = r0 === 'СЕТИ' ? 'СЕТИ' : 'ОПТ';
  if (r0 === 'Россия') {
    var ok = stack[1];
    if (ok === 'Россия' || ok == null) return ['Не распределён', 'Не распределён', 'ОПТ'];
    return [ok, VLAD_OKRUGA[ok] ? 'Владивосток' : 'Москва', 'ОПТ'];
  }
  if (SNG[r0]) return [r0, 'Москва', 'ОПТ'];
  if (r0 === 'Москва') return ['Москва', 'Москва', 'ОПТ'];
  if (r0 === 'СЕТИ') return ['СЕТИ', 'Москва', 'СЕТИ'];
  return ['Не распределён', 'Не распределён', 'ОПТ'];
}
function groupCat(cat, brand, name) {
  function infer(t) {
    if (!t) return null;
    if (t.indexOf('пена') >= 0) return 'Монтажные пены';
    if (t.indexOf('герметик') >= 0) return 'Герметики';
    if (t.indexOf('клей') >= 0 || t.indexOf('жидкие гвозд') >= 0) return 'Клей';
    if (t.indexOf('краск') >= 0 || t.indexOf('эмал') >= 0 || t.indexOf('грунт') >= 0 || t.indexOf('аэрозол') >= 0) return 'Аэрозольные краски';
    if (t.indexOf('пистолет') >= 0) return 'Пистолеты и оснастка';
    return null;
  }
  if (brand === 'KRONbuild' || brand === 'ENKI') {
    var g = infer((cat || '').toLowerCase()) || infer((name || '').toLowerCase());
    if (g) return g;
  }
  return cat;
}
function mkPrice(refs) {
  var m = new Map(), artCat = new Map();
  (refs.price || []).forEach(function (r) { m.set(r[0], [r[1], r[2], r[3], r[4]]); });
  m.forEach(function (v) { if (v[0] && v[3]) artCat.set(v[0], v[3]); });
  return { map: m, artCat: artCat, artPrice: refs.artPrice || {} };
}
function brandOf(name, P) {
  var m = P.map.get(nrm(name)); if (m) return m[2];
  var u = name.toLowerCase();
  if (u.indexOf('kronbuild') >= 0 || u.indexOf('кронбилд') >= 0) return 'KRONbuild';
  if (u.indexOf('enki') >= 0 || u.indexOf('энки') >= 0) return 'ENKI';
  return 'HeadRock';
}
function catArt(name, P) { var m = P.map.get(nrm(name)); return m ? [m[3], m[0]] : [null, null]; }
function pad2(n) { return (n < 10 ? '0' : '') + n; }
function dim(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
function dayNo(y, m, d) { return Math.round(Date.UTC(y, m - 1, d) / 86400000); }

/* ───────────── продажи: строки → список строк отчёта ───────────── */
KI.loadSalesRows = async function (blob, progress) {
  var rows = [], hdr = [];
  await readSheet(blob, { sheet: 'Лист_1', progress: progress }, function (r, lv, c) {
    if (r <= 13) hdr.push(c);
    if (r < 8) return;
    var c1 = c[0] == null ? null : c[0], art = c[3], q = c[5], gr = c[6], rt = c[7], nt = c[8];
    if (c1 == null && nt == null && q == null) return;
    var num = function (x) { return typeof x === 'number' ? x : null; };
    rows.push([lv, c1 != null ? String(c1).trim() : null, num(q), num(gr), num(rt), num(nt), art ? String(art).trim() : null]);
  });
  rows.header = hdr;
  return rows;
};

/* ───────────── порт parse_sales ───────────── */
KI.parseSales = function (rows, P, curYear) {
  var n = rows.length, stack = [], orderStack = [];
  var company = { gross: 0, ret: 0, net: 0, qty: 0 };
  var compYear = new Map(), brandYear = new Map(), clientYear = new Map();
  var clientRegion = new Map(), clientFilial = new Map(), clientChannel = new Map(), clientCity = new Map(), clientSubject = new Map();
  var client_net = new Map(), client_qty = new Map(), client_orders = new Map(), client_last = new Map(), client_orderhist = new Map();
  var client_prod = new Map(), comp = new Map(), regSales = new Map();
  var yearMonth = {}, ymGross = {}, ymRet = {}, posm = [0, 0];
  var tx = [], tx_clients = [], tx_cats = [], tx_skus = [], ci_m = new Map(), ct_m = new Map(), sk_m = new Map(), inwork = [], maxdate = 0;
  var rx = /от (\d{2})\.(\d{2})\.(\d{4})/;
  var add = function (m, k, v) { m.set(k, (m.get(k) || 0) + v); };
  function ci(x) { var v = ci_m.get(x); if (v === undefined) { v = tx_clients.length; ci_m.set(x, v); tx_clients.push(x); } return v; }
  function kt(c) { c = c || 'Без категории'; var v = ct_m.get(c); if (v === undefined) { v = tx_cats.length; ct_m.set(c, v); tx_cats.push(c); } return v; }
  function si(name, art) {
    var ca = catArt(name, P), cat = ca[0]; art = art || ca[1];
    if (art && P.artCat.get(art)) cat = P.artCat.get(art);
    var b = brandOf(name, P); cat = groupCat(cat, b, name);
    var key = art || ('n:' + nrm(name)), v = sk_m.get(key);
    if (v === undefined) { v = tx_skus.length; sk_m.set(key, v); tx_skus.push([art || '—', name, kt(cat), b]); }
    return v;
  }
  for (var i = 0; i < n; i++) {
    var row = rows[i], lvl = row[0], name = row[1], q = row[2], gr = row[3], rt = row[4], nt = row[5], art = row[6];
    if (name == null) { stack[lvl] = null; stack.length = lvl + 1; continue; }
    stack[lvl] = name; stack.length = lvl + 1; orderStack.length = Math.min(orderStack.length, lvl);
    if (lvl === 0) continue;
    var low = name.toLowerCase(), parent = stack[lvl - 1];
    if (isDoc(name)) {
      if (!isDoc(parent || '')) {
        var comp_name = parent || stack[1] || 'Прочее', cl = comp_name.toLowerCase();
        if (cl.indexOf('интернет решения') >= 0 || cl.indexOf('(озон)') >= 0 || cl.indexOf('ozon') >= 0) comp_name = 'OZON (маркетплейс)';
        else if (cl.indexOf('вайлдберриз') >= 0 || cl.indexOf('wildberries') >= 0) comp_name = 'Wildberries (маркетплейс)';
        var g = geoOf(stack), reg = g[0], fil = g[1], chan = g[2];
        if (comp_name.toLowerCase().indexOf('нордлогистик') >= 0) { reg = 'Сибирский ФО (Запад)'; fil = 'Владивосток'; }
        if (comp_name.trim().toLowerCase().indexOf('частное лицо') === 0) comp_name = comp_name + ' · ' + reg;
        var m = rx.exec(name);
        var di = m ? parseInt(m[3] + m[2] + m[1], 10) : 0, iso = m ? m[3] + '-' + m[2] + '-' + m[1] : '';
        orderStack[lvl] = [di, iso, name.split('от')[0].trim(), nt || 0, comp_name, reg, fil, chan];
        if (!clientRegion.has(comp_name)) clientRegion.set(comp_name, reg);
        if (!clientFilial.has(comp_name)) clientFilial.set(comp_name, fil);
        if (!clientChannel.has(comp_name)) clientChannel.set(comp_name, chan);
        var CL = lvl - 1, subj = '', city = '';
        if (stack[0] === 'Россия') { if (CL >= 3) subj = stack[2] || ''; if (CL >= 4) city = stack[3] || ''; }
        if (!clientSubject.has(comp_name)) clientSubject.set(comp_name, subj);
        if (!clientCity.has(comp_name)) clientCity.set(comp_name, city);
        if (di > maxdate) maxdate = di;
        if (iso && iso > (client_last.get(comp_name) || '')) client_last.set(comp_name, iso);
        if (low.indexOf('заказ клиента') === 0) {
          add(client_orders, comp_name, 1);
          if (m) {
            var hl = client_orderhist.get(comp_name); if (!hl) { hl = []; client_orderhist.set(comp_name, hl); }
            hl.push({ date: m[1] + '.' + m[2] + '.' + m[3], num: name.split('от')[0].replace('Заказ клиента', '').trim(), sum: pyRound(nt || 0), qty: Math.trunc(q || 0) });
            if (Math.floor(di / 10000) === curYear) {
              var shipped = 0, j = i + 1;
              while (j < n && rows[j][0] > lvl && rows[j][1] != null) { if (rows[j][1].toLowerCase().indexOf('реализация') === 0) shipped += rows[j][5] || 0; j++; }
              var nsh = (nt || 0) - shipped;
              if (nsh > 1) inwork.push({ client: comp_name, order: name.split('от')[0].replace('Заказ клиента', '').trim(), date: m[1] + '.' + m[2] + '.' + m[3],
                ordered: pyRound(nt || 0), shipped: pyRound(shipped), not_shipped: pyRound(nsh) });
            }
          }
        }
      }
      continue;
    }
    var nxt = i + 1 < n ? rows[i + 1][0] : -1;
    if (nxt > lvl) continue;
    if (nt == null) continue;
    var co = null; for (var k = Math.min(lvl - 1, orderStack.length - 1); k >= 0; k--) if (orderStack[k]) { co = orderStack[k]; break; }
    if (!co) continue;
    if (isPosm(name)) { posm[0]++; posm[1] += nt || 0; continue; }
    var d2 = co[0], client = co[4], reg2 = co[5];
    q = q || 0; var net = nt, b = brandOf(name, P);
    company.net += net; company.gross += (gr || 0); company.ret += (rt || 0); company.qty += q;
    add(regSales, reg2, net); add(client_net, client, net); add(client_qty, client, q);
    var cp = client_prod.get(client); if (!cp) { cp = new Map(); client_prod.set(client, cp); }
    var e1 = cp.get(name); if (!e1) { e1 = [0, 0]; cp.set(name, e1); } e1[0] += q; e1[1] += net;
    var e2 = comp.get(name); if (!e2) { e2 = [0, 0]; comp.set(name, e2); } e2[0] += q; e2[1] += net;
    if (d2 > 0) {
      var yr = Math.floor(d2 / 10000), ym = String(d2).slice(0, 6);
      yearMonth[ym] = (yearMonth[ym] || 0) + net; ymGross[ym] = (ymGross[ym] || 0) + (gr || 0); ymRet[ym] = (ymRet[ym] || 0) + (rt || 0);
      var cy = compYear.get(yr); if (!cy) { cy = { gross: 0, ret: 0, net: 0, qty: 0 }; compYear.set(yr, cy); }
      cy.net += net; cy.gross += (gr || 0); cy.ret += (rt || 0); cy.qty += q;
      var by = brandYear.get(yr); if (!by) { by = new Map(); brandYear.set(yr, by); } add(by, b, net);
      var cly = clientYear.get(client); if (!cly) { cly = new Map(); clientYear.set(client, cly); } add(cly, yr, net);
      if (net !== 0) tx.push([d2, ci(client), si(name, art), round1(q), pyRound(net)]);
    }
  }
  inwork.sort(function (a, b) { return b.not_shipped - a.not_shipped; });
  return { company: company, comp: comp, compYear: compYear, brandYear: brandYear, clientYear: clientYear, clientRegion: clientRegion, clientFilial: clientFilial,
    clientChannel: clientChannel, clientCity: clientCity, clientSubject: clientSubject, client_net: client_net, client_qty: client_qty, client_orders: client_orders,
    client_last: client_last, client_orderhist: client_orderhist, client_prod: client_prod, regSales: regSales, yearMonth: yearMonth, ymGross: ymGross, ymRet: ymRet,
    tx: tx, tx_clients: tx_clients, tx_cats: tx_cats, tx_skus: tx_skus, inwork: inwork, maxdate: maxdate, posm: posm };
};

/* ───────────── остатки и сроки годности ───────────── */
KI.loadStock = async function (blob, refs, progress) {
  var WH = refs.stockWh || 'Адресный Лакония ОПТ', agg = new Map(), cur = null, seenHdr = false;
  function whBucket(nm) {
    var n = (nm || '').toLowerCase();
    if (n.indexOf('лакония опт') >= 0) return 'Лакония ОПТ';
    if (n.indexOf('лакония сети') >= 0) return 'Лакония СЕТИ';
    if (n.indexOf('владивосток') >= 0 || n.indexOf('янковск') >= 0 || n.indexOf('мангут') >= 0) return 'Владивосток';
    if (n.indexOf('москва') >= 0 || n.indexOf('лист ложистик') >= 0 || n.indexOf('гидд') >= 0) return 'Москва';
    if (n.indexOf('новосибирск') >= 0) return 'Новосибирск';
    if (n.indexOf('лакония') >= 0) return 'Лакония прочее';
    return 'Прочее';
  }
  var num = function (x) { return typeof x === 'number' ? x : 0; };
  await readSheet(blob, { sheet: 'Лист_1', progress: progress }, function (r, lv, v) {
    if (v[0] === 'Артикул' && v[2] === 'Номенклатура') seenHdr = true;
    var c0 = v[0], c2 = v[2], c5 = v[5], isart = !!(c0 && c2 && c5);
    if (c0 && !isart) { if (lv === 0) cur = String(c0).trim(); return; }
    if (!isart) return;
    var a = String(c0).trim(), d = agg.get(a);
    if (!d) { d = { name: '', in_stock: 0, shipping: 0, reserved: 0, available: 0, incoming: 0, company: 0, wh: {} }; agg.set(a, d); }
    if (!d.name) d.name = String(c2).trim();
    var q7 = num(v[7]); d.company += q7;
    if (q7) { var wb = whBucket(cur); d.wh[wb] = (d.wh[wb] || 0) + q7; }
    if (cur === WH) { d.in_stock += q7; d.shipping += num(v[8]); d.reserved += num(v[9]); d.available += num(v[10]); d.incoming += num(v[11]); }
  });
  if (!seenHdr) throw new Error('Не похоже на отчёт «Остатки и доступность по сериям» (нет шапки «Артикул / Номенклатура»)');
  var stock = {};
  agg.forEach(function (d, a) {
    var wh = {}; Object.keys(d.wh).forEach(function (k) { var x = round1(d.wh[k]); if (Math.abs(d.wh[k]) >= 0.5) wh[k] = x; });
    stock[a] = { name: d.name, in_stock: round1(d.in_stock), shipping: round1(d.shipping), reserved: round1(d.reserved), available: round1(d.available),
      incoming: round1(d.incoming), company_in_stock: round1(d.company), wh: wh };
  });
  return stock;
};

KI.loadExpiry = async function (blob, progress) {
  var batches = {}, ok = false;
  function pdate(v) {
    if (typeof v === 'number' && v > 20000) { var dt = new Date(Math.round((v - 25569) * 86400000)); return dt.getUTCFullYear() + '-' + pad2(dt.getUTCMonth() + 1) + '-' + pad2(dt.getUTCDate()); }
    if (typeof v === 'string') { var m = /^(\d{2})\.(\d{2})\.(\d{4})/.exec(v.trim()); if (m) return m[3] + '-' + m[2] + '-' + m[1]; }
    return null;
  }
  await readSheet(blob, { sheet: 'Лист_1', progress: progress }, function (r, lv, v) {
    if (v[9] === 'Дата окончания срока годности') ok = true;
    if (r < 11) return;
    var art = v[0]; if (!art) return;
    var a = String(art).trim(); if (a.toLowerCase().indexOf('лакония') >= 0 || a.toLowerCase().indexOf('склад') >= 0) return;
    var d = pdate(v[9]); if (!d) return;
    var q = typeof v[11] === 'number' ? v[11] : 0;
    if (q > 0) (batches[a] = batches[a] || []).push([d, q]);
  });
  if (!ok) throw new Error('Не похоже на отчёт «Товары на складах с окончанием срока годности» (нет колонки «Дата окончания срока годности»)');
  Object.keys(batches).forEach(function (a) { batches[a].sort(function (x, y) { return x[0] < y[0] ? -1 : x[0] > y[0] ? 1 : x[1] - y[1]; }); });
  return batches;
};

function expiryOut(raw, refiso) {
  var ry = +refiso.slice(0, 4), rm = +refiso.slice(5, 7), rd = +refiso.slice(8), refN = dayNo(ry, rm, rd), out = {};
  Object.keys(raw).forEach(function (a) {
    var withq = raw[a], exp = 0, soon = 0, lst = [];
    withq.forEach(function (b) {
      var p = b[0].split('-'), dN = dayNo(+p[0], +p[1], +p[2]), days = dN - refN, q = b[1];
      if (dN < refN) exp += q; else if (days <= 90) soon += q;
      lst.push({ d: p[2] + '.' + p[1] + '.' + p[0], q: pyRound(q), st: dN < refN ? 'exp' : (days <= 90 ? 'soon' : 'ok') });
    });
    var near = withq[0], nd = null, nl = null;
    if (near) { var pp = near[0].split('-'); nl = pp[2] + '.' + pp[1] + '.' + pp[0]; nd = dayNo(+pp[0], +pp[1], +pp[2]) - refN; }
    out[a] = { nearest: nl, nearest_days: nd, expired: pyRound(exp), soon: pyRound(soon), list: lst };
  });
  return out;
}

/* ───────────── порт build_krs ───────────── */
function abcOf(comp) {
  var items = []; comp.forEach(function (d, n) { if (d[1] > 0) items.push([n, d[1]]); });
  items.sort(function (a, b) { return b[1] - a[1]; });
  var tot = 0; items.forEach(function (x) { tot += x[1]; }); tot = tot || 1;
  var cum = 0, abc = new Map();
  items.forEach(function (x) { cum += x[1]; var sh = cum / tot; abc.set(x[0], sh <= 0.8 ? 'A' : (sh <= 0.95 ? 'B' : 'C')); });
  comp.forEach(function (d, n) { if (!abc.has(n)) abc.set(n, 'C'); });
  return [abc, tot];
}
function grow(a, b) { return b ? pyRound((a - b) / b * 1000) / 10 : null; }
function sortBy(arr, f) { return arr.map(function (x, i) { return [f(x), i, x]; }).sort(function (a, b) { return (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] - b[1]); }).map(function (z) { return z[2]; }); }
function obj(map) { var o = {}; map.forEach(function (v, k) { o[k] = v; }); return o; }

KI.buildKRS = function (S, stock, expiryRaw, ctx) {
  var D = ctx.D, P = ctx.P, plan = D.plan || {}, refs = D.refs || {}, TURN = refs.turnWindow || 90;
  var cy = S.compYear, by = S.brandYear, comp = S.comp;
  var years = []; cy.forEach(function (v, y) { years.push(y); });
  var maxd = String(S.maxdate), CUR = maxd.length === 8 ? +maxd.slice(0, 4) : (years.length ? Math.max.apply(null, years) : 2026);
  if (!cy.has(CUR)) CUR = years.length ? Math.max.apply(null, years) : CUR;
  var PREV = CUR - 1;
  var refiso = maxd.length === 8 ? maxd.slice(0, 4) + '-' + maxd.slice(4, 6) + '-' + maxd.slice(6) : CUR + '-12-31';
  var refN = dayNo(+refiso.slice(0, 4), +refiso.slice(5, 7), +refiso.slice(8));
  var expiry = expiryOut(expiryRaw || {}, refiso);
  var excluded = function (nm) { return (S.client_net.get(nm) || 0) <= 0; };
  var nExcl = 0; S.client_net.forEach(function (v, nm) { if (excluded(nm)) nExcl++; });
  var abcR = abcOf(comp), abc = abcR[0], totRev = abcR[1];
  /* CRM */
  var CRM_OKRUG = { 'УФО': ['Уральский ФО', 'Владивосток'], 'ДВФО': ['Дальневосточный ФО', 'Владивосток'], 'Западная Сибирь': ['Сибирский ФО (Запад)', 'Владивосток'],
    'Восточная Сибирь': ['Сибирский ФО (Восток)', 'Владивосток'], 'ЦФО': ['Центральный ФО', 'Москва'], 'ЮФО': ['Южный ФО', 'Москва'], 'СЗФО': ['Северо-Западный ФО', 'Москва'],
    'СКФО': ['Северо-Кавказский ФО', 'Москва'], 'ПФО': ['Поволжский ФО', 'Москва'], 'Москва и МО': ['Москва', 'Москва'], 'СНГ': ['СНГ', 'Москва'] };
  var oldCrm = D.crm || {}, cinn = new Map(); Object.keys(D.clientInn || {}).forEach(function (n) { cinn.set(nrm(n), D.clientInn[n]); });
  var crmByInn = {}; Object.keys(oldCrm).forEach(function (k) { var i = String((oldCrm[k] || {}).inn || '').trim(); if (i && !(i in crmByInn)) crmByInn[i] = oldCrm[k]; });
  var crmOf = function (nm) { var i = cinn.get(nrm(nm)); if (i && crmByInn[i]) return crmByInn[i]; return oldCrm[cid(nm)] || {}; };
  var cReg = new Map(S.clientRegion), cFil = new Map(S.clientFilial), cChan = new Map(S.clientChannel);
  cReg.forEach(function (v, nm) { if (v === 'Не распределён') { var ok = (crmOf(nm) || {}).okrug || ''; if (CRM_OKRUG[ok]) { cReg.set(nm, CRM_OKRUG[ok][0]); cFil.set(nm, CRM_OKRUG[ok][1]); } } });
  var SNG_REG = { 'Беларусь': 1, 'Казахстан': 1, 'Кыргызстан': 1, 'СНГ': 1 };
  var bucket = function (nm) {
    if (cChan.get(nm) === 'СЕТИ') return 'Сети';
    if (SNG_REG[cReg.get(nm)]) return 'СНГ';
    var f = cFil.get(nm);
    if (f === 'Владивосток') return 'ОПТ Владивосток'; if (f === 'Москва') return 'ОПТ Москва'; return 'Не распределён';
  };
  var OKRUG_MGR = { 'Центральный ФО': 'Сергей Сидоров', 'Северо-Западный ФО': 'Сергей Сидоров', 'Поволжский ФО': 'Семён Комиссаров', 'Южный ФО': 'Семён Комиссаров',
    'Дальневосточный ФО': 'Александр Федотов', 'Сибирский ФО (Восток)': 'Евгений Швайгерт', 'Сибирский ФО (Запад)': 'Дмитрий Жигалов', 'Уральский ФО': 'Михаил Гнипель' };
  var cMgr = {};
  cReg.forEach(function (reg, nm) {
    var ch = cChan.get(nm);
    if (ch === 'СЕТИ') { cMgr[nm] = null; return; }
    if (reg === 'Москва') { var resp = String((crmOf(nm) || {}).responsible || '').trim(); cMgr[nm] = (resp === 'Максим Чашников' || resp === 'Василий Димитрюк') ? resp : 'Максим Чашников'; }
    else if (SNG_REG[reg]) cMgr[nm] = 'Сергей Сидоров';
    else if (OKRUG_MGR[reg]) cMgr[nm] = OKRUG_MGR[reg];
    else cMgr[nm] = 'Не назначен';
  });
  var crmJoined = {}; S.client_net.forEach(function (v, nm) { var rec = crmOf(nm); if (rec && Object.keys(rec).length) crmJoined[cid(nm)] = rec; });
  /* каталог + ABC */
  var catItems = []; comp.forEach(function (d, n) { if (d[1] > 0) catItems.push([n, d]); });
  catItems = sortBy(catItems, function (x) { return -x[1][1]; });
  var catalog = [], name2cat = new Map();
  catItems.forEach(function (x) { var ca = catArt(x[0], P); name2cat.set(x[0], catalog.length); catalog.push([ca[1] || '—', x[0], ca[0] || 'Без категории', abc.get(x[0]) || 'C', pyRound(x[1][1])]); });
  var A = [], B = [], C = []; abc.forEach(function (v, n) { (v === 'A' ? A : v === 'B' ? B : C).push(n); });
  var sumRev = function (L) { var s = 0; L.forEach(function (n) { s += comp.get(n)[1]; }); return s; };
  var rev_by = { A: sumRev(A), B: sumRev(B), C: sumRev(C) };
  var buyers = new Map(); S.client_prod.forEach(function (prods) { prods.forEach(function (v, n) { if (v[1] !== 0) buyers.set(n, (buyers.get(n) || 0) + 1); }); });
  var nclients = S.client_prod.size;
  var aSorted = sortBy(A, function (n) { return -comp.get(n)[1]; });
  var a_list = aSorted.slice(0, 20).map(function (n) { return { article: catArt(n, P)[1] || '—', name: n, rev: pyRound(comp.get(n)[1]) }; });
  var reserve = sortBy(A, function (n) { return -((nclients - (buyers.get(n) || 0)) * comp.get(n)[1]); }).slice(0, 15)
    .map(function (n) { return { name: n, nonbuyers: nclients - (buyers.get(n) || 0), rev: pyRound(comp.get(n)[1]) }; });
  var catTotal = {}; comp.forEach(function (d, n) { var c = catArt(n, P)[0]; if (c) catTotal[c] = (catTotal[c] || 0) + 1; });
  /* оборачиваемость */
  var cutoff = refN - TURN, recentQty = new Map();
  S.tx.forEach(function (t) {
    var s = String(t[0]); var o = dayNo(+s.slice(0, 4), +s.slice(4, 6), +s.slice(6));
    if (o >= cutoff) { var a = S.tx_skus[t[2]][0]; if (a && a !== '—') recentQty.set(a, (recentQty.get(a) || 0) + t[3]); }
  });
  var artPrice = P.artPrice, turn = {};
  Object.keys(stock).forEach(function (art) {
    var s = stock[art], inst = s.in_stock || 0; if (inst <= 0) return;
    var sold = recentQty.get(art) || 0, perday = sold / TURN, dos = perday > 0 ? inst / perday : null, p1 = artPrice[art] || 0;
    var ex = expiry[art] || {}, expQ = ex.expired || 0, soonQ = ex.soon || 0;
    turn[art] = { dos: dos !== null ? pyRound(dos) : null, in_stock: inst, available: s.available, reserved: s.reserved || 0, incoming: s.incoming || 0, wh: s.wh || {},
      company_in_stock: s.company_in_stock || 0, sold: pyRound(sold), frozen: pyRound(inst * p1), name: s.name,
      exp_near: ex.nearest === undefined ? null : ex.nearest, exp_days: ex.nearest_days === undefined ? null : ex.nearest_days, exp_expired: expQ, exp_soon: soonQ,
      exp_froz: pyRound(expQ * p1), soon_froz: pyRound(soonQ * p1), batches: ex.list || [],
      status: dos === null ? 'dead' : (dos < 30 ? 'short' : (dos > 180 ? 'slow' : 'ok')) };
  });
  /* overview */
  var o_cur = cy.get(CUR) || { net: 0, gross: 0, ret: 0, qty: 0 }, o_prev = cy.get(PREV) || { net: 0 };
  var plancur = plan[CUR] || {}, ym = S.yearMonth, by_cur = by.get(CUR) || new Map();
  var mfact = function (y, mi) { return pyRound((ym[y + pad2(mi)] || 0) / 1e6 * 1000) / 1000; };
  var pm = (plancur.months || {});
  var months = []; for (var mi = 1; mi <= 12; mi++) months.push({ m: MN[mi - 1], fact: mfact(CUR, mi), prev: mfact(PREV, mi), plan: plancur.months ? pyRound((pm[mi] || 0) / 1e6 * 1000) / 1000 : null });
  var clientsCur = new Map();
  S.client_net.forEach(function (v, cl) { if (excluded(cl)) return; var x = (S.clientYear.get(cl) || new Map()).get(CUR) || 0; if (x) clientsCur.set(cl, x); });
  var ordersCur = 0; S.client_orderhist.forEach(function (h) { h.forEach(function (x) { if (x.date.slice(-4) === String(CUR)) ordersCur++; }); });
  var curMonth = maxd.length === 8 ? +maxd.slice(4, 6) : 12, cutD = maxd.length === 8 ? +maxd.slice(6, 8) : 31, dm = dim(CUR, curMonth), partial = cutD < dm;
  var prevCut = PREV * 10000 + curMonth * 100 + (partial ? cutD : 31), prevSame = 0;
  S.tx.forEach(function (t) { if (t[0] >= PREV * 10000 + 101 && t[0] <= prevCut) prevSame += t[4]; });
  var planYtd = null;
  if (Object.keys(plancur).length) { planYtd = 0; for (var mj = 1; mj <= curMonth; mj++) planYtd += (pm[mj] || 0) * ((mj === curMonth && partial) ? cutD / dm : 1); }
  var oc = 0, op = 0, uc = 0, up = 0, cc = new Set(), cp = new Set();
  S.client_orderhist.forEach(function (hist, cl) {
    if (excluded(cl)) return;
    hist.forEach(function (h) {
      var p = (h.date || '').split('.'), y = +p[2] || 0, mo = +p[1] || 0;
      if (mo < 1 || mo > curMonth) return;
      if (y === PREV && mo === curMonth && partial && (+h.date.slice(0, 2) || 0) > cutD) return;
      if (y === CUR) { oc++; uc += h.qty || 0; cc.add(cl); } else if (y === PREV) { op++; up += h.qty || 0; cp.add(cl); }
    });
  });
  var yoy = { sales: grow(o_cur.net, prevSame), clients: grow(cc.size, cp.size), orders: grow(oc, op), units: grow(uc, up) };
  var catnet = new Map(), prodnet = new Map(), regnet = new Map(), cr = S.clientRegion;
  S.tx.forEach(function (t) {
    if (Math.floor(t[0] / 10000) !== CUR || t[4] <= 0) return;
    var sk = S.tx_skus[t[2]], cn = S.tx_cats[sk[2]];
    if (cn && cn !== 'Без категории') catnet.set(cn, (catnet.get(cn) || 0) + t[4]);
    prodnet.set(sk[1], (prodnet.get(sk[1]) || 0) + t[4]);
    var rg = cr.get(S.tx_clients[t[1]]); if (rg && rg !== 'Не распределён') regnet.set(rg, (regnet.get(rg) || 0) + t[4]);
  });
  var arr = function (m) { var a = []; m.forEach(function (v, k) { a.push([k, v]); }); return sortBy(a, function (x) { return -x[1]; }); };
  var catsSorted = arr(catnet).slice(0, 8), ctot = 0; catsSorted.forEach(function (x) { ctot += x[1]; }); ctot = ctot || 1;
  var ovCats = catsSorted.map(function (x) { return { name: x[0], sales: pyRound(x[1]), share: pyRound(x[1] / ctot * 1000) / 10 }; });
  var ovTop = arr(prodnet).slice(0, 10).map(function (x) { var ar = catArt(x[0], P)[1] || ''; return { sku: ar || '—', name: x[0], abc: abc.get(x[0]) || 'C', sales: pyRound(x[1]), stock: (stock[ar] || {}).in_stock || 0 }; });
  var ovRegs = arr(regnet).map(function (x) { return { name: x[0], value: pyRound(x[1]) }; });
  var growth = grow(o_cur.net, prevSame), avgOrder = ordersCur ? pyRound(o_cur.net / ordersCur) : 0;
  var byv = function (b) { return pyRound(by_cur.get(b) || 0); };
  var overview = { sales: pyRound(o_cur.net), grossSales: pyRound(o_cur.gross), returns: pyRound(o_cur.ret), units: Math.trunc(o_cur.qty),
    returnsRate: o_cur.gross ? pyRound(-o_cur.ret / o_cur.gross * 10000) / 100 : 0, plan: plancur.annual === undefined ? null : plancur.annual,
    planYTD: planYtd ? pyRound(planYtd) : null, planDone: planYtd ? pyRound(o_cur.net / planYtd * 1000) / 10 : null, growth: growth,
    clients: clientsCur.size, orders: ordersCur, avgOrder: avgOrder, yoy: yoy, months: months,
    branches: [{ id: 'all', name: 'Вся компания', sales: pyRound(o_cur.net), plan: plancur.annual === undefined ? null : plancur.annual, growth: grow(o_cur.net, o_prev.net || 0),
      headrock: byv('HeadRock'), kron: byv('KRONbuild'), enki: byv('ENKI') }],
    topClients: arr(clientsCur).slice(0, 10).map(function (x) { return { name: x[0], sales: pyRound(x[1]), share: pyRound(x[1] / (o_cur.net || 1) * 1000) / 10 }; }),
    topProducts: ovTop, regions: ovRegs, categories: ovCats };
  var catNames = Object.keys(catTotal).sort(function (a, b) { return catTotal[b] - catTotal[a]; });
  var sdBr = function (nm) { return { name: nm, sales: byv(nm), clients: clientsCur.size, orders: ordersCur, avgOrder: overview.avgOrder, brands: [] }; };
  var salesDetail = { branches: { moscow: { name: 'Вся компания', sales: pyRound(o_cur.net), clients: clientsCur.size, orders: ordersCur, avgOrder: overview.avgOrder,
      brands: [{ id: 'headrock', name: 'HeadRock', fact: byv('HeadRock') }, { id: 'kron', name: 'KRONbuild', fact: byv('KRONbuild') }, { id: 'enki', name: 'ENKI', fact: byv('ENKI') }] },
      vlad: { name: 'Владивосток', sales: null, missing: true } },
    brands: { headrock: sdBr('HeadRock'), kron: sdBr('KRONbuild'), enki: sdBr('ENKI') },
    categories: catNames.map(function (c) { var s = 0; comp.forEach(function (d, n) { if (catArt(n, P)[0] === c) s += d[1]; }); return { name: c, sales: pyRound(s), growth: null, sku: catTotal[c] }; }),
    products: [] };
  /* клиенты */
  var clients_list = [], clientDetail = {}, clientAssort = {}, clientRegion = obj(cReg);
  var cl_sorted = []; S.client_net.forEach(function (v, nm) { cl_sorted.push(nm); });
  cl_sorted = sortBy(cl_sorted, function (c) { return -((S.clientYear.get(c) || new Map()).get(CUR) || 0); });
  var refDate = new Date(Date.UTC(+refiso.slice(0, 4), +refiso.slice(5, 7) - 1, +refiso.slice(8)));
  cl_sorted.forEach(function (name) {
    if (excluded(name)) return;
    var id = cid(name), prods = S.client_prod.get(name) || new Map(), bought = [];
    prods.forEach(function (v, n) { if (v[1] !== 0) bought.push(n); });
    var cyr = S.clientYear.get(name) || new Map(), rev_cur = cyr.get(CUR) || 0, hist = S.client_orderhist.get(name) || [];
    var hist_cur = hist.filter(function (h) { return h.date.slice(-4) === String(CUR); }), ordn = hist_cur.length, last = S.client_last.get(name) || '';
    var qty = 0; prods.forEach(function (v) { qty += v[0]; }); qty = pyRound(qty);
    var catpen = new Map(), btot = 0;
    bought.forEach(function (n) { var c = catArt(n, P)[0]; btot += prods.get(n)[1]; if (c) { var e = catpen.get(c); if (!e) { e = [0, 0]; catpen.set(c, e); } e[0]++; e[1] += prods.get(n)[1]; } });
    var cats = []; var cpa = []; catpen.forEach(function (v, c) { cpa.push([c, v]); }); cpa = sortBy(cpa, function (x) { return -x[1][1]; });
    cpa.forEach(function (x) {
      var c = x[0], got = x[1][0], r = x[1][1], totc = catTotal[c] || got, pen = totc ? pyRound(got / totc * 100) : 0;
      cats.push({ name: c, total: totc, buy: got, share: pyRound(r / (btot || 1) * 1000) / 10, status: pen >= 50 ? 'Хорошо' : (pen >= 25 ? 'Среднее' : (pen > 0 ? 'Зона роста' : 'Не представлена')) });
    });
    var bset = new Set(bought);
    var cov = {}; [['A', A], ['B', B], ['C', C]].forEach(function (kv) { var buy = 0; kv[1].forEach(function (n) { if (bset.has(n)) buy++; }); cov[kv[0]] = { total: kv[1].length, buy: buy }; });
    var gapsAll = A.filter(function (n) { return !bset.has(n); }), gaps = sortBy(gapsAll, function (n) { return -comp.get(n)[1]; }).slice(0, 12);
    var lastD = null; if (last) { var lp = last.split('-').map(Number); lastD = Math.round((refN - dayNo(lp[0], lp[1], lp[2]))); }
    var st = (lastD !== null && lastD <= 45) ? 'Активный' : ((lastD || 999) > 90 ? 'Риск' : 'Снижение'); if (rev_cur <= 0) st = 'Риск';
    var reg = clientRegion[name] || '—', gr = grow(rev_cur, cyr.get(PREV) || 0), lastTxt = last ? last.replace(/-/g, '.') : '—';
    clients_list.push({ id: id, name: name, manager: reg, region: reg, sales: pyRound(rev_cur), plan: null, growth: gr, brands: 1, categories: cats.length, sku: bought.length, last: lastTxt,
      potential: gaps.length >= 100 ? 'Высокий' : (gaps.length >= 40 ? 'Средний' : 'Низкий'), status: st, avgCheck: ordn ? pyRound(rev_cur / ordn) : 0, freq: ordn,
      aGap: cov.A.total - cov.A.buy, days: lastD });
    clientDetail[id] = { name: name, manager: reg, region: reg, city: '—', segment: '—', status: st, sales: pyRound(rev_cur), units: qty, orders: ordn, avgOrder: ordn ? pyRound(rev_cur / ordn) : 0,
      last: lastTxt, growth: gr, discount: null, brands: [{ name: 'HeadRock', sales: pyRound(rev_cur), share: 100 }], months: [], categories: [], abc: cov, recommended: [],
      ordersHistory: hist_cur.slice(-12).map(function (h) { return { date: h.date, order: h.num, brand: '—', sum: h.sum, sku: h.qty, discount: null, status: 'Из отчёта' }; }), buySku: bought.length };
    var as = {}; bought.forEach(function (n) { if (name2cat.has(n)) as[String(name2cat.get(n))] = pyRound(prods.get(n)[1]); }); clientAssort[id] = as;
  });
  /* регионы / филиалы */
  var regionsData = {}, rc = new Map();
  S.client_net.forEach(function (net, nm) { if (excluded(nm)) return; var r = cReg.get(nm) || 'Не распределён', m = rc.get(r); if (!m) { m = new Map(); rc.set(r, m); } m.set(nm, net); });
  rc.forEach(function (cls, reg) {
    var lst = [], tot = 0; cls.forEach(function (v, n) { tot += v; if (v > 0) lst.push({ name: n, sales: pyRound(v) }); });
    regionsData[reg] = { sales: pyRound(tot), clients: lst.length, list: sortBy(lst, function (x) { return -x.sales; }) };
  });
  var factB = {}; S.tx.forEach(function (t) { var y = Math.floor(t[0] / 10000), b = bucket(S.tx_clients[t[1]]); (factB[y] = factB[y] || {})[b] = ((factB[y] || {})[b] || 0) + t[4]; });
  var salesFilial = {}, salesFilial3 = {};
  [CUR, PREV, CUR - 2].filter(function (v, i, a) { return a.indexOf(v) === i; }).sort().forEach(function (y) {
    var pl = ((plan[y] || {}).bucketAnnual) || {}, f = function (bk) { return pyRound((factB[y] || {})[bk] || 0); };
    var pp = function (bk) { return bk in pl ? pyRound(pl[bk]) : null; };
    salesFilial[y] = {}; ['ОПТ Москва', 'ОПТ Владивосток', 'СНГ', 'Сети'].forEach(function (bk) { salesFilial[y][bk] = { fact: f(bk), plan: pp(bk) }; });
    salesFilial3[y] = { 'Москва': { fact: f('ОПТ Москва') + f('СНГ'), plan: pp('ОПТ Москва') }, 'Владивосток': { fact: f('ОПТ Владивосток'), plan: pp('ОПТ Владивосток') }, 'Сети': { fact: f('Сети'), plan: pp('Сети') } };
  });
  var foreign = []; S.regSales.forEach(function (v, r) { if (SNG[r]) foreign.push(r); });
  var regionsFull = arr(S.regSales).map(function (x) { return { name: x[0], sales: pyRound(x[1]) }; });
  /* остатки */
  var stockOut = {}, low = [];
  Object.keys(stock).forEach(function (a) {
    var s = stock[a];
    stockOut[a] = { name: s.name, in_stock: s.in_stock, shipping: s.shipping, reserved_client: s.reserved, reserved_sale: 0, available: s.available, incoming: s.incoming, total_available: s.available };
    if (s.available <= 0) low.push({ article: a, name: s.name, available: s.available, incoming: s.incoming, in_stock: s.in_stock });
  });
  low = sortBy(low, function (x) { return x.available; }).slice(0, 40);
  var clientInn = {}; S.client_net.forEach(function (v, cn) { if (cinn.has(nrm(cn))) clientInn[cn] = cinn.get(nrm(cn)); });
  var ymd = refiso.slice(8) + '.' + refiso.slice(5, 7) + '.' + refiso.slice(0, 4);
  var revR = {}, revS = {}; ['A', 'B', 'C'].forEach(function (k) { revR[k] = pyRound(rev_by[k]); revS[k] = pyRound(rev_by[k] / (totRev || 1) * 1000) / 10; });
  var K = {};
  Object.keys(D).forEach(function (k) { K[k] = D[k]; });                                  // справочники: гео, карта, план, население, Контур…
  var upd = {
    meta: { prototype: true, note: 'Пересчитано в браузере из загруженных файлов 1С. Вся компания, ' + CUR + ' (АППГ ' + PREV + ').',
      sourceCoverage: { period: CUR + ' (АППГ ' + PREV + ')', scope: 'Вся компания / все бренды', actualNetSales: pyRound(o_cur.net), actualUnits: Math.trunc(o_cur.qty), actualReturns: pyRound(o_cur.ret),
        missing: ['маржинальность', 'себестоимость', 'скидки'] } },
    periods: [{ id: 'ytd', label: CUR + ' год', available: true }], overview: overview,
    abcSummary: { totalRevenue: pyRound(totRev), counts: { A: A.length, B: B.length, C: C.length }, totalSku: catalog.length },
    salesDetail: salesDetail, regions: {}, managers: {}, clients: clients_list, clientDetail: clientDetail, defaultClient: clients_list.length ? clients_list[0].id : '',
    sourceLists: { groups: regionsFull }, stock: stockOut, ordersInWork: S.inwork, lowStock: low, managersFull: {}, regionsFull: regionsFull,
    abcDetail: { counts: { A: A.length, B: B.length, C: C.length }, revenue: revR, revenue_share: revS, total_rev: pyRound(totRev), total_sku: catalog.length, nclients: nclients, a_list: a_list, reserve: reserve },
    tx: S.tx, txClients: S.tx_clients, txCats: S.tx_cats, txSkus: S.tx_skus, catTotal: catTotal, ym: S.yearMonth, ymGross: S.ymGross, ymRet: S.ymRet,
    regionsData: regionsData, clientRegion: clientRegion, foreignRegions: foreign.length ? foreign : ['Беларусь'], clientFilial: obj(cFil), clientChannel: obj(cChan), clientManager: cMgr,
    clientCity: obj(S.clientCity), clientSubject: obj(S.clientSubject), salesFilial: salesFilial, salesFilial3: salesFilial3,
    refDate: refiso, dataAsOf: refiso, dataAsOfHuman: ymd, catalog: catalog, clientAssort: clientAssort, crm: crmJoined, clientInn: clientInn, turnover: turn };
  Object.keys(upd).forEach(function (k) { K[k] = upd[k]; });
  K.__nExcluded = nExcl;
  return K;
};

/* ───────────── сверка с прошлыми данными и проверки ───────────── */
KI.checkSales = function (rows, S, D) {
  var warns = [], info = {};
  var maxd = String(S.maxdate), iso = maxd.slice(0, 4) + '-' + maxd.slice(4, 6) + '-' + maxd.slice(6);
  info.maxDate = iso; info.net = S.company.net; info.tx = S.tx.length; info.clients = S.client_net.size; info.rows = rows.length;
  if (S.tx.length < 1000) warns.push('Очень мало строк продаж — похоже, это не тот отчёт или выгрузка обрезана.');
  if (D.dataAsOf && iso < D.dataAsOf) warns.push('Новая выгрузка заканчивается ' + iso + ', а в дашборде уже данные на ' + D.dataAsOf + ' — файл старее текущих данных.');
  var ks = Object.keys(S.yearMonth).sort(), firstY = ks.length ? +ks[0].slice(0, 4) : null;
  info.firstYear = firstY;
  // дрейф закрытых периодов: полный прошлый год должен совпасть с тем, что уже в дашборде (если период тот же)
  var oldYm = D.ym || {}, drift = [];
  Object.keys(S.yearMonth).forEach(function (k) {
    if (k >= iso.slice(0, 4) + iso.slice(5, 7)) return;                 // текущий (неполный) месяц не сверяем
    var a = S.yearMonth[k], b = oldYm[k];
    if (b != null && Math.abs(a - b) > Math.max(5000, Math.abs(b) * 0.002)) drift.push(k.slice(4) + '.' + k.slice(0, 4) + ': было ' + Math.round(b / 1e6 * 10) / 10 + ' → стало ' + Math.round(a / 1e6 * 10) / 10 + ' млн');
  });
  info.drift = drift;
  if (drift.length) warns.push('В закрытых месяцах цифры отличаются от прежних (' + drift.length + ' мес.): ' + drift.slice(0, 4).join('; ') + (drift.length > 4 ? '…' : '') + '. Это допустимо, если в 1С правили документы задним числом.');
  var oldYears = Object.keys(oldYm).map(function (k) { return k.slice(0, 4); }), oldMin = oldYears.length ? oldYears.sort()[0] : null;
  if (oldMin && firstY && firstY > +oldMin) warns.push('Новая выгрузка начинается с ' + firstY + ' года, а раньше в дашборде были данные с ' + oldMin + ' — часть истории пропадёт.');
  var known = D.clientRegion || {}, nNew = 0, nNewNoCrm = 0;
  S.client_net.forEach(function (v, nm) { if (v > 0 && !(nm in known)) { nNew++; if (!(cid(nm) in (D.crm || {})) && S.clientChannel.get(nm) !== 'СЕТИ') nNewNoCrm++; } });
  info.newClients = nNew; info.noCrm = nNewNoCrm;
  return { warns: warns, info: info };
};

KI.nrm = nrm; KI.cid = cid; KI.mkPrice = mkPrice;

/* Полный цикл: файлы → новая модель данных. files: {sales: Blob, stock?: Blob, expiry?: Blob}. */
KI.run = async function (files, D, onStage) {
  var P = mkPrice(D.refs || {}), st = onStage || function () {}, rows = null, S = null, stock, expiryRaw, notes = [], head = '', pm = null;
  var cur = (D.refDate ? +String(D.refDate).slice(0, 4) : 2026);
  if (files.sales) {
    st('Читаю файл продаж…', 0);
    rows = await KI.loadSalesRows(files.sales, function (p) { st('Читаю файл продаж…', p); });
    st('Считаю продажи…', 1);
    await new Promise(function (r) { setTimeout(r, 20); });
    var hdr = rows.header || []; head = hdr.map(function (c) { return c.filter(function (x) { return x != null; }).join(' '); }).join('\n');
    if (head.indexOf('Партнер.Бизнес-регион') < 0) throw new Error('Это не отчёт «Продажи по бизнес регионам»: в шапке нет колонки «Партнер.Бизнес-регион». Выгрузите отчёт из 1С в том же виде, что и раньше (иерархия: бизнес-регион → партнёр → заказ → номенклатура).');
    pm = /Период:\s*(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})/.exec(head);
    S = KI.parseSales(rows, P, cur);
    var maxy = Math.floor(S.maxdate / 10000); if (maxy && maxy !== cur) { cur = maxy; S = KI.parseSales(rows, P, cur); }
  } else throw new Error('Нужен файл продаж');
  var chk = KI.checkSales(rows, S, D);
  if (pm) { chk.info.periodFrom = pm[1]; chk.info.periodTo = pm[2]; }
  var fi = head.indexOf('Подразделение'); var fe = head.indexOf('Партнер.Бизнес-регион'); chk.info.filter = fi >= 0 ? head.slice(fi, fe > fi ? fe : fi + 400).replace(/\s+/g, ' ').trim() : '';
  if (files.stock) { st('Читаю остатки…', 0); stock = await KI.loadStock(files.stock, D.refs || {}, function (p) { st('Читаю остатки…', p); }); notes.push('остатки: ' + Object.keys(stock).length + ' артикулов'); }
  else { stock = (D.refs || {}).stock || {}; notes.push('остатки взяты прежние'); }
  if (files.expiry) { st('Читаю сроки годности…', 0); expiryRaw = await KI.loadExpiry(files.expiry); notes.push('сроки годности: ' + Object.keys(expiryRaw).length + ' артикулов'); }
  else { expiryRaw = (D.refs || {}).expiry || {}; }
  st('Собираю модель…', 1);
  await new Promise(function (r) { setTimeout(r, 20); });
  var K = KI.buildKRS(S, stock, expiryRaw, { D: D, P: P });
  K.refs = Object.assign({}, D.refs || {}, { stock: stock, expiry: expiryRaw });
  K.__ingest = { at: new Date().toISOString(), files: { sales: files.sales && files.sales.name || null, stock: files.stock && files.stock.name || null, expiry: files.expiry && files.expiry.name || null },
    warns: chk.warns, info: chk.info, notes: notes };
  return K;
};

if (typeof module !== 'undefined' && module.exports) module.exports = KI; else root.KairosIngest = KI;
})(typeof window !== 'undefined' ? window : globalThis);
