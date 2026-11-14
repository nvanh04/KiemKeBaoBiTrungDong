/* shim.js - thay thế lucide, html5-qrcode, xlsx. Không cần mạng, không cần thư viện ngoài. */
(function () {
'use strict';

/* ---------- 1. ICON (thay lucide) ---------- */
var G = {boxes:'📦','clipboard-edit':'📝',history:'🕘','shield-check':'🛡️','building-2':'🏢',camera:'📷',check:'✔','plus-circle':'➕','list-checks':'📋',save:'💾',calendar:'📅','rotate-ccw':'↺','file-spreadsheet':'📊','trash-2':'🗑️',lock:'🔒','shield-alert':'⚠️','log-out':'🚪',factory:'🏭',plus:'＋','folder-tree':'🗂️','settings-2':'⚙️',download:'⬇️',upload:'⬆️',pencil:'✏️',scan:'🔍',x:'✕','alert-triangle':'⚠️','check-circle-2':'✅','alert-circle':'❗'};
window.lucide = { createIcons: function () {
  document.querySelectorAll('[data-lucide]').forEach(function (el) {
    var g = G[el.getAttribute('data-lucide')] || '•';
    if (el.textContent !== g) el.textContent = g;
    var s = el.style;
    s.display = 'inline-block'; s.fontStyle = 'normal'; s.lineHeight = '1'; s.textAlign = 'center';
    var w = el.getBoundingClientRect().width;
    s.fontSize = (w > 0 ? Math.round(w * 0.85) : 13) + 'px';
  });
}};

/* ---------- 2. QUÉT QR (thay html5-qrcode, dùng BarcodeDetector của trình duyệt) ---------- */
function Q(id) { this.el = document.getElementById(id); this.stream = null; this.timer = null; this.dead = false; }
Q.prototype.start = async function (cam, cfg, onOk) {
  var el = this.el, self = this;
  if (!('BarcodeDetector' in window) || !navigator.mediaDevices) {
    el.innerHTML = '<div style="padding:16px;color:#fff;font-size:13px;text-align:center">Trình duyệt này chưa hỗ trợ quét QR trực tiếp.<br>Hãy đóng cửa sổ và chọn vật tư ở ô <b>"Hoặc chọn nhanh từ danh mục"</b>.</div>';
    return;
  }
  var det = new BarcodeDetector({ formats: ['qr_code'] });
  var st = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' } }, audio: false });
  if (this.dead) { st.getTracks().forEach(function (t) { t.stop(); }); return; }
  this.stream = st;
  var v = document.createElement('video');
  v.setAttribute('playsinline', 'true'); v.muted = true; v.style.minHeight = '240px';
  v.srcObject = st; el.innerHTML = ''; el.appendChild(v);
  await v.play();
  var busy = false;
  this.timer = setInterval(async function () {
    if (busy || v.readyState < 2) return;
    busy = true;
    try {
      var r = await det.detect(v);
      if (r.length) { clearInterval(self.timer); self.timer = null; onOk(r[0].rawValue); }
    } catch (e) {}
    busy = false;
  }, 200);
};
Q.prototype.stop = function () {
  this.dead = true; clearInterval(this.timer); this.timer = null;
  if (this.stream) this.stream.getTracks().forEach(function (t) { t.stop(); });
  this.stream = null;
  return Promise.resolve();
};
Q.prototype.clear = function () { this.el.innerHTML = ''; };
window.Html5Qrcode = Q;

/* ---------- 3. EXCEL (thay xlsx): ghi/đọc .xlsx, đọc .csv ---------- */
function colName(i) { var s = ''; for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + (i - 1) % 26) + s; return s; }
function colIdx(ref) { var m = /^[A-Z]+/.exec(ref)[0], n = 0; for (var i = 0; i < m.length; i++) n = n * 26 + m.charCodeAt(i) - 64; return n - 1; }
function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

var T = (function () { var t = []; for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc(b) { var c = -1; for (var i = 0; i < b.length; i++) c = T[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ -1) >>> 0; }

function zip(files) { // trả về mảng các phần để đưa vào Blob (lưu không nén)
  var parts = [], cd = [], off = 0, cs = 0, enc = new TextEncoder();
  files.forEach(function (f) {
    var nm = enc.encode(f.n), c = crc(f.d), n = f.d.length;
    var h = new DataView(new ArrayBuffer(30));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x800, true); h.setUint16(12, 0x21, true);
    h.setUint32(14, c, true); h.setUint32(18, n, true); h.setUint32(22, n, true); h.setUint16(26, nm.length, true);
    var e = new DataView(new ArrayBuffer(46));
    e.setUint32(0, 0x02014b50, true); e.setUint16(4, 20, true); e.setUint16(6, 20, true); e.setUint16(8, 0x800, true); e.setUint16(14, 0x21, true);
    e.setUint32(16, c, true); e.setUint32(20, n, true); e.setUint32(24, n, true); e.setUint16(28, nm.length, true); e.setUint32(42, off, true);
    parts.push(h.buffer, nm, f.d); cd.push(e.buffer, nm);
    off += 30 + nm.length + n; cs += 46 + nm.length;
  });
  var z = new DataView(new ArrayBuffer(22));
  z.setUint32(0, 0x06054b50, true); z.setUint16(8, files.length, true); z.setUint16(10, files.length, true);
  z.setUint32(12, cs, true); z.setUint32(16, off, true);
  return parts.concat(cd, [z.buffer]);
}

function sheetXml(rows) {
  var keys = [];
  rows.forEach(function (r) { Object.keys(r).forEach(function (k) { if (keys.indexOf(k) < 0) keys.push(k); }); });
  function cell(r, c, v) {
    var ref = colName(c) + r;
    if (typeof v === 'number' && isFinite(v)) return '<c r="' + ref + '"><v>' + v + '</v></c>';
    return '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v == null ? '' : v) + '</t></is></c>';
  }
  var x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';
  x += '<row r="1">' + keys.map(function (k, c) { return cell(1, c, k); }).join('') + '</row>';
  rows.forEach(function (o, i) { x += '<row r="' + (i + 2) + '">' + keys.map(function (k, c) { return cell(i + 2, c, o[k]); }).join('') + '</row>'; });
  return x + '</sheetData></worksheet>';
}

function rowsToWb(t) {
  var h = (t[0] || []).map(function (x) { return String(x == null ? '' : x).trim(); }), rows = [];
  for (var i = 1; i < t.length; i++) {
    var r = t[i]; if (!r) continue;
    var o = {}, any = false;
    h.forEach(function (k, c) { if (!k) return; var v = r[c] == null ? '' : r[c]; if (v !== '') any = true; o[k] = v; });
    if (any) rows.push(o);
  }
  return { SheetNames: ['Sheet1'], Sheets: { Sheet1: { rows: rows } } };
}

function parseCsv(t) {
  t = t.replace(/^\uFEFF/, '');
  var first = t.split(/\r?\n/)[0], d = first.split(';').length > first.split(',').length ? ';' : ',';
  var rows = [], row = [], f = '', q = false;
  for (var i = 0; i < t.length; i++) {
    var c = t[i];
    if (q) { if (c === '"') { if (t[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true;
    else if (c === d) { row.push(f); f = ''; }
    else if (c === '\n' || c === '\r') { if (c === '\r' && t[i + 1] === '\n') i++; row.push(f); f = ''; rows.push(row); row = []; }
    else f += c;
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return rows;
}

async function inflate(u8) {
  var ds = new DecompressionStream('deflate-raw'), w = ds.writable.getWriter();
  w.write(u8); w.close();
  return new Uint8Array(await new Response(ds.readable).arrayBuffer());
}

async function readXlsx(b) {
  var dv = new DataView(b.buffer, b.byteOffset, b.byteLength), dec = new TextDecoder(), i = b.length - 22;
  while (i >= 0 && dv.getUint32(i, true) !== 0x06054b50) i--;
  if (i < 0) throw new Error('zip');
  var n = dv.getUint16(i + 10, true), p = dv.getUint32(i + 16, true), ent = {};
  for (var k = 0; k < n; k++) {
    var nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
    ent[dec.decode(b.subarray(p + 46, p + 46 + nl))] = { m: dv.getUint16(p + 10, true), cs: dv.getUint32(p + 20, true), o: dv.getUint32(p + 42, true) };
    p += 46 + nl + el + cl;
  }
  async function get(name) {
    var e = ent[name]; if (!e) return null;
    var s = e.o + 30 + dv.getUint16(e.o + 26, true) + dv.getUint16(e.o + 28, true), d = b.subarray(s, s + e.cs);
    return dec.decode(e.m === 0 ? d : await inflate(d));
  }
  var num = function (x) { return parseInt(x.match(/(\d+)\.xml$/)[1], 10); };
  var names = Object.keys(ent).filter(function (x) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(x); }).sort(function (x, y) { return num(x) - num(y); });
  if (!names.length) throw new Error('sheet');
  var P = new DOMParser(), ss = [], sx = await get('xl/sharedStrings.xml');
  if (sx) [].forEach.call(P.parseFromString(sx, 'text/xml').getElementsByTagName('si'), function (si) {
    ss.push([].map.call(si.getElementsByTagName('t'), function (t) { return t.textContent; }).join(''));
  });
  var doc = P.parseFromString(await get(names[0]), 'text/xml'), tbl = [];
  [].forEach.call(doc.getElementsByTagName('c'), function (c) {
    var ref = c.getAttribute('r'), r = parseInt(ref.replace(/\D/g, ''), 10) - 1, col = colIdx(ref),
        t = c.getAttribute('t'), v = c.getElementsByTagName('v')[0], val = '';
    if (t === 'inlineStr') val = [].map.call(c.getElementsByTagName('t'), function (x) { return x.textContent; }).join('');
    else if (v) val = t === 's' ? ss[+v.textContent] : v.textContent;
    (tbl[r] = tbl[r] || [])[col] = val;
  });
  return rowsToWb(tbl);
}

window.XLSX = {
  utils: {
    json_to_sheet: function (a) { return { rows: a || [] }; },
    book_new: function () { return { SheetNames: [], Sheets: {} }; },
    book_append_sheet: function (wb, ws, n) { wb.SheetNames.push(n); wb.Sheets[n] = ws; },
    sheet_to_json: function (ws) { return ws.rows; }
  },
  read: function () { var w = XLSX._pre; XLSX._pre = null; if (!w) throw new Error('empty'); return w; },
  writeFile: function (wb, name) {
    var H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        PR = 'http://schemas.openxmlformats.org/package/2006/relationships',
        CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.',
        ns = wb.SheetNames, enc = new TextEncoder();
    var files = [
      ['[Content_Types].xml', H + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="' + CT + 'sheet.main+xml"/>' + ns.map(function (_, i) { return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="' + CT + 'worksheet+xml"/>'; }).join('') + '</Types>'],
      ['_rels/.rels', H + '<Relationships xmlns="' + PR + '"><Relationship Id="rId1" Type="' + R + '/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
      ['xl/workbook.xml', H + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="' + R + '"><sheets>' + ns.map(function (n, i) { return '<sheet name="' + esc(n) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join('') + '</sheets></workbook>'],
      ['xl/_rels/workbook.xml.rels', H + '<Relationships xmlns="' + PR + '">' + ns.map(function (_, i) { return '<Relationship Id="rId' + (i + 1) + '" Type="' + R + '/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>'; }).join('') + '</Relationships>']
    ];
    ns.forEach(function (n, i) { files.push(['xl/worksheets/sheet' + (i + 1) + '.xml', sheetXml(wb.Sheets[n].rows)]); });
    var blob = new Blob(zip(files.map(function (f) { return { n: f[0], d: enc.encode(f[1]) }; })), { type: CT + 'sheet' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
};

/* Nhập Excel: đọc file (bất đồng bộ) rồi chuyển cho handleExcelImport có sẵn của ứng dụng */
document.addEventListener('change', function (e) {
  var inp = e.target;
  if (!inp || inp.id !== 'excel-file-input') return;
  var f = inp.files && inp.files[0];
  if (!f) return;
  e.stopImmediatePropagation();
  f.arrayBuffer().then(function (buf) {
    var b = new Uint8Array(buf);
    if (b[0] === 0x50 && b[1] === 0x4B) return readXlsx(b);
    if (b[0] === 0xD0 && b[1] === 0xCF) throw new Error('xls');
    return rowsToWb(parseCsv(new TextDecoder('utf-8').decode(b)));
  }).then(function (wb) {
    XLSX._pre = wb;
    window.handleExcelImport({ target: { files: [new Blob([new Uint8Array(1)])], value: '' } });
    inp.value = '';
  }).catch(function () {
    if (window.showToast) window.showToast('Không đọc được file. Hãy lưu lại dạng .xlsx hoặc .csv rồi thử lại', 'error');
    inp.value = '';
  });
}, true);

})();
