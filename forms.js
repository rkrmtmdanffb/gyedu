/* 행사 서류 자동 채우기: 엑셀(.xlsx) 양식을 올리면 채울 칸을 찾아 두고, 행사 정보를 한 번 입력해 모든 양식에 넣는다.
   파일은 이 브라우저 안에서만 처리하고 서버로 보내지 않는다. */
(function () {
  'use strict';
  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var BASE_FIELDS = [
    { id: 'name', label: '행사명', syn: ['행사명', '행사이름', '행사제목', '제목', '사업명'] },
    { id: 'date', label: '행사일', syn: ['행사일', '행사일시', '행사일자', '일시', '일자', '날짜', '개최일', '개최일시', '기간'] },
    { id: 'pres', label: '회장 이름', syn: ['회장', '회장명', '회장이름', '회장성명'] },
    { id: 'place', label: '장소', syn: ['장소', '행사장소', '개최장소'] },
    { id: 'body', label: '내용', syn: ['내용', '행사내용', '주요내용', '세부내용', '사업내용'] }
  ];
  var KINDS = { cell: '옆(아래) 칸에 넣기', inline: '같은 칸 뒤에 이어 쓰기', token: '{{표시}} 바꾸기' };

  /* ---------- 저장 ---------- */
  var mem = { templates: [], custom: [], values: {} };
  function lsGet(k, d) { try { var v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
  function idb() {
    return new Promise(function (res) {
      try {
        var rq = indexedDB.open('gyedu-forms', 1);
        rq.onupgradeneeded = function () { rq.result.createObjectStore('templates', { keyPath: 'id' }); };
        rq.onsuccess = function () { res(rq.result); };
        rq.onerror = function () { res(null); };
      } catch (e) { res(null); }
    });
  }
  function idbAll() {
    return idb().then(function (db) {
      if (!db) return [];
      return new Promise(function (res) {
        try { var rq = db.transaction('templates').objectStore('templates').getAll(); rq.onsuccess = function () { res(rq.result || []); }; rq.onerror = function () { res([]); }; }
        catch (e) { res([]); }
      });
    });
  }
  function idbPut(t) { return idb().then(function (db) { if (!db) return; try { db.transaction('templates', 'readwrite').objectStore('templates').put(t); } catch (e) {} }); }
  function idbDel(id) { return idb().then(function (db) { if (!db) return; try { db.transaction('templates', 'readwrite').objectStore('templates').delete(id); } catch (e) {} }); }

  function fields() { return BASE_FIELDS.concat(mem.custom.map(function (n) { return { id: 'c:' + n, label: n, syn: [n] }; })); }
  function fieldById(id) { return fields().filter(function (f) { return f.id === id; })[0]; }

  /* ---------- 엑셀 읽기 ---------- */
  function norm(s) { return String(s || '').replace(/[\s ▮■□●○◆◇▶▷※·ㆍ\-\*\[\]\(\)<>〈〉《》「」]/g, ''); }
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  function colName(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
  function parseAddr(a) { var m = /^([A-Z]+)(\d+)$/.exec(String(a || '').toUpperCase().replace(/\$/g, '').trim()); return m ? { c: colNum(m[1]), r: +m[2] } : null; }
  function addr(r, c) { return colName(c) + r; }
  function xml(s) { return new DOMParser().parseFromString(s, 'application/xml'); }
  function kids(el, name) { var out = []; for (var n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && n.localName === name) out.push(n); return out; }
  function siText(si) {
    var out = '';
    (function walk(n) { for (var k = n.firstChild; k; k = k.nextSibling) { if (k.nodeType !== 1) continue; if (k.localName === 'rPh') continue; if (k.localName === 't') out += k.textContent; else walk(k); } })(si);
    return out;
  }

  function openBook(buf) {
    var book = { sheets: [], shared: [] };
    return JSZip.loadAsync(buf).then(function (zip) {
      book.zip = zip;
      var wbFile = zip.file('xl/workbook.xml'), relFile = zip.file('xl/_rels/workbook.xml.rels');
      if (!wbFile || !relFile) throw new Error('not-xlsx');
      return Promise.all([wbFile.async('string'), relFile.async('string'), zip.file('xl/sharedStrings.xml') ? zip.file('xl/sharedStrings.xml').async('string') : '']);
    }).then(function (r) {
      var wb = xml(r[0]), rels = xml(r[1]), map = {};
      Array.prototype.forEach.call(rels.getElementsByTagName('Relationship'), function (e) {
        var t = e.getAttribute('Target'); map[e.getAttribute('Id')] = t.charAt(0) === '/' ? t.slice(1) : 'xl/' + t;
      });
      if (r[2]) Array.prototype.forEach.call(xml(r[2]).getElementsByTagName('si'), function (si) { book.shared.push(siText(si)); });
      var list = Array.prototype.map.call(wb.getElementsByTagName('sheet'), function (s) {
        return { name: s.getAttribute('name'), path: map[s.getAttribute('r:id')] };
      }).filter(function (s) { return s.path && book.zip.file(s.path); });
      return Promise.all(list.map(function (s) {
        return book.zip.file(s.path).async('string').then(function (x) {
          s.doc = xml(x); s.cells = {}; s.merges = [];
          Array.prototype.forEach.call(s.doc.getElementsByTagName('c'), function (c) {
            var a = c.getAttribute('r'); if (!a) return;
            var t = c.getAttribute('t'), v = kids(c, 'v')[0], text = '';
            if (t === 's') text = v ? (book.shared[+v.textContent] || '') : '';
            else if (t === 'inlineStr') { var is = kids(c, 'is')[0]; text = is ? siText(is) : ''; }
            else text = v ? v.textContent : '';
            s.cells[a] = { el: c, text: text };
          });
          Array.prototype.forEach.call(s.doc.getElementsByTagName('mergeCell'), function (m) {
            var p = (m.getAttribute('ref') || '').split(':'), a = parseAddr(p[0]), b = parseAddr(p[1] || p[0]);
            if (a && b) s.merges.push({ r1: a.r, c1: a.c, r2: b.r, c2: b.c });
          });
          book.sheets.push(s);
        });
      })).then(function () {
        book.sheets.sort(function (a, b) { return list.indexOf(a) - list.indexOf(b); });
        return book;
      });
    });
  }
  function mergeAt(s, r, c) { for (var i = 0; i < s.merges.length; i++) { var m = s.merges[i]; if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m; } return null; }
  function master(s, r, c) { var m = mergeAt(s, r, c); return m ? addr(m.r1, m.c1) : addr(r, c); }
  function textAt(s, a) { return s.cells[a] ? s.cells[a].text : ''; }

  function synIndex() {
    var idx = {};
    fields().forEach(function (f) { f.syn.forEach(function (w) { var k = norm(w); if (!(k in idx)) idx[k] = f.id; }); });
    return idx;
  }

  /* 채울 칸 찾기 */
  function detect(book) {
    var idx = synIndex(), found = [], seen = {};
    function add(m) { var k = m.sheet + '!' + m.addr + '!' + m.field; if (seen[k]) return; seen[k] = 1; m.on = true; found.push(m); }
    book.sheets.forEach(function (s) {
      Object.keys(s.cells).forEach(function (a) {
        var text = s.cells[a].text; if (!text || !String(text).trim()) return;
        var p = parseAddr(a); if (!p) return;
        if (master(s, p.r, p.c) !== a) return;
        fields().forEach(function (f) {
          if (text.indexOf('{{' + f.label + '}}') >= 0 || text.indexOf('{{' + f.label.replace(/\s/g, '') + '}}') >= 0)
            add({ field: f.id, sheet: s.name, addr: a, kind: 'token', label: text });
        });
        var n = norm(text).replace(/[:：]$/, '');
        if (idx[n] && !/[:：]/.test(text)) {
          var m = mergeAt(s, p.r, p.c), rc = (m ? m.c2 : p.c) + 1, right = master(s, p.r, rc), rt = norm(textAt(s, right)).replace(/[:：]$/, '');
          var target = right;
          if (rt && (idx[rt] || /[:：]/.test(textAt(s, right)) && idx[norm(textAt(s, right).split(/[:：]/)[0])])) {
            target = master(s, (m ? m.r2 : p.r) + 1, p.c);
          }
          add({ field: idx[n], sheet: s.name, addr: target, kind: 'cell', label: text.trim(), cur: textAt(s, target) });
          return;
        }
        var ci = text.search(/[:：]/);
        if (ci > 0 && ci <= 14) {
          var head = norm(text.slice(0, ci));
          if (idx[head]) add({ field: idx[head], sheet: s.name, addr: a, kind: 'inline', label: text.slice(0, ci + 1).trim(), cur: text.slice(ci + 1).trim() });
        }
      });
    });
    var order = fields().map(function (f) { return f.id; });
    found.sort(function (x, y) { return order.indexOf(x.field) - order.indexOf(y.field); });
    return found;
  }

  /* 칸에 글자 쓰기 */
  function setCell(s, a, value) {
    var p = parseAddr(a); if (!p) return false;
    a = addr(p.r, p.c);
    var doc = s.doc, ns = doc.documentElement.namespaceURI, data = doc.getElementsByTagName('sheetData')[0];
    if (!data) return false;
    var rows = kids(data, 'row'), row = null, before = null, i;
    for (i = 0; i < rows.length; i++) { var rn = +rows[i].getAttribute('r'); if (rn === p.r) { row = rows[i]; break; } if (rn > p.r) { before = rows[i]; break; } }
    if (!row) { row = doc.createElementNS(ns, 'row'); row.setAttribute('r', String(p.r)); data.insertBefore(row, before); }
    var cs = kids(row, 'c'), cell = null, cb = null, hadFormula = false;
    for (i = 0; i < cs.length; i++) { var q = parseAddr(cs[i].getAttribute('r')); if (!q) continue; if (q.c === p.c) { cell = cs[i]; break; } if (q.c > p.c) { cb = cs[i]; break; } }
    if (!cell) { cell = doc.createElementNS(ns, 'c'); cell.setAttribute('r', a); row.insertBefore(cell, cb); }
    if (kids(cell, 'f').length) hadFormula = true;
    while (cell.firstChild) cell.removeChild(cell.firstChild);
    cell.setAttribute('t', 'inlineStr');
    var is = doc.createElementNS(ns, 'is'), t = doc.createElementNS(ns, 't');
    t.setAttribute('xml:space', 'preserve'); t.textContent = value; is.appendChild(t); cell.appendChild(is);
    s.cells[a] = { el: cell, text: value }; s.dirty = true;
    return hadFormula ? 'formula' : true;
  }

  function serialize(doc) { var out = new XMLSerializer().serializeToString(doc); return out.indexOf('<?xml') === 0 ? out : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out; }

  function fill(tpl, values) {
    var log = [];
    return openBook(tpl.buf.slice(0)).then(function (book) {
      var dropCalc = false;
      tpl.maps.forEach(function (m) {
        if (!m.on) return;
        var f = fieldById(m.field), v = values[m.field];
        if (!f || v == null || v === '') return;
        var s = book.sheets.filter(function (x) { return x.name === m.sheet; })[0], p = parseAddr(m.addr);
        if (!s || !p) { log.push({ ok: false, f: f.label, where: m.sheet + ' ' + m.addr, note: '칸을 찾지 못했습니다' }); return; }
        var a = master(s, p.r, p.c), cur = textAt(s, a), out = v;
        if (m.kind === 'inline') { var ci = cur.search(/[:：]/); out = (ci >= 0 ? cur.slice(0, ci + 1) + ' ' : (cur ? cur + ' ' : '')) + v; }
        else if (m.kind === 'token') { out = cur.split('{{' + f.label + '}}').join(v).split('{{' + f.label.replace(/\s/g, '') + '}}').join(v); }
        var r = setCell(s, a, out);
        if (r === 'formula') dropCalc = true;
        log.push({ ok: !!r, f: f.label, where: m.sheet + ' ' + a, note: out });
      });
      book.sheets.forEach(function (s) { if (s.dirty) book.zip.file(s.path, serialize(s.doc)); });
      var jobs = [];
      if (dropCalc && book.zip.file('xl/calcChain.xml')) {
        book.zip.remove('xl/calcChain.xml');
        jobs.push(book.zip.file('xl/_rels/workbook.xml.rels').async('string').then(function (x) {
          var d = xml(x); Array.prototype.slice.call(d.getElementsByTagName('Relationship')).forEach(function (e) { if (/calcChain/.test(e.getAttribute('Target'))) e.parentNode.removeChild(e); });
          book.zip.file('xl/_rels/workbook.xml.rels', serialize(d));
        }));
        jobs.push(book.zip.file('[Content_Types].xml').async('string').then(function (x) {
          var d = xml(x); Array.prototype.slice.call(d.getElementsByTagName('Override')).forEach(function (e) { if (/calcChain/.test(e.getAttribute('PartName'))) e.parentNode.removeChild(e); });
          book.zip.file('[Content_Types].xml', serialize(d));
        }));
      }
      return Promise.all(jobs).then(function () { return book.zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: XLSX_MIME }); });
    }).then(function (blob) { return { blob: blob, log: log }; });
  }

  /* ---------- 내려받기 ---------- */
  function save(filename, blob) {
    function plain() {
      var u = URL.createObjectURL(blob), a = document.createElement('a');
      a.href = u; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(u); }, 4000);
      return 'saved';
    }
    if (!(window.claude && typeof window.claude.use === 'function')) return Promise.resolve(plain());
    return window.claude.use('downloads').then(function (d) {
      if (!d) return plain();
      return d.save({ filename: filename, data: blob }).then(function () { return 'saved'; }, function (e) {
        if (e && e.code === 'declined') return 'declined';
        if (e && e.code === 'rate_limited') return 'busy';
        return 'failed';
      });
    }, function () { return plain(); });
  }

  /* ---------- 화면 ---------- */
  function h(tag, props, children) {
    var n = document.createElement(tag), k;
    if (props) for (k in props) { if (k === 'class') n.className = props[k]; else if (k === 'text') n.textContent = props[k]; else if (k.slice(0, 2) === 'on') n.addEventListener(k.slice(2), props[k]); else if (props[k] != null && props[k] !== false) n.setAttribute(k, props[k] === true ? '' : props[k]); }
    (children || []).forEach(function (c) { if (c != null) n.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
    return n;
  }
  var $tpl, $vals, $out, $msg;
  function say(text, bad) { $msg.textContent = text || ''; $msg.className = 'fmsg' + (bad ? ' bad' : ''); }
  function fmtDate(d, time) {
    if (!d) return '';
    var p = d.split('-'), dt = new Date(+p[0], +p[1] - 1, +p[2]);
    return p[0] + '년 ' + (+p[1]) + '월 ' + (+p[2]) + '일(' + '일월화수목금토'.charAt(dt.getDay()) + ')' + (time ? ' ' + time : '');
  }
  function currentValues() {
    var v = {};
    fields().forEach(function (f) {
      if (f.id === 'date') { v.date = fmtDate(document.getElementById('fv-date').value, document.getElementById('fv-time').value.trim()); return; }
      var el = document.getElementById('fv-' + f.id); v[f.id] = el ? el.value.trim() : '';
    });
    return v;
  }
  function rememberValues() {
    var raw = { date: document.getElementById('fv-date').value, time: document.getElementById('fv-time').value };
    fields().forEach(function (f) { if (f.id !== 'date') { var el = document.getElementById('fv-' + f.id); if (el) raw[f.id] = el.value; } });
    mem.values = raw; lsSet('gyedu-values', raw);
  }

  function drawValues() {
    var raw = mem.values || {};
    $vals.textContent = '';
    fields().forEach(function (f) {
      var id = 'fv-' + f.id, ctl;
      if (f.id === 'date') {
        ctl = h('div', { class: 'frow2' }, [
          h('input', { type: 'date', id: 'fv-date', value: raw.date || '', onchange: rememberValues }),
          h('input', { type: 'text', id: 'fv-time', placeholder: '시간 (예: 오후 5시 30분)', value: raw.time || '', 'aria-label': '행사 시간', oninput: rememberValues })
        ]);
        $vals.appendChild(h('div', { class: 'ffield' }, [h('label', { for: 'fv-date', text: f.label }), ctl]));
        return;
      }
      if (f.id === 'body') { ctl = h('textarea', { id: id, rows: '4', oninput: rememberValues }); ctl.value = raw[f.id] || ''; }
      else ctl = h('input', { type: 'text', id: id, value: raw[f.id] || '', oninput: rememberValues });
      var lab = h('label', { for: id, text: f.label });
      var box = h('div', { class: 'ffield' + (f.id === 'body' ? ' wide' : '') }, [lab, ctl]);
      if (f.id.indexOf('c:') === 0) lab.appendChild(h('button', { type: 'button', class: 'flink', text: '항목 삭제', onclick: function () { rememberValues(); mem.custom = mem.custom.filter(function (n) { return 'c:' + n !== f.id; }); lsSet('gyedu-custom', mem.custom); drawValues(); drawTemplates(); } }));
      $vals.appendChild(box);
    });
  }

  function saveTpl(t) { idbPut({ id: t.id, name: t.name, buf: t.buf, maps: t.maps, sheets: t.sheets }); }

  function drawTemplates() {
    $tpl.textContent = '';
    if (!mem.templates.length) { $tpl.appendChild(h('p', { class: 'fempty', text: '아직 올린 양식이 없습니다. 위 버튼으로 엑셀 양식(.xlsx)을 올려 주세요.' })); return; }
    mem.templates.forEach(function (t) {
      var body = h('div', { class: 'fmaps' });
      function row(m) {
        var fsel = h('select', { 'aria-label': '항목', onchange: function () { m.field = fsel.value; saveTpl(t); } }, fields().map(function (f) { return h('option', { value: f.id, text: f.label, selected: f.id === m.field }); }));
        var ssel = h('select', { 'aria-label': '시트', onchange: function () { m.sheet = ssel.value; saveTpl(t); } }, t.sheets.map(function (n) { return h('option', { value: n, text: n, selected: n === m.sheet }); }));
        var ain = h('input', { type: 'text', value: m.addr, 'aria-label': '채울 칸', class: 'faddr', onchange: function () { var p = parseAddr(ain.value); if (!p) { ain.value = m.addr; say('칸 이름은 B3처럼 적어 주세요.', true); return; } m.addr = addr(p.r, p.c); ain.value = m.addr; saveTpl(t); say('저장했습니다.'); } });
        var ksel = h('select', { 'aria-label': '넣는 방식', onchange: function () { m.kind = ksel.value; saveTpl(t); } }, Object.keys(KINDS).map(function (k) { return h('option', { value: k, text: KINDS[k], selected: k === m.kind }); }));
        var on = h('input', { type: 'checkbox', 'aria-label': '사용', onchange: function () { m.on = on.checked; saveTpl(t); } }); on.checked = !!m.on;
        var note = m.label ? '찾은 글자: ' + String(m.label).slice(0, 30) + (m.cur ? ' · 지금 값: ' + String(m.cur).slice(0, 30) : '') : '직접 추가한 칸';
        return h('div', { class: 'fmap' }, [on, fsel, ssel, ain, ksel, h('small', { text: note }),
          h('button', { type: 'button', class: 'flink', text: '삭제', onclick: function () { t.maps.splice(t.maps.indexOf(m), 1); saveTpl(t); drawTemplates(); } })]);
      }
      if (!t.maps.length) body.appendChild(h('p', { class: 'fempty', text: '채울 칸을 자동으로 찾지 못했습니다. 아래 버튼으로 칸을 직접 정해 주세요.' }));
      t.maps.forEach(function (m) { body.appendChild(row(m)); });
      var card = h('div', { class: 'ftpl' }, [
        h('div', { class: 'ftpl-h' }, [h('strong', { text: t.name }), h('span', { class: 'pill', text: '채울 칸 ' + t.maps.filter(function (m) { return m.on; }).length + '곳' }),
          h('button', { type: 'button', class: 'flink', text: '다시 찾기', onclick: function () { openBook(t.buf.slice(0)).then(function (b) { t.maps = detect(b); saveTpl(t); drawTemplates(); say('채울 칸을 다시 찾았습니다.'); }); } }),
          h('button', { type: 'button', class: 'flink', text: '양식 삭제', onclick: function () { mem.templates = mem.templates.filter(function (x) { return x !== t; }); idbDel(t.id); drawTemplates(); } })]),
        h('div', { class: 'fmap fhead' }, [h('span', { text: '사용' }), h('span', { text: '항목' }), h('span', { text: '시트' }), h('span', { text: '채울 칸' }), h('span', { text: '넣는 방식' }), h('span', { text: '' }), h('span', { text: '' })]),
        body,
        h('button', { type: 'button', class: 'chip', text: '+ 채울 칸 직접 추가', onclick: function () { t.maps.push({ field: 'name', sheet: t.sheets[0], addr: 'A1', kind: 'cell', on: true }); saveTpl(t); drawTemplates(); } })
      ]);
      $tpl.appendChild(card);
    });
  }

  function addFiles(list) {
    var files = Array.prototype.slice.call(list), chain = Promise.resolve(), okN = 0, bad = [];
    files.forEach(function (file) {
      chain = chain.then(function () {
        if (!/\.xlsx$/i.test(file.name)) { bad.push(file.name + ' (xlsx 파일이 아닙니다)'); return; }
        return file.arrayBuffer().then(function (buf) {
          return openBook(buf.slice(0)).then(function (book) {
            var t = { id: 't' + Date.now() + Math.random().toString(36).slice(2, 7), name: file.name, buf: buf, maps: detect(book), sheets: book.sheets.map(function (s) { return s.name; }) };
            mem.templates.push(t); saveTpl(t); okN++;
          });
        }).catch(function () { bad.push(file.name + ' (읽지 못했습니다)'); });
      });
    });
    return chain.then(function () {
      drawTemplates();
      if (bad.length) say('올리지 못한 파일: ' + bad.join(', ') + '. 엑셀에서 "다른 이름으로 저장 → Excel 통합 문서(.xlsx)"로 저장해 다시 올려 주세요.', true);
      else say(okN + '개 양식을 올렸습니다. 아래에서 채울 칸을 확인해 주세요.');
    });
  }

  function run() {
    var values = currentValues();
    if (!mem.templates.length) { say('먼저 엑셀 양식을 올려 주세요.', true); return; }
    if (!Object.keys(values).some(function (k) { return values[k]; })) { say('행사 정보를 한 가지 이상 입력해 주세요.', true); return; }
    rememberValues(); $out.textContent = ''; say('채우는 중입니다…');
    var stamp = (document.getElementById('fv-date').value || '').replace(/-/g, '');
    Promise.all(mem.templates.map(function (t) {
      return fill(t, values).then(function (r) { return { t: t, r: r }; }, function () { return { t: t, r: null }; });
    })).then(function (all) {
      all.forEach(function (x) {
        if (!x.r) { $out.appendChild(h('div', { class: 'file' }, [h('span', { class: 'meta', text: x.t.name + ' — 채우지 못했습니다.' })])); return; }
        var base = x.t.name.replace(/\.xlsx$/i, ''), fname = ((values.name ? values.name + '_' : '') + base + (stamp ? '_' + stamp : '') + '.xlsx').replace(/[\\\/:*?"<>|]/g, ' ');
        var done = x.r.log.filter(function (l) { return l.ok; });
        var status = h('small', { text: '' });
        var btn = h('button', { type: 'button', class: 'act', text: '내려받기', onclick: function () {
          save(fname, x.r.blob).then(function (res) { status.textContent = res === 'saved' ? ' 저장을 요청했습니다.' : res === 'declined' ? ' 저장을 취소했습니다.' : res === 'busy' ? ' 잠시 뒤 다시 눌러 주세요.' : ' 이 화면에서는 저장할 수 없습니다.'; });
        } });
        var detail = h('ul', { class: 'flog' }, x.r.log.map(function (l) { return h('li', { text: (l.ok ? '' : '실패 · ') + l.f + ' → ' + l.where + ' : ' + String(l.note).slice(0, 60) }); }));
        $out.appendChild(h('div', { class: 'ftpl' }, [h('div', { class: 'ftpl-h' }, [h('strong', { text: fname }), h('span', { class: 'pill', text: done.length + '곳 채움' }), btn, status]), detail]));
      });
      say('완성했습니다. 아래에서 파일을 내려받으세요.');
      $out.scrollIntoView({ block: 'nearest' });
    });
  }

  function init() {
    $tpl = document.getElementById('f-templates'); $vals = document.getElementById('f-values'); $out = document.getElementById('f-results'); $msg = document.getElementById('f-msg');
    if (!$tpl) return;
    mem.custom = lsGet('gyedu-custom', []); mem.values = lsGet('gyedu-values', {});
    drawValues(); drawTemplates();
    idbAll().then(function (list) { if (list.length) { mem.templates = list.concat(mem.templates); drawTemplates(); } });
    var input = document.getElementById('f-file');
    input.addEventListener('change', function () { if (input.files.length) addFiles(input.files).then(function () { input.value = ''; }); });
    var drop = document.getElementById('f-drop');
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
    drop.addEventListener('drop', function (e) { e.preventDefault(); drop.classList.remove('over'); if (e.dataTransfer && e.dataTransfer.files.length) addFiles(e.dataTransfer.files); });
    document.getElementById('f-run').addEventListener('click', run);
    document.getElementById('f-addfield').addEventListener('submit', function (e) {
      e.preventDefault();
      var inp = document.getElementById('f-newfield'), n = inp.value.trim();
      if (!n) return;
      if (fields().some(function (f) { return f.label === n; })) { say('이미 있는 항목입니다.', true); return; }
      rememberValues(); mem.custom.push(n); lsSet('gyedu-custom', mem.custom); inp.value = ''; drawValues(); drawTemplates();
      say('"' + n + '" 항목을 추가했습니다. 양식에서 "다시 찾기"를 누르면 이 항목의 칸도 찾습니다.');
    });
  }
  window.GyeduForms = { openBook: openBook, detect: detect, fill: fill, save: save };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
