/* 회원 전용 명부와 임원 전용 자료.
   내용은 비밀번호로 암호화(AES-GCM, 키는 PBKDF2)해서 secure 폴더에 두고, 맞는 비밀번호를 넣은 브라우저에서만 풀린다.
   회원 비밀번호: 명부만.  임원 비밀번호: 명부 + 임원 전용 자료. */
(function () {
  'use strict';
  var ITER = 250000, te = new TextEncoder(), td = new TextDecoder();
  function b64(buf) { var u = new Uint8Array(buf), s = '', i; for (i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }
  function unb64(s) { var b = atob(s), u = new Uint8Array(b.length), i; for (i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; }
  function rand(n) { return crypto.getRandomValues(new Uint8Array(n)); }
  function kdf(pw, salt, iter) {
    return crypto.subtle.importKey('raw', te.encode(pw), 'PBKDF2', false, ['deriveKey']).then(function (base) {
      return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: salt, iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
    });
  }
  function rawKey(bytes) { return crypto.subtle.importKey('raw', bytes, { name: 'AES-GCM' }, false, ['encrypt', 'decrypt']); }
  function seal(key, bytes) { var iv = rand(12); return crypto.subtle.encrypt({ name: 'AES-GCM', iv: iv }, key, bytes).then(function (ct) { return { iv: b64(iv), ct: b64(ct) }; }); }
  function unseal(key, box) { return crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(box.iv) }, key, unb64(box.ct)).then(function (p) { return new Uint8Array(p); }); }
  function json(bytes) { return JSON.parse(td.decode(bytes)); }

  var vault = null, vaultTried = null, state = { dataKey: null, fileKey: null, roster: null, files: null };
  function loadVault() {
    if (vaultTried) return vaultTried;
    vaultTried = fetch('secure/vault.json', { cache: 'no-store' }).then(function (r) { return r.ok ? r.json() : null; }).then(function (v) { vault = v && v.v === 1 ? v : null; return vault; }, function () { return null; });
    return vaultTried;
  }
  function openRoster() { return unseal(state.dataKey, vault.roster).then(function (b) { state.roster = json(b); }); }
  function unlockMember(pw) {
    return kdf(pw, unb64(vault.saltM), vault.iter).then(function (k) { return unseal(k, vault.m); }).then(rawKey).then(function (k) { state.dataKey = k; return openRoster(); });
  }
  function unlockOfficer(pw) {
    return kdf(pw, unb64(vault.saltO), vault.iter).then(function (k) { return unseal(k, vault.o); }).then(function (b) {
      var o = json(b);
      return Promise.all([rawKey(unb64(o.dataKey)), rawKey(unb64(o.fileKey))]);
    }).then(function (ks) {
      state.dataKey = ks[0]; state.fileKey = ks[1];
      return Promise.all([openRoster(), unseal(state.fileKey, vault.files).then(function (b) { state.files = json(b); })]);
    });
  }

  /* ---------- 화면 도우미 ---------- */
  function h(tag, props, children) {
    var n = document.createElement(tag), k;
    if (props) for (k in props) { if (k === 'class') n.className = props[k]; else if (k === 'text') n.textContent = props[k]; else if (k.slice(0, 2) === 'on') n.addEventListener(k.slice(2), props[k]); else if (props[k] != null && props[k] !== false) n.setAttribute(k, props[k] === true ? '' : props[k]); }
    (children || []).forEach(function (c) { if (c != null) n.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
    return n;
  }
  function lockBox(id, label, hint, onTry) {
    var msg = h('p', { class: 'fmsg', role: 'status' });
    var input = h('input', { type: 'password', id: id, autocomplete: 'off' });
    var btn = h('button', { class: 'btn', type: 'submit', text: '열기' });
    var form = h('form', { class: 'lock', onsubmit: function (e) {
      e.preventDefault(); if (!input.value) return;
      btn.disabled = true; msg.className = 'fmsg'; msg.textContent = '확인하는 중입니다…';
      onTry(input.value).then(function () { input.value = ''; }, function () { msg.className = 'fmsg bad'; msg.textContent = '비밀번호가 맞지 않습니다.'; }).then(function () { btn.disabled = false; });
    } }, [
      h('div', { class: 'ffield' }, [h('label', { for: id, text: label }), input]),
      btn, h('p', { class: 'lock-hint', text: hint }), msg
    ]);
    return form;
  }

  /* ---------- 회원 명부 ---------- */
  function drawMembers() {
    var box = document.getElementById('s-members'); if (!box) return;
    box.textContent = '';
    if (!vault) { box.appendChild(h('p', { class: 'fempty', text: '아직 등록된 명부가 없습니다. 임원이 "명부·임원 자료 관리"에서 명부를 올리면 여기에 표시됩니다.' })); return; }
    if (!state.roster) {
      box.appendChild(lockBox('s-mpw', '회원 비밀번호', '회원 명부는 광양교육사랑 회원만 볼 수 있습니다. 비밀번호는 가입할 때 사무국에서 알려 드립니다.', function (pw) {
        return unlockMember(pw).catch(function () { return unlockOfficer(pw); }).then(function () { drawMembers(); drawOfficerFiles(); });
      }));
      return;
    }
    var r = state.roster, q = h('input', { type: 'search', id: 's-q', placeholder: '이름이나 업체로 찾기' }), tbody = h('tbody');
    function rows() {
      var k = q.value.trim().toLowerCase(); tbody.textContent = '';
      r.rows.filter(function (row) { return !k || row.join(' ').toLowerCase().indexOf(k) >= 0; }).forEach(function (row) {
        tbody.appendChild(h('tr', null, r.headers.map(function (_, i) { return h('td', { text: row[i] == null ? '' : String(row[i]) }); })));
      });
    }
    q.addEventListener('input', rows); rows();
    box.appendChild(h('div', { class: 'ftpl-h' }, [h('span', { class: 'pill', text: '회원 ' + r.rows.length + '명' }),
      h('div', { class: 'ffield', style: 'flex:1 1 220px;max-width:320px' }, [h('label', { for: 's-q', text: '찾기', style: 'position:absolute;width:1px;height:1px;overflow:hidden' }), q]),
      h('button', { type: 'button', class: 'flink', text: '잠그기', onclick: function () { state.roster = null; state.dataKey = null; state.files = null; state.fileKey = null; drawMembers(); drawOfficerFiles(); } })]));
    box.appendChild(h('div', { class: 'tblwrap' }, [h('table', { class: 'tbl' }, [h('thead', null, [h('tr', null, r.headers.map(function (x) { return h('th', { text: x }); }))]), tbody])]));
    box.appendChild(h('p', { class: 'lock-hint', text: '명부는 회원끼리 연락하는 데만 써 주세요. 화면을 찍거나 밖으로 전달하지 말아 주세요.' }));
  }

  /* ---------- 임원 전용 자료 ---------- */
  function drawOfficerFiles() {
    var box = document.getElementById('s-officer'); if (!box) return;
    box.textContent = '';
    if (!vault) { box.appendChild(h('p', { class: 'fempty', text: '아직 등록된 임원 전용 자료가 없습니다.' })); return; }
    if (!state.files) {
      box.appendChild(lockBox('s-opw', '임원 비밀번호', '시나리오, 감사장 원본처럼 임원진만 보는 자료입니다.', function (pw) { return unlockOfficer(pw).then(function () { drawMembers(); drawOfficerFiles(); }); }));
      return;
    }
    if (!state.files.length) box.appendChild(h('p', { class: 'fempty', text: '등록된 임원 전용 자료가 없습니다.' }));
    var list = h('div', { class: 'files' });
    state.files.forEach(function (f) {
      var status = h('small', { text: '' });
      var btn = h('button', { type: 'button', class: 'act', text: '내려받기', onclick: function () {
        status.textContent = ' 여는 중…';
        fetch('secure/' + f.id + '.json').then(function (r) { if (!r.ok) throw 0; return r.json(); }).then(function (box) { return unseal(state.fileKey, box); }).then(function (bytes) {
          return window.GyeduForms.save(f.name, new Blob([bytes]));
        }).then(function (res) {
          status.textContent = res === 'saved' ? ' 저장을 요청했습니다.' : res === 'declined' ? ' 저장을 취소했습니다.' : ' 이 미리보기 화면에서는 이 형식을 저장할 수 없습니다. 공개 주소에서는 저장됩니다.';
        }, function () { status.textContent = ' 파일을 열지 못했습니다.'; });
      } });
      list.appendChild(h('div', { class: 'file' }, [h('span', { class: 'pill gold', text: '임원 전용' }), h('span', { class: 'meta' }, [h('b', { text: f.name }), h('small', { text: Math.max(1, Math.round(f.size / 1024)) + 'KB' }), status]), btn]));
    });
    box.appendChild(list);
  }

  /* ---------- 관리: 보안 묶음 만들기 ---------- */
  function readRoster(file) {
    return file.arrayBuffer().then(function (buf) { return window.GyeduForms.openBook(buf); }).then(function (book) {
      var s = book.sheets[0], grid = {}, maxC = 0, rowNums = [];
      Object.keys(s.cells).forEach(function (a) {
        var m = /^([A-Z]+)(\d+)$/.exec(a), t = String(s.cells[a].text || '').trim(); if (!m || !t) return;
        var c = 0, i; for (i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) - 64);
        var r = +m[2]; if (!grid[r]) { grid[r] = {}; rowNums.push(r); } grid[r][c] = t; if (c > maxC) maxC = c;
      });
      rowNums.sort(function (a, b) { return a - b; });
      var hi = -1, i;
      for (i = 0; i < rowNums.length; i++) if (Object.keys(grid[rowNums[i]]).length >= 2) { hi = i; break; }
      if (hi < 0) throw new Error('empty');
      var cols = Object.keys(grid[rowNums[hi]]).map(Number).sort(function (a, b) { return a - b; });
      var headers = cols.map(function (c) { return grid[rowNums[hi]][c]; });
      var rows = rowNums.slice(hi + 1).map(function (r) { return cols.map(function (c) { return grid[r][c] || ''; }); }).filter(function (r) { return r.some(function (x) { return x; }); });
      return { headers: headers, rows: rows };
    });
  }
  function build(roster, files, mpw, opw) {
    var dataRaw = rand(32), fileRaw = rand(32), saltM = rand(16), saltO = rand(16), zip = new JSZip(), list = [];
    return Promise.all([kdf(mpw, saltM, ITER), kdf(opw, saltO, ITER), rawKey(dataRaw), rawKey(fileRaw)]).then(function (k) {
      var chain = Promise.resolve();
      files.forEach(function (f, i) {
        chain = chain.then(function () { return f.arrayBuffer(); }).then(function (buf) { return seal(k[3], buf); }).then(function (box) {
          var id = 'f' + (i + 1) + '-' + b64(rand(6)).replace(/[^a-zA-Z0-9]/g, '').slice(0, 6);
          zip.file('secure/' + id + '.json', JSON.stringify(box)); list.push({ id: id, name: f.name, size: f.size });
        });
      });
      return chain.then(function () {
        return Promise.all([seal(k[0], dataRaw), seal(k[1], te.encode(JSON.stringify({ dataKey: b64(dataRaw), fileKey: b64(fileRaw) }))), seal(k[2], te.encode(JSON.stringify(roster))), seal(k[3], te.encode(JSON.stringify(list)))]);
      });
    }).then(function (b) {
      zip.file('secure/vault.json', JSON.stringify({ v: 1, iter: ITER, saltM: b64(saltM), saltO: b64(saltO), m: b[0], o: b[1], roster: b[2], files: b[3] }));
      return zip.generateAsync({ type: 'blob' });
    });
  }
  function initManage() {
    var form = document.getElementById('s-build'); if (!form) return;
    var msg = document.getElementById('s-bmsg');
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var rf = document.getElementById('s-roster').files[0], fs = Array.prototype.slice.call(document.getElementById('s-files').files);
      var mpw = document.getElementById('s-newm').value, opw = document.getElementById('s-newo').value;
      function bad(t) { msg.className = 'fmsg bad'; msg.textContent = t; }
      if (!rf) return bad('회원 명부 엑셀(.xlsx)을 골라 주세요.');
      if (mpw.length < 8 || opw.length < 8) return bad('비밀번호는 각각 8자 이상으로 정해 주세요.');
      if (mpw === opw) return bad('회원 비밀번호와 임원 비밀번호는 서로 다르게 정해 주세요.');
      msg.className = 'fmsg'; msg.textContent = '만드는 중입니다…';
      readRoster(rf).then(function (roster) { return build(roster, fs, mpw, opw).then(function (blob) { return { blob: blob, n: roster.rows.length }; }); }).then(function (r) {
        var out = document.getElementById('s-bout'); out.textContent = '';
        var status = h('small', { text: '' });
        out.appendChild(h('div', { class: 'file' }, [h('span', { class: 'meta' }, [h('b', { text: 'secure.zip' }), h('small', { text: '회원 ' + r.n + '명 · 임원 전용 자료 ' + fs.length + '개' }), status]),
          h('button', { type: 'button', class: 'act', text: '내려받기', onclick: function () { window.GyeduForms.save('secure.zip', r.blob).then(function (res) { status.textContent = res === 'saved' ? ' 저장을 요청했습니다.' : ' 저장하지 못했습니다.'; }); } })]));
        msg.textContent = '보안 묶음을 만들었습니다. 내려받아 압축을 풀고, secure 폴더를 홈페이지 파일이 있는 곳에 올려 주세요.';
      }, function () { bad('명부를 읽지 못했습니다. 첫 줄에 "이름, 연락처" 같은 제목이 있는 .xlsx 파일인지 확인해 주세요.'); });
    });
  }

  function init() { initManage(); loadVault().then(function () { drawMembers(); drawOfficerFiles(); }); }
  window.GyeduSecure = { build: build, readRoster: readRoster };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
