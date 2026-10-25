(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var S = { file: '', aoa: null, recs: [], names: [], open: {}, pay: loadPay() };

  /* 薪資設定記在本機，下次不用重打 */
  function loadPay() { try { return JSON.parse(localStorage.getItem('wh-pay-v1')) || {}; } catch (e) { return {}; } }
  function savePay() { try { localStorage.setItem('wh-pay-v1', JSON.stringify(S.pay)); } catch (e) {} }

  var money = function (n) { return '$' + Math.round(n).toLocaleString('en-US'); };
  var fix = function (n) { return (Math.round(n * 10) / 10).toFixed(1); };
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };

  /* ---------- 讀檔 ---------- */
  function showErr(msg) { var e = $('err'); e.textContent = msg; e.hidden = !msg; }

  function handleFile(f) {
    if (!f) return;
    showErr('');
    var rd = new FileReader();
    rd.onload = function () {
      try {
        var wb = XLSX.read(new Uint8Array(rd.result), { type: 'array', raw: true });
        var aoa = null;
        for (var i = 0; i < wb.SheetNames.length && !aoa; i++) {
          var a = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[i]], { header: 1, raw: true, defval: null });
          try { WH.parseRows(a); aoa = a; } catch (e) { /* 換下一張工作表 */ }
        }
        if (!aoa) throw new Error('找不到「員工、上班時間、下班時間」欄位，請確認是打卡系統匯出的原始表格。');
        load(aoa, f.name);
      } catch (e) { showErr(e.message || '無法讀取這個檔案'); }
    };
    rd.onerror = function () { showErr('無法讀取這個檔案'); };
    rd.readAsArrayBuffer(f);
  }

  function load(aoa, name) {
    var rows = WH.parseRows(aoa);
    if (!rows.length) throw new Error('檔案裡沒有任何打卡紀錄。');
    var b = WH.buildRecords(rows);
    S.aoa = aoa; S.file = name; S.recs = b.recs; S.names = b.names; S.open = {};
    b.names.forEach(function (n) { if (b.recs.some(function (r) { return r.name === n && WH.status(r); })) S.open[n] = true; });
    b.names.forEach(function (n) { if (!S.pay[n]) S.pay[n] = { type: '時薪', amt: '' }; });
    $('home').hidden = true; $('result').hidden = false; $('bar').hidden = false;
    render();
    window.scrollTo(0, 0);
  }

  /* ---------- 計算 ---------- */
  function totals() {
    var t = {};
    S.names.forEach(function (n) { t[n] = { h: 0, miss: 0 }; });
    S.recs.forEach(function (r) {
      var c = WH.calc(r.inAbs, r.outAbs);
      if (c) t[r.name].h += c.net; else t[r.name].miss++;
    });
    return t;
  }
  function payOf(n, h) {
    var p = S.pay[n] || {}, a = parseFloat(p.amt);
    if (!(a > 0)) return 0;
    return p.type === '月薪' ? a : a * h;
  }
  function period() {
    var ts = S.recs.map(function (r) { return r.inAbs != null ? r.inAbs : r.outAbs; });
    var a = Math.min.apply(null, ts), b = Math.max.apply(null, ts);
    return { a: WH.absToDate(a), b: WH.absToDate(b) };
  }

  /* ---------- 畫面 ---------- */
  function render() {
    var t = totals(), p = period(), allH = 0, allM = 0, miss = 0;
    S.names.forEach(function (n) { allH += t[n].h; allM += payOf(n, t[n].h); miss += t[n].miss; });
    $('period').textContent = p.a.slice(5).replace('-', '/') + ' – ' + p.b.slice(5).replace('-', '/') + ' 工時';
    $('fname').textContent = S.file;
    $('s1').className = ''; $('s2').className = 'on'; $('s3').className = '';
    $('sum').innerHTML =
      '<div><span class="muted">人數</span><b>' + S.names.length + '</b></div>' +
      '<div><span class="muted">總時數（已扣休息）</span><b>' + fix(allH) + '</b></div>' +
      '<div class="' + (miss ? 'bad' : '') + '"><span class="muted">缺打卡待補</span><b>' + miss + '</b></div>';
    $('t-n').textContent = S.names.length; $('t-h').textContent = fix(allH); $('t-m').textContent = money(allM);

    var h = '<thead><tr><th>員工</th><th class="num">時數</th><th>狀態</th><th>薪資類型</th><th class="num">金額</th><th class="num">應付</th><th></th></tr></thead><tbody>';
    S.names.forEach(function (n) {
      var p2 = S.pay[n], pay = payOf(n, t[n].h), open = !!S.open[n];
      h += '<tr class="emp" data-n="' + esc(n) + '"><td class="name">' + esc(n) + '</td>' +
        '<td class="num" data-h>' + fix(t[n].h) + '</td>' +
        '<td data-st>' + (t[n].miss ? '<button class="tag" data-act="open">缺 ' + t[n].miss + ' 筆</button>' : '<span class="muted">正常</span>') + '</td>' +
        '<td><div class="seg" role="group" aria-label="薪資類型"><button type="button" data-act="type" data-v="月薪" aria-pressed="' + (p2.type === '月薪') + '">月薪</button><button type="button" data-act="type" data-v="時薪" aria-pressed="' + (p2.type === '時薪') + '">時薪</button></div></td>' +
        '<td class="num"><label class="money">$<input inputmode="decimal" data-act="amt" aria-label="' + esc(n) + ' 金額" value="' + esc(p2.amt) + '" placeholder="' + (p2.type === '月薪' ? '月薪' : '時薪') + '"></label></td>' +
        '<td class="num" data-pay>' + (pay ? money(pay) : '–') + '</td>' +
        '<td><button class="btn link" data-act="open">' + (open ? '收合明細' : '看明細') + '</button></td></tr>';
      if (open) {
        h += '<tr class="sub"><td colspan="7" style="padding:0"><div class="scroll"><table><thead><tr><th>日期</th><th>週</th><th>上班</th><th>下班</th><th class="num">換算</th><th class="num">休息</th><th class="num">計算時數</th><th>備註</th></tr></thead><tbody>';
        S.recs.filter(function (r) { return r.name === n; }).forEach(function (r) {
          var c = WH.calc(r.inAbs, r.outAbs), st = WH.status(r), ref = r.inAbs != null ? r.inAbs : r.outAbs;
          h += '<tr class="' + (st ? 'miss' : '') + '" data-id="' + r.id + '"><td>' + WH.absToDate(ref).slice(5) + '</td><td>' + WH.weekday(ref) + '</td>' +
            '<td><input type="text" data-act="time" data-k="in" value="' + (r.inAbs != null ? WH.absToHM(r.inAbs) : '') + '" placeholder="hh:mm" aria-label="上班時間"></td>' +
            '<td><input type="text" data-act="time" data-k="out" value="' + (r.outAbs != null ? WH.absToHM(r.outAbs) : '') + '" placeholder="hh:mm" aria-label="下班時間"></td>' +
            '<td class="num">' + (c ? fix(c.gross) : '') + '</td><td class="num">' + (c ? fix(c.brk) : '') + '</td><td class="num"><b>' + (c ? fix(c.net) : '–') + '</b></td>' +
            '<td class="note ' + (st ? '' : 'ok') + '">' + esc(st || r.note) + '</td></tr>';
        });
        h += '</tbody></table></div></td></tr>';
      }
    });
    $('tbl').innerHTML = h + '</tbody>';
  }

  /* 只更新數字，輸入時不重繪整張表，游標才不會跳掉 */
  function refreshNumbers() {
    var t = totals(), allH = 0, allM = 0;
    S.names.forEach(function (n) { allH += t[n].h; allM += payOf(n, t[n].h); });
    $('t-h').textContent = fix(allH); $('t-m').textContent = money(allM);
    Array.prototype.forEach.call(document.querySelectorAll('tr.emp'), function (tr) {
      var n = tr.getAttribute('data-n'), pay = payOf(n, t[n].h);
      tr.querySelector('[data-h]').textContent = fix(t[n].h);
      tr.querySelector('[data-pay]').textContent = pay ? money(pay) : '–';
    });
  }

  $('tbl').addEventListener('click', function (e) {
    var b = e.target.closest('[data-act]'); if (!b) return;
    var tr = b.closest('tr'), n = (tr.classList.contains('emp') ? tr : tr.closest('tr.sub').previousElementSibling).getAttribute('data-n');
    var act = b.getAttribute('data-act');
    if (act === 'open') { S.open[n] = !S.open[n]; render(); }
    else if (act === 'type') { S.pay[n].type = b.getAttribute('data-v'); savePay(); render(); }
  });
  $('tbl').addEventListener('input', function (e) {
    var el = e.target;
    if (el.getAttribute('data-act') === 'amt') {
      var n = el.closest('tr').getAttribute('data-n');
      S.pay[n].amt = el.value.replace(/[^\d.]/g, ''); savePay(); refreshNumbers();
    }
  });
  $('tbl').addEventListener('change', function (e) {
    var el = e.target;
    if (el.getAttribute('data-act') !== 'time') return;
    var r = S.recs[+el.closest('tr').getAttribute('data-id')], k = el.getAttribute('data-k');
    var m = el.value.trim().replace('：', ':').match(/^(\d{1,2}):?(\d{2})$/);
    var day = Math.floor((r.inAbs != null ? r.inAbs : r.outAbs) / 1440) * 1440;
    if (!el.value.trim()) { if (k === 'in') r.inAbs = null; else r.outAbs = null; }
    else if (m && +m[1] < 24 && +m[2] < 60) {
      var v = day + (+m[1]) * 60 + (+m[2]);
      if (k === 'in') r.inAbs = v; else { r.outAbs = v; if (r.inAbs != null && v <= r.inAbs) r.outAbs = v + 1440; }
    }
    if (r.inAbs == null && r.outAbs == null) r.outAbs = day; /* 保底 */
    r.note = '已手動補時間';
    render();
  });

  /* ---------- 檔案進來 ---------- */
  $('pick').onclick = function () { $('file').click(); };
  $('file').onchange = function () { handleFile(this.files[0]); this.value = ''; };
  var drop = $('home');
  ['dragenter', 'dragover'].forEach(function (ev) { document.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { document.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
  document.addEventListener('drop', function (e) { if (e.dataTransfer && e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });
  $('again').onclick = function () { $('home').hidden = false; $('result').hidden = true; $('bar').hidden = true; $('s1').className = 'on'; $('s2').className = ''; $('s3').className = ''; };

  $('demo').onclick = function () {
    var D = [['打卡紀錄 : 範例資料'], ['員工', '周', '上班時間', '下班時間', '工時']];
    var days = ['2026-09-01', '2026-09-02', '2026-09-03', '2026-09-04', '2026-09-07'];
    days.forEach(function (d, i) { D.push(['王小美', '', d + ' 0' + (8 + (i % 2)) + ':5' + (5 + i), d + ' 16:3' + (3 + i), '']); });
    days.forEach(function (d, i) { D.push(['陳大明', '', d + ' 13:0' + i, d + (i === 2 ? ' 22:21' : ' 21:3' + i), '']); });
    D.push(['陳大明', '', '2026-09-08 17:02', null, '']);
    load(D, '範例資料（非真實）');
  };

  /* ---------- 匯出 Excel：保留公式，可直接在 Excel 繼續編輯 ---------- */
  var ORANGE = 'FFCF5E06', SOFT = 'FFFFEBDA', RED = 'FFFDE9E6';
  function head(row) { row.eachCell(function (c) { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ORANGE } }; c.alignment = { vertical: 'middle', horizontal: 'center' }; }); row.height = 24; }

  $('export').onclick = function () {
    var wb = new ExcelJS.Workbook(), t = totals(), p = period();
    var sm = wb.addWorksheet('薪資總表'), dt = wb.addWorksheet('每日明細'), raw = wb.addWorksheet('原始資料');

    /* 每日明細：換算小時可手改，休息與計算時數是公式 */
    dt.columns = [{ width: 14 }, { width: 12 }, { width: 6 }, { width: 10 }, { width: 10 }, { width: 11 }, { width: 11 }, { width: 14 }, { width: 28 }];
    head(dt.addRow(['員工', '日期', '週', '上班', '下班', '換算小時', '休息', '計算時數', '備註']));
    S.recs.forEach(function (r, i) {
      var c = WH.calc(r.inAbs, r.outAbs), st = WH.status(r), ref = r.inAbs != null ? r.inAbs : r.outAbs, n = i + 2;
      var row = dt.addRow([r.name, WH.absToDate(ref), WH.weekday(ref),
        r.inAbs != null ? WH.absToHM(r.inAbs) : '', r.outAbs != null ? WH.absToHM(r.outAbs) : '',
        c ? c.gross : null,
        { formula: 'IF(F' + n + '="","",IF(F' + n + '>=8,1,IF(F' + n + '>=4,0.5,0)))', result: c ? c.brk : '' },
        { formula: 'IF(F' + n + '="","",F' + n + '-G' + n + ')', result: c ? c.net : '' },
        st || r.note]);
      ['F', 'G', 'H'].forEach(function (L) { row.getCell(L).numFmt = '0.0'; row.getCell(L).alignment = { horizontal: 'right' }; });
      if (st) row.eachCell({ includeEmpty: true }, function (cell) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } }; });
    });
    dt.views = [{ state: 'frozen', ySplit: 1 }];
    dt.autoFilter = { from: 'A1', to: 'I' + (S.recs.length + 1) };

    /* 薪資總表：時數用 SUMIF 抓明細，改明細會跟著算 */
    sm.columns = [{ width: 16 }, { width: 10 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 14 }];
    sm.mergeCells('A1:F1'); sm.getCell('A1').value = '工時薪資總表　' + p.a + ' ～ ' + p.b;
    sm.getCell('A1').font = { bold: true, size: 15, color: { argb: ORANGE } }; sm.getRow(1).height = 28;
    sm.mergeCells('A2:F2'); sm.getCell('A2').value = '橘色底的格子可以直接改：類型、月薪、時薪。時數來自「每日明細」。';
    sm.getCell('A2').font = { color: { argb: 'FF75655A' }, size: 10 };
    head(sm.addRow(['員工', '類型', '月薪', '時薪', '總時數', '應付金額']));
    var total = 0, first = 4;
    S.names.forEach(function (nm, i) {
      var n = i + first, pay = S.pay[nm] || { type: '時薪', amt: '' }, a = parseFloat(pay.amt) || null;
      var res = payOf(nm, t[nm].h);
      var row = sm.addRow([nm, pay.type, pay.type === '月薪' ? a : null, pay.type === '時薪' ? a : null,
        { formula: 'SUMIF(每日明細!$A:$A,A' + n + ',每日明細!$H:$H)', result: Math.round(t[nm].h * 10) / 10 },
        { formula: 'IF(B' + n + '="月薪",N(C' + n + '),E' + n + '*N(D' + n + '))', result: res }]);
      total += res;
      row.getCell('B').dataValidation = { type: 'list', allowBlank: true, formulae: ['"月薪,時薪"'] };
      ['B', 'C', 'D'].forEach(function (L) { row.getCell(L).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SOFT } }; });
      row.getCell('B').alignment = { horizontal: 'center' };
      row.getCell('C').numFmt = row.getCell('D').numFmt = row.getCell('F').numFmt = '#,##0';
      row.getCell('E').numFmt = '0.0';
    });
    var last = first + S.names.length - 1, tr = sm.addRow(['合計', '', '', '',
      { formula: 'SUM(E' + first + ':E' + last + ')', result: Math.round(S.names.reduce(function (s, n) { return s + t[n].h; }, 0) * 10) / 10 },
      { formula: 'SUM(F' + first + ':F' + last + ')', result: total }]);
    tr.font = { bold: true }; tr.getCell('E').numFmt = '0.0'; tr.getCell('F').numFmt = '#,##0';
    tr.eachCell({ includeEmpty: true }, function (c) { c.border = { top: { style: 'medium', color: { argb: ORANGE } } }; });
    sm.addRow([]);
    ['計算規則：上班就近進位到整點／半點；下班往前扣到整點／半點，:20～:29、:50～:59 視為 :20、:50（尾數 .3、.8）；滿 4 小時休息 0.5，滿 8 小時休息 1。',
      '紅底列是缺打卡，請在「每日明細」補上時數（換算小時欄），休息與計算時數會自動更新。'].forEach(function (s) {
      var r = sm.addRow([s]); r.getCell(1).font = { color: { argb: 'FF75655A' }, size: 10 };
    });
    sm.views = [{ state: 'frozen', ySplit: 3 }];

    /* 原始資料：保留匯入的內容 */
    S.aoa.forEach(function (r) { raw.addRow(r.map(function (v) { return v == null ? '' : v; })); });
    raw.columns.forEach(function (c) { c.width = 18; });
    head(raw.getRow(2));

    wb.xlsx.writeBuffer().then(function (buf) {
      var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = '工時薪資_' + p.a.slice(0, 7) + '.xlsx';
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 2000);
      $('s3').className = 'on';
    });
  };
})();
