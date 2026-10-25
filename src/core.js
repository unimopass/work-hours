/* 工時計算核心規則（純函式，無畫面相依） */
(function (root) {
  'use strict';
  var WD = ['日', '一', '二', '三', '四', '五', '六'];
  var p2 = function (n) { return String(n).padStart(2, '0'); };
  var r1 = function (x) { return Math.round(x * 10) / 10; };

  /* 讀取時間：支援 "2026-09-01 08:58" 字串、Excel 序號、Date。abs = 自 1970 起的分鐘數 */
  function parseStamp(v) {
    if (v == null || v === '') return null;
    var y, mo, d, h, mi, dt;
    if (v instanceof Date) {
      y = v.getFullYear(); mo = v.getMonth() + 1; d = v.getDate(); h = v.getHours(); mi = v.getMinutes();
    } else if (typeof v === 'number') {
      dt = new Date(Math.round((v - 25569) * 1440) * 60000);
      y = dt.getUTCFullYear(); mo = dt.getUTCMonth() + 1; d = dt.getUTCDate(); h = dt.getUTCHours(); mi = dt.getUTCMinutes();
    } else {
      var m = String(v).trim().match(/(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})[T\s]+(\d{1,2}):(\d{2})/);
      if (!m) return null;
      y = +m[1]; mo = +m[2]; d = +m[3]; h = +m[4]; mi = +m[5];
    }
    return { abs: Date.UTC(y, mo - 1, d, h, mi) / 60000 };
  }

  function absToHM(abs) { var m = ((abs % 1440) + 1440) % 1440; return p2(Math.floor(m / 60)) + ':' + p2(m % 60); }
  function absToDate(abs) { var d = new Date(abs * 60000); return d.getUTCFullYear() + '-' + p2(d.getUTCMonth() + 1) + '-' + p2(d.getUTCDate()); }
  function weekday(abs) { return WD[new Date(abs * 60000).getUTCDay()]; }

  /* 規則 1：上班就近進位到 :00 / :30（08:58 → 09:00，09:12 → 09:00，09:22 → 09:30）
     規則 2：下班一律往前扣到 :00 / :30（16:37 → 16:30，16:52 → 16:30）
     規則 3：下班落在 :20～:29 視為 :20、落在 :50～:59 視為 :50（小時尾數 .3 / .8）；到 :30 / :00 就回到規則 2 */
  function snap(abs, isOut) {
    var m = ((abs % 60) + 60) % 60, base = abs - m;
    if (isOut) {
      if (m >= 50) return base + 50;
      if (m >= 30) return base + 30;
      if (m >= 20) return base + 20;
      return base;
    }
    if (m < 15) return base;
    if (m < 45) return base + 30;
    return base + 60;
  }

  /* 規則 2：滿 4 小時休息 0.5、滿 8 小時休息 1 */
  function breakFor(gross) { return gross >= 8 ? 1 : gross >= 4 ? 0.5 : 0; }

  function calc(inAbs, outAbs) {
    if (inAbs == null || outAbs == null) return null;
    var mins = snap(outAbs, true) - snap(inAbs, false);
    if (mins <= 0) return null;
    var gross = r1(mins / 60), brk = breakFor(gross);
    return { gross: gross, brk: brk, net: r1(gross - brk) };
  }

  /* 找到表頭並取出每一列打卡 */
  function parseRows(aoa) {
    var hi = -1, ci = {}, i, j;
    for (i = 0; i < Math.min(aoa.length, 30); i++) {
      var row = aoa[i] || [], c = {};
      for (j = 0; j < row.length; j++) {
        var t = String(row[j] == null ? '' : row[j]).trim();
        if (t === '員工') c.name = j;
        else if (t === '上班時間') c.inn = j;
        else if (t === '下班時間') c.out = j;
        else if (t === '工時') c.orig = j;
      }
      if (c.name != null && c.inn != null && c.out != null) { hi = i; ci = c; break; }
    }
    if (hi < 0) throw new Error('找不到「員工、上班時間、下班時間」欄位，請確認是打卡系統匯出的原始表格。');
    var rows = [];
    for (i = hi + 1; i < aoa.length; i++) {
      var r = aoa[i] || [];
      var name = String(r[ci.name] == null ? '' : r[ci.name]).trim();
      if (!name) continue; /* 空白列、「總工時」小計列 */
      var a = parseStamp(r[ci.inn]), b = parseStamp(r[ci.out]);
      if (!a && !b) continue;
      rows.push({ name: name, in: a, out: b });
    }
    return rows;
  }

  /* 同一人同一天合併：完整的一筆各自成列；單邊打卡若落在完整班內視為重複；
     剩下的「只有上班＋只有下班」配成一班；仍無法配對的標記為缺打卡 */
  function buildRecords(rows) {
    var groups = new Map(), order = [];
    rows.forEach(function (r) {
      var ref = (r.in || r.out).abs, key = r.name + '|' + absToDate(ref);
      if (!groups.has(key)) groups.set(key, { name: r.name, rows: [] });
      groups.get(key).rows.push(r);
      if (order.indexOf(r.name) < 0) order.push(r.name);
    });
    var recs = [];
    groups.forEach(function (g) {
      /* 上下班只差幾分鐘的「班」是重複刷卡，當作單邊打卡處理 */
      var comp = g.rows.filter(function (r) { return r.in && r.out && r.out.abs - r.in.abs > 5; });
      var shorts = g.rows.filter(function (r) { return r.in && r.out && r.out.abs - r.in.abs <= 5; });
      var ins = g.rows.filter(function (r) { return r.in && !r.out; }).map(function (r) { return r.in.abs; });
      var outs = g.rows.filter(function (r) { return !r.in && r.out; }).map(function (r) { return r.out.abs; });
      shorts.forEach(function (r) { outs.push(r.out.abs); });
      var segs = comp.map(function (r) { return [r.in.abs, r.out.abs]; });
      /* 完整班前後 15 分鐘內的單邊打卡，視為同一班的重複刷卡 */
      var absorbed = function (t) { return segs.some(function (s) { return t >= s[0] - 15 && t <= s[1] + 15; }); };
      var li = ins.filter(function (t) { return !absorbed(t); });
      var lo = outs.filter(function (t) { return !absorbed(t); });
      segs.forEach(function (s) { recs.push({ name: g.name, inAbs: s[0], outAbs: s[1], note: '' }); });
      if (!segs.length && li.length && lo.length && Math.min.apply(null, li) < Math.max.apply(null, lo)) {
        recs.push({ name: g.name, inAbs: Math.min.apply(null, li), outAbs: Math.max.apply(null, lo), note: '同日上下班合併' });
      } else {
        if (li.length) recs.push({ name: g.name, inAbs: Math.min.apply(null, li), outAbs: null, note: li.length > 1 ? '同日有 ' + li.length + ' 筆上班打卡' : '' });
        if (lo.length) recs.push({ name: g.name, inAbs: null, outAbs: Math.max.apply(null, lo), note: lo.length > 1 ? '同日有 ' + lo.length + ' 筆下班打卡，取最晚' : '' });
      }
    });
    recs.sort(function (a, b) {
      return (order.indexOf(a.name) - order.indexOf(b.name)) || ((a.inAbs != null ? a.inAbs : a.outAbs) - (b.inAbs != null ? b.inAbs : b.outAbs));
    });
    recs.forEach(function (r, k) { r.id = k; });
    return { recs: recs, names: order };
  }

  function status(rec) {
    if (rec.inAbs == null) return '缺上班打卡，請補';
    if (rec.outAbs == null) return '缺下班打卡，請補';
    if (!calc(rec.inAbs, rec.outAbs)) return '時間異常，請檢查';
    return '';
  }

  var api = { parseStamp: parseStamp, absToHM: absToHM, absToDate: absToDate, weekday: weekday,
    snap: snap, breakFor: breakFor, calc: calc, parseRows: parseRows, buildRecords: buildRecords, status: status, r1: r1 };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.WH = api;
})(typeof window !== 'undefined' ? window : globalThis);
