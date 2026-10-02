/* Goal Board data processing — shared by the page and the node tests. Needs global XLSX (SheetJS). */
var GSProc = (function () {
  var MAIN = {
    PRGL: { w: 'Paragon Lady', t: 'L', s: 'BKK', tier: 'A' }, ZENL: { w: 'Zen Lady', t: 'L', s: 'BKK', tier: 'B' }, STKL: { w: 'Siam Takashimaya Lady', t: 'L', s: 'BKK', tier: 'C' },
    PRGM: { w: 'Paragon Men', t: 'M', s: 'BKK', tier: 'C' }, CHDM: { w: 'Central Chidlom Men', t: 'M', s: 'BKK', tier: 'B' }, ZENM: { w: 'Zen Men', t: 'M', s: 'BKK', tier: 'B' },
    GSVS: { w: 'Gaysorn Shop', t: 'S', s: 'BKK', tier: 'A' }, ICDS: { w: 'Icon Siam Shop', t: 'S', s: 'BKK', tier: 'A' }, TMNS: { w: 'Terminal 21 Shop', t: 'S', s: 'BKK', tier: 'A' }, RGRS: { w: 'Riverside Plaza Shop', t: 'S', s: 'BKK', tier: 'C' },
    PKCL: { w: 'Phuket Central Lady', t: 'L', s: 'UPC', tier: 'B' }, PHCL: { w: 'Patong Central Lady', t: 'L', s: 'UPC', tier: 'C' }, SMUL: { w: 'Samui Lady', t: 'L', s: 'UPC', tier: 'C' },
    PKCM: { w: 'Phuket Central Men', t: 'M', s: 'UPC', tier: 'C' }, PHCM: { w: 'Patong Central Men', t: 'M', s: 'UPC', tier: 'C' },
    JCPS: { w: 'Jungceylon Plaza Shop', t: 'S', s: 'UPC', tier: 'C' }, SMAS: { w: 'Samui Airport Shop', t: 'S', s: 'UPC', tier: 'A' }
  };
  var POOL_EXTRA = ['Phuket Central Event', 'Jungceylon Plaza Event', 'Terminal Rama 3 Event'];
  var TIER_ORDER = { C: 0, B: 1, A: 2 };
  var TH_MONTH = { 'มกราคม': 1, 'กุมภาพันธ์': 2, 'มีนาคม': 3, 'เมษายน': 4, 'พฤษภาคม': 5, 'มิถุนายน': 6, 'กรกฎาคม': 7, 'สิงหาคม': 8, 'กันยายน': 9, 'ตุลาคม': 10, 'พฤศจิกายน': 11, 'ธันวาคม': 12 };

  function norm(s) { return String(s == null ? '' : s).trim(); }
  function pad(n) { return String(n).padStart(2, '0'); }
  function ymd(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function toDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date && !isNaN(v)) return v;
    if (typeof v === 'number') { var p = XLSX.SSF.parse_date_code(v); return p ? new Date(p.y, p.m - 1, p.d) : null; }
    var s = String(v).trim(), m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); // M/D/YYYY (Google Sheets US locale)
    if (m) { var y = +m[3]; if (y > 2400) y -= 543; return new Date(y, +m[1] - 1, +m[2]); }
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    return null;
  }
  function firstNum(v) { if (typeof v === 'number') return v; var m = String(v == null ? '' : v).replace(/,/g, '').match(/-?\d+(\.\d+)?/); return m ? +m[0] : 0; }

  // Who counts for which board: BKK = คุณเอ + คุณโอ๋, UPC = คุณโบว์ (written คุณโบ in the sheet)
  function scopeOfPerson(name) {
    var n = norm(name).replace(/^คุณ/, '').replace(/\s+/g, '');
    if (!n) return null;
    if (n === 'เอ' || n === 'โอ๋' || n === 'โอ') return 'BKK';
    if (n === 'โบ' || n === 'โบว์' || n === 'โบ้' || n === 'มิ้ว' || n === 'มิว' || n === 'มิ๊ว' || n === 'มิ้วส์') return 'UPC';
    return null;
  }

  // one spelling per Sup on every record
  function supName(name) { var n = norm(name).replace(/^คุณ/, '').replace(/\s+/g, ''); return n === 'โบ' || n === 'โบว์' || n === 'โบ้' ? 'คุณโบว์' : norm(name); }

  // Free-text branch names in the evaluation sheet -> branch codes
  function branchCode(raw) {
    var s = norm(raw).toLowerCase(); if (!s) return null;
    var has = function (k) { return s.indexOf(k) >= 0; };
    var ev = has('event') || has('อีเว้น') || has('fair');
    if (has('paragon') || has('พารากอน')) return has('men') ? 'PRGM' : 'PRGL';
    if (has('zen') || has('เซ็น')) return ev ? 'ZENE' : (has('men') ? 'ZENM' : 'ZENL');
    if (has('chidlom') || has('ชิดลม')) return 'CHDM';
    if (has('taka') || has('ทาคา')) return 'STKL';
    if (has('gaysorn') || has('gasorn') || has('gatorn') || has('เกษร')) return 'GSVS';
    if (has('icon') || has('ไอคอน')) return ev ? 'ICDE' : 'ICDS';
    if (has('rama') || has('พระราม')) return ev ? 'TMTE' : 'TMTS';
    if (has('terminal') || has('asok') || has('อโศก')) return ev ? 'TMNE' : 'TMNS';
    if (has('riverside') || has('ริเวอร์')) return 'RGRS';
    if (has('emporium')) return has('men') ? 'EMPM' : 'EMPL';
    if (has('patong') || has('pathong') || has('ป่าตอง')) return has('men') ? 'PHCM' : 'PHCL';
    if (has('jungceylon') || has('jungcylon') || has('จังซีลอน')) return ev ? 'JCPE' : 'JCPS';
    if (has('samui') || has('สมุย')) return has('airport') || has('สนามบิน') ? 'SMAS' : 'SMUL';
    if (has('phuket') || has('ภูเก็ต') || has('floresta') || has('foresta')) return ev ? 'PKCE' : (has('men') ? 'PKCM' : 'PKCL');
    return null;
  }
  var BRANCH_SCOPE = { ZENE: 'BKK', ICDE: 'BKK', TMTE: 'BKK', TMTS: 'BKK', TMNE: 'BKK', EMPL: 'BKK', EMPM: 'BKK', PKCE: 'UPC', JCPE: 'UPC' };
  function scopeOfBranch(code) { return code ? (MAIN[code] ? MAIN[code].s : BRANCH_SCOPE[code] || null) : null; }

  function sheetRows(wb, mustHave) {
    var best = null;
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[i]], { header: 1, raw: true, defval: null });
      for (var r = 0; r < Math.min(rows.length, 5); r++) {
        var h = (rows[r] || []).map(norm);
        if (mustHave.every(function (k) { return h.some(function (x) { return x.indexOf(k) >= 0; }); })) {
          var cand = { name: wb.SheetNames[i], header: h, rows: rows.slice(r + 1), hr: r };
          if (!best || cand.rows.length > best.rows.length) best = cand; // a manual/guide sheet can repeat the headers; the data sheet has the most rows
          break;
        }
      }
    }
    return best;
  }
  function focusFromWorkbook(wb) {
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[i]], { header: 1, raw: true, defval: null });
      for (var r = 0; r < Math.min(rows.length, 10); r++) {
        var h = (rows[r] || []).map(norm), cm = h.indexOf('Model'), cl = h.indexOf('Lady'), cmen = h.indexOf('Men'), cs = h.indexOf('Shop');
        if (cm < 0 || cl < 0 || cmen < 0 || cs < 0) continue;
        var out = [], grp = null;
        rows.slice(r + 1).forEach(function (x) {
          if (!x) return;
          if (cm >= 2 && x[cm - 2]) grp = norm(x[cm - 2]);
          var no = x[cm - 1], name = norm(x[cm]);
          if (no == null || no === '' || !name || /TOTAL/i.test(name)) return;
          var yes = function (v) { return /^y/i.test(norm(v)); };
          out.push({ no: no, group: grp, name: name, L: yes(x[cl]), M: yes(x[cmen]), S: yes(x[cs]) });
        });
        return out;
      }
    }
    throw new Error('ไม่พบหัวคอลัมน์ Model / Lady / Men / Shop ในไฟล์ Focus Product');
  }
  function col(h, test) { for (var i = 0; i < h.length; i++) if (test(h[i])) return i; return -1; }

  /* ---------- Evaluation sheet: audits + visits + exam ---------- */
  function parseAudit(wb) {
    var t = sheetRows(wb, ['SC ที่ถูกประเมิน', 'ผู้ประเมิน']);
    if (!t) throw new Error('ไม่พบชีทที่มีหัวคอลัมน์ "SC ที่ถูกประเมิน" และ "ผู้ประเมิน"');
    var h = t.header;
    var c = {
      ts: col(h, function (x) { return x === 'Timestamp'; }),
      date: col(h, function (x) { return x.indexOf('วันที่เข้าประเมิน') >= 0; }),
      year: col(h, function (x) { return x.indexOf('ประจำปี') >= 0; }),
      month: col(h, function (x) { return x.indexOf('ประจำเดือน') >= 0; }),
      method: col(h, function (x) { return x.indexOf('วิธีการประเมิน') >= 0; }),
      sc: col(h, function (x) { return x.indexOf('SC ที่ถูกประเมิน') >= 0; }),
      br: col(h, function (x) { return x.indexOf('สาขาที่เข้าไปประเมิน') >= 0; }),
      ev: col(h, function (x) { return x.indexOf('ผู้ประเมิน') >= 0; })
    };
    var scoreCols = [], examCols = [], noteCols = [], items = [], noteFrom = [];
    h.forEach(function (x, i) {
      var m = x.match(/^(Store Management|Product Knowledge|Selling Roleplay|Skill การขาย)[^\[]*\[\s*(.*?)\s*\]\s*$/);
      if (m) { scoreCols.push(i); items.push({ c: m[1], l: m[2] }); }
      else if (/^(Store Management|Product Knowledge|Selling Roleplay|Skill การขาย)/.test(x)) { scoreCols.push(i); items.push({ c: x.split(/[\s(]/)[0], l: x }); }
      else if (/^Exam/.test(x)) examCols.push(i);
      else if (x.indexOf('ข้อเสนอแนะ') >= 0) { noteCols.push(i); noteFrom.push(/Director/i.test(x) ? 'Director' : /พนักงาน/.test(x) ? 'SC' : ''); }
    });
    var diag = { sheet: t.name, total: 0, inYear: 0, audits: 0, visitOnly: 0, otherEvaluator: 0, crossScope: 0, unknownEvaluators: {}, unknownBranches: {}, monthDateMismatch: 0, scoreCols: scoreCols.length, examCols: examCols.length, items: items };
    var audits = [], visitMap = {}, logs = [], auditIdx = {}; diag.resubmitted = 0;
    t.rows.forEach(function (r) {
      if (!r || r.every(function (v) { return v == null || v === ''; })) return;
      diag.total++;
      var d = toDate(r[c.date]) || toDate(r[c.ts]);
      var tsD = toDate(r[c.ts]), lag = tsD && d ? Math.max(0, Math.round((tsD - d) / 864e5)) : null; // days between the visit and the form entry
      var yr = firstNum(r[c.year]); if (yr > 2400) yr -= 543;
      if (!yr && d) yr = d.getFullYear();
      if (yr !== 2026) return;
      diag.inYear++;
      var evName = supName(r[c.ev]), scope = scopeOfPerson(evName);
      if (!scope) { diag.otherEvaluator++; if (evName) diag.unknownEvaluators[evName] = (diag.unknownEvaluators[evName] || 0) + 1; return; }
      var brRaw = norm(r[c.br]), br = branchCode(brRaw);
      if (!br && brRaw) diag.unknownBranches[brRaw] = (diag.unknownBranches[brRaw] || 0) + 1;
      var bs = scopeOfBranch(br);
      var cross = bs && bs !== scope; if (cross) diag.crossScope++; // kept but flagged: team boards skip it, the Senior board shows it outside the KPI
      var mName = norm(r[c.month]), mNum = TH_MONTH[mName] || (d ? d.getMonth() + 1 : null);
      var ymAudit = mNum ? '2026-' + pad(mNum) : null;
      if (d && mNum && d.getMonth() + 1 !== mNum) diag.monthDateMismatch++;
      var isNum = function (v) { return typeof v === 'number' || (v !== null && v !== '' && !isNaN(+v)); };
      var sArr = scoreCols.map(function (i) { return isNum(r[i]) ? Number(r[i]) : null; });
      var scores = sArr.filter(function (v) { return v !== null; });
      var exams = examCols.map(function (i) { return r[i]; }).filter(function (v) { return typeof v === 'number' || (v !== null && v !== '' && !isNaN(+v)); }).map(Number);
      var sc = norm(r[c.sc]);
      var method = norm(r[c.method]);
      var online = /online|ออนไลน์/i.test(method);
      var isAudit = sc && !/ไม่ได้ประเมิน|visit|^-$/i.test(sc) && (scores.length > 0 || exams.length > 0);
      var notes = noteCols.map(function (i, j) { var x = norm(r[i]); return x ? { f: noteFrom[j], t: x.slice(0, 400) } : null; }).filter(Boolean);
      var note = notes.map(function (x) { return x.t; }).join(' / ').slice(0, 240);
      // the same SC, same day, same evaluator sent twice = a corrected resubmission: keep the later one only
      var aKey = (d ? ymd(d) : '') + '|' + sc.replace(/\s+/g, '') + '|' + evName, resub = isAudit && auditIdx[aKey] != null;
      if (isAudit) {
        if (resub) diag.resubmitted++; else diag.audits++;
        var aRec = ({ d: d ? ymd(d) : null, ym: ymAudit, sc: sc, br: br, brRaw: br ? undefined : brRaw, ev: evName, scope: scope,
          avg: scores.length ? Math.round(scores.reduce(function (a, b) { return a + b; }, 0) / scores.length * 100) / 100 : null, n: scores.length,
          exam: exams.length ? Math.round(exams.reduce(function (a, b) { return a + b; }, 0) / exams.length * 100) / 100 : null,
          online: online || undefined, x: cross ? 1 : undefined, lag: lag == null ? undefined : lag, note: note || undefined, s: scores.length ? sArr : undefined, notes: notes.length ? notes : undefined });
        if (resub) audits[auditIdx[aKey]] = aRec; else { auditIdx[aKey] = audits.length; audits.push(aRec); }
      } else diag.visitOnly++;
      // Every onsite row is a store visit: one per evaluator + branch + day
      if (!online && d && (br || brRaw)) {
        var k = scope + '|' + evName + '|' + (br || brRaw) + '|' + ymd(d);
        var v = visitMap[k] || (visitMap[k] = { d: ymd(d), ym: ymd(d).slice(0, 7), br: br, brRaw: br ? undefined : brRaw, ev: evName, scope: scope, x: cross ? 1 : undefined, audits: 0 });
        if (isAudit && !resub) v.audits++;
        if (lag != null && (v.lag == null || lag < v.lag)) v.lag = lag;
        notes.forEach(function (n) { var t = n.t.slice(0, 300); v.notes = v.notes || []; if (!v.notes.some(function (o) { return o.t === t; })) v.notes.push({ f: n.f, t: t, sc: isAudit ? sc : undefined }); });
      } else if (!online && d && notes.length) {
        // a Sup's work day with no branch (head office, meeting, other site): keep the note as a log line
        logs.push({ d: ymd(d), ym: ymd(d).slice(0, 7), ev: evName, scope: scope, lag: lag == null ? undefined : lag, t: notes.map(function (n) { return n.t; }).join(' / ').slice(0, 300) });
      }
    });
    var visits = Object.keys(visitMap).map(function (k) { return visitMap[k]; });
    visits.forEach(function (v) { Object.keys(v).forEach(function (k) { if (v[k] === undefined) delete v[k]; }); });
    audits.forEach(function (v) { Object.keys(v).forEach(function (k) { if (v[k] === undefined) delete v[k]; }); });
    diag.visits = visits.length;
    visits.forEach(function (v) { (v.notes || []).forEach(function (n) { if (n.sc === undefined) delete n.sc; }); });
    return { audits: audits, visits: visits, logs: logs, diag: diag };
  }

  /* ---------- Product demand sheet ---------- */
  function parseDemand(wb) {
    var t = sheetRows(wb, ['ผู้แจ้ง', 'Model', 'Status']);
    if (!t) throw new Error('ไม่พบชีทที่มีหัวคอลัมน์ "ผู้แจ้ง", "Model", "Status"');
    var h = t.header, c = {
      ts: col(h, function (x) { return x === 'Timestamp'; }), model: col(h, function (x) { return x === 'Model'; }), size: col(h, function (x) { return x === 'Size'; }),
      skin1: col(h, function (x) { return x === 'Skin1'; }), q1: col(h, function (x) { return x.indexOf('จำนวนSkin1') >= 0; }), skin2: col(h, function (x) { return x === 'Skin2'; }),
      q2: col(h, function (x) { return x.indexOf('จำนวนSkin2') >= 0; }), by: col(h, function (x) { return x.indexOf('ผู้แจ้ง') >= 0; }), others: col(h, function (x) { return x === 'Others'; }),
      plan: col(h, function (x) { return x.indexOf('Plan') >= 0; }), status: col(h, function (x) { return x === 'Status'; })
    };
    var diag = { sheet: t.name, total: 0, inYear: 0, counted: 0, otherRequester: {}, qtyUnreadable: 0 }, rows = [];
    t.rows.forEach(function (r) {
      if (!r || r.every(function (v) { return v == null || v === ''; })) return;
      diag.total++;
      var d = toDate(r[c.ts]); if (!d || d.getFullYear() !== 2026) return;
      diag.inYear++;
      var by = supName(r[c.by]), scope = scopeOfPerson(by);
      if (!scope) { var key = by || '(ว่าง)'; diag.otherRequester[key] = (diag.otherRequester[key] || 0) + 1; return; }
      var q1raw = r[c.q1], q2raw = r[c.q2];
      var qty = firstNum(q1raw) + firstNum(q2raw);
      if ((q1raw == null || q1raw === '') && (q2raw == null || q2raw === '')) diag.qtyUnreadable++;
      var model = norm(r[c.model]) || norm(r[c.others]).slice(0, 60) || '(ไม่ระบุรุ่น)';
      var status = norm(r[c.status]);
      diag.counted++;
      rows.push({ d: ymd(d), ym: ymd(d).slice(0, 7), by: by, scope: scope, item: model, size: norm(r[c.size]) || undefined,
        skin: [norm(r[c.skin1]), norm(r[c.skin2])].filter(Boolean).join(' / ') || undefined, qty: qty, done: /^done/i.test(status), status: status || undefined, plan: norm(r[c.plan]) || undefined });
    });
    rows.forEach(function (v) { Object.keys(v).forEach(function (k) { if (v[k] === undefined) delete v[k]; }); });
    return { rows: rows, diag: diag };
  }

  /* ---------- Main Stock file + Focus list ---------- */
  function focusBase(name) {
    var parts = String(name).replace(/\s+3\/4/, '').split('/');
    return parts.map(function (p) { return p.split(',')[0].replace(/\s+V\.?\s*\d+$/i, '').trim(); }).filter(Boolean);
  }
  function makeMatcher(focus) {
    var bases = [];
    focus.forEach(function (f) { if (f.name === 'All Model') return; focusBase(f.name).forEach(function (b) { bases.push({ b: b, item: f.name }); }); });
    bases.sort(function (a, b) { return b.b.length - a.b.length; });
    var exact = { 'Key Chain': true };
    var glasses = focus.some(function (f) { return f.name === 'All Model'; });
    return function (model, group) {
      if (glasses && group && /Glasses Case/i.test(group)) return 'All Model';
      var m = norm(model); if (!m) return null;
      for (var i = 0; i < bases.length; i++) {
        var b = bases[i].b;
        if (exact[b]) { if (m.toLowerCase() === b.toLowerCase()) return bases[i].item; continue; }
        if (m.toLowerCase().indexOf(b.toLowerCase()) === 0) { var nx = m.charAt(b.length); if (!nx || /[\s.]/.test(nx)) return bases[i].item; }
      }
      return null;
    };
  }
  // excl = { BRANCH: [focus item names the branch does not carry] } (e.g. Paragon has no sunglass)
  function parseStock(wb, focus, excl) {
    excl = excl || {}; var skip = function (k, item) { return (excl[k] || []).indexOf(item) >= 0; };
    // two layouts: "Main Stock" (Barcode / Prod.Group / Qty., date in A1 "Stock: 25.09.69")
    // and "Stock Card" (Serial No. / Product Category / QTY, date in B1, sheet named "Stock 30.09.69")
    var t = sheetRows(wb, ['Barcode', 'Model', 'Branch']) || sheetRows(wb, ['Serial No', 'Model', 'Branch']);
    if (!t) throw new Error('ไม่พบชีทที่มีหัวคอลัมน์ "Barcode" หรือ "Serial No.", "Model", "Branch"');
    var sheet = wb.Sheets[t.name], a1 = sheet.A1 ? String(sheet.A1.v) : '', stockDate = null;
    var dmy = function (txt) { var md = String(txt || '').match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/); if (!md) return null; var y = +md[3]; if (y < 100) y += 2500; if (y > 2400) y -= 543; return y + '-' + pad(+md[2]) + '-' + pad(+md[1]); };
    stockDate = dmy(a1) || dmy(t.name);
    if (!stockDate) { for (var ci = 0; ci < 6 && !stockDate; ci++) { var cell = sheet[XLSX.utils.encode_cell({ r: 0, c: ci })]; if (cell && typeof cell.v === 'number') { var dd = toDate(cell.v); if (dd) stockDate = ymd(dd); } } }
    var h = t.header, find = function (names) { for (var i = 0; i < names.length; i++) { var k = h.indexOf(names[i]); if (k >= 0) return k; } return -1; };
    var cB = find(['Barcode', 'Serial No.', 'Serial No']), cM = h.indexOf('Model'), cG = find(['Prod.Group', 'Product Category']), cBr = h.indexOf('Branch'), cQ = find(['Qty.', 'QTY', 'Qty']);
    if (!/stock/i.test(a1) || !/\d/.test(a1)) a1 = stockDate ? 'Stock: ' + stockDate : a1;
    var match = makeMatcher(focus);
    var wToCode = {}; Object.keys(MAIN).forEach(function (k) { wToCode[MAIN[k].w] = k; });
    var poolSet = {}; POOL_EXTRA.forEach(function (w) { poolSet[w] = 1; }); Object.keys(MAIN).forEach(function (k) { poolSet[MAIN[k].w] = 1; });
    var seen = {}, have = {}, pool = {}, rowsRead = 0, dup = 0;
    Object.keys(MAIN).forEach(function (k) { have[k] = {}; });
    t.rows.forEach(function (r) {
      if (!r || !r[cB]) return;
      if (seen[r[cB]]) { dup++; return; } seen[r[cB]] = 1; // 1 barcode = 1 piece
      rowsRead++;
      var w = norm(r[cBr]); if (!poolSet[w]) return;
      var item = match(r[cM], cG >= 0 ? r[cG] : null); if (!item) return;
      var q = cQ >= 0 ? (+r[cQ] || 1) : 1;
      pool[item] = (pool[item] || 0) + q;
      var code = wToCode[w]; if (code) have[code][item] = (have[code][item] || 0) + q;
    });
    var need = {}; // item -> list of main branches that must carry it
    focus.forEach(function (f) { need[f.name] = Object.keys(MAIN).filter(function (k) { return f[MAIN[k].t] && !skip(k, f.name); }); });
    var exempt = {}; Object.keys(MAIN).forEach(function (k) { exempt[k] = []; });
    var shortages = [];
    focus.forEach(function (f) {
      var br = need[f.name], p = pool[f.name] || 0, short = Math.max(0, br.length - p);
      if (!short) return;
      var missing = br.filter(function (k) { return !have[k][f.name]; })
        .sort(function (a, b) { return TIER_ORDER[MAIN[a].tier] - TIER_ORDER[MAIN[b].tier] || a.localeCompare(b); });
      var given = missing.slice(0, short);
      given.forEach(function (k) { exempt[k].push(f.name); });
      shortages.push({ item: f.name, pool: p, need: br.length, exempted: given });
    });
    var branches = {};
    Object.keys(MAIN).forEach(function (k) {
      var req = focus.filter(function (f) { return f[MAIN[k].t] && !skip(k, f.name); }).map(function (f) { return f.name; });
      var h2 = req.filter(function (i) { return have[k][i]; }).map(function (i) { return [i, have[k][i]]; });
      var ex = exempt[k];
      var miss = req.filter(function (i) { return !have[k][i] && ex.indexOf(i) < 0; });
      branches[k] = { req: req.length, have: h2, exempt: ex, missing: miss, ok: h2.length + ex.length, pcs: h2.reduce(function (a, x) { return a + x[1]; }, 0) };
    });
    var poolAll = {}; focus.forEach(function (f) { poolAll[f.name] = pool[f.name] || 0; });
    return { stockDate: stockDate, stockLabel: a1, rowsRead: rowsRead, dupBarcodes: dup, branches: branches, shortages: shortages, pool: poolAll };
  }


  /* ---------- Sales_Normalized (daily sales form) ---------- */
  var OLD_CODE = { ZENS: 'ZENL', STKS: 'STKL', PKCS: 'PKCL', PHCS: 'PHCL', SMUA: 'SMAS' };
  var EVENT_SCOPE = { ZENE: 'BKK', ICDE: 'BKK', TMNE: 'BKK', TMTE: 'BKK', CTWE: 'BKK', CTWF: 'BKK', DISE: 'BKK', BDHE: 'BKK', PKCE: 'UPC', PKCE2: 'UPC', JCPE: 'UPC' };
  function salesScope(code) {
    if (MAIN[code]) return MAIN[code].s;
    if (code === 'O2O') return 'BKK';
    if (EVENT_SCOPE[code]) return EVENT_SCOPE[code];
    if (code === 'TMTS' || code === 'EMPL' || code === 'EMPM') return 'BKK';
    if (code === 'JCLS') return 'UPC';
    return null;
  }
  function isEvent(code) { return !!EVENT_SCOPE[code] || /^[A-Z]{3}[EF]\d?$/.test(code); }
  function parseSales(wb) {
    var ws = wb.Sheets['Sales_Normalized'];
    if (!ws) throw new Error('ไม่พบชีท Sales_Normalized');
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    var h = (rows[0] || []).map(norm);
    var cD = h.indexOf('ยอดขายวันที่'), cB = h.indexOf('สาขา'), cP = h.indexOf('พนักงานขาย'), who = {}, cT = col(h, function (x) { return x.indexOf('ยอดขายประจำสาขาวันนี้') === 0; });
    if (cD < 0 || cB < 0 || cT < 0) throw new Error('ไม่พบหัวคอลัมน์ ยอดขายวันที่ / สาขา / ยอดขายประจำสาขาวันนี้ ในชีท Sales_Normalized');
    var diag = { rows: 0, unknown: {}, wholesale: 0 }, day = {}, lastDate = {};
    rows.slice(1).forEach(function (r) {
      if (!r) return; var d = toDate(r[cD]), code = norm(r[cB]).toUpperCase(); if (!d || !code) return;
      if (d.getFullYear() !== 2026) return;
      diag.rows++;
      code = OLD_CODE[code] || code;
      var raw = cP >= 0 ? norm(r[cP]) : '', mm = raw.match(/^(BBYD\d+)\s*-\s*(.+)$/);
      if (raw && !(mm && /^BBYD1\d{3}$/.test(mm[1])) && !/^PT\b/i.test(mm ? mm[2] : raw)) { // PT pools are not people
        var nm = mm ? mm[2].trim() : raw, w = who[nm] || (who[nm] = { name: nm, first: ymd(d), last: ymd(d), br: {} });
        if (ymd(d) < w.first) w.first = ymd(d); if (ymd(d) > w.last) w.last = ymd(d);
        var s0 = salesScope(code); if (s0) w.br[s0] = (w.br[s0] || 0) + 1;
      }
      var k = ymd(d) + '|' + code, v = +r[cT] || 0;
      (day[k] = day[k] || {})[v] = 1; // same shop total typed by every SC -> count each distinct value once
      var ymk = ymd(d).slice(0, 7); if (!lastDate[ymk] || ymd(d) > lastDate[ymk]) lastDate[ymk] = ymd(d);
    });
    var months = {};
    Object.keys(day).forEach(function (k) {
      var p = k.split('|'), d = p[0], code = p[1], ymk = d.slice(0, 7);
      var amt = Object.keys(day[k]).reduce(function (a, x) { return a + Number(x); }, 0);
      var M = months[ymk] || (months[ymk] = { lastDate: lastDate[ymk], wholesale: 0, BKK: { sales: 0, branches: {}, events: {} }, UPC: { sales: 0, branches: {}, events: {} } });
      if (code === 'WHOLESALE') { M.wholesale += amt; return; }
      var sc = salesScope(code);
      if (!sc) { diag.unknown[code] = (diag.unknown[code] || 0) + 1; return; }
      var S2 = M[sc]; S2.sales += amt; S2.branches[code] = (S2.branches[code] || 0) + amt;
      if (isEvent(code)) { var e = S2.events[code] || (S2.events[code] = { code: code, days: 0, sales: 0 }); e.days++; e.sales += amt; }
    });
    Object.keys(months).forEach(function (k) {
      var M = months[k], n = 0;
      ['BKK', 'UPC'].forEach(function (sc) {
        var ev = Object.keys(M[sc].events).map(function (c) { return M[sc].events[c]; });
        M[sc].events = ev; M[sc].eventDays = ev.reduce(function (a, e) { return a + e.days; }, 0);
        M[sc].eventsWithSales = ev.filter(function (e) { return e.sales > 0; }).length; n += M[sc].eventsWithSales;
        M[sc].sales = Math.round(M[sc].sales * 100) / 100;
      });
      M.eventsWithSales = n;
    });
    var sc = Object.keys(who).map(function (k) { var w = who[k], b = w.br.BKK || 0, u = w.br.UPC || 0; return { name: w.name, first: w.first, last: w.last, scope: b || u ? (u > b ? 'UPC' : 'BKK') : null }; });
    return { months: months, sc: sc, diag: diag };
  }

  /* ---------- HR staff sheet: only nickname / team / start & end month are kept ---------- */
  function parseStaff(wb) {
    var t = sheetRows(wb, ['ชื่อเล่น', 'สถานะพนักงาน', 'สถานที่ปฏิบัติงาน']);
    if (!t) throw new Error('ไม่พบชีทที่มีหัวคอลัมน์ "ชื่อเล่น", "สถานะพนักงาน", "สถานที่ปฏิบัติงาน"');
    var h = t.header, c = {
      nick: h.indexOf('ชื่อเล่น'), start: h.indexOf('วันเริ่มงาน'), pos: h.indexOf('ตำแหน่งปัจจุบัน'), loc: h.indexOf('สถานที่ปฏิบัติงาน'),
      status: h.indexOf('สถานะพนักงาน'), end: col(h, function (x) { return x.indexOf('สถานะพนักงานสิ้นสุด') === 0; })
    };
    var diag = { total: 0, notSales: [], leftNoDate: [], unknownLoc: {} }, people = [];
    t.rows.forEach(function (r) {
      if (!r) return; var name = norm(r[c.nick]); if (!name) return;
      diag.total++;
      var pos = norm(r[c.pos]);
      if (pos && pos !== 'พนักงานขาย') { diag.notSales.push(name + ' (' + pos + ')'); return; }
      var loc = norm(r[c.loc]), scope = /ภูเก็ต|สมุย|phuket|samui/i.test(loc) ? 'UPC' : /พระนคร|ธน|กรุงเทพ|bangkok|bkk/i.test(loc) ? 'BKK' : null;
      if (!scope) { diag.unknownLoc[loc || '(ว่าง)'] = (diag.unknownLoc[loc || '(ว่าง)'] || 0) + 1; return; }
      var left = /ลาออก|พ้นสภาพ|resign/i.test(norm(r[c.status]));
      var sd = toDate(r[c.start]), ed = c.end >= 0 ? toDate(r[c.end]) : null;
      if (left && !ed) diag.leftNoDate.push(name); // month decided later from the last sales-form entry
      var o = { name: name, scope: scope, start: sd ? ymd(sd) : null };
      if (left) { o.left = true; if (ed) o.end = ymd(ed); }
      people.push(o);
    });
    diag.people = people.length;
    return { people: people, diag: diag };
  }

  return { parseSales: parseSales, parseStaff: parseStaff, salesScope: salesScope, focusFromWorkbook: focusFromWorkbook, MAIN: MAIN, POOL_EXTRA: POOL_EXTRA, parseAudit: parseAudit, parseDemand: parseDemand, parseStock: parseStock, branchCode: branchCode, scopeOfPerson: scopeOfPerson, scopeOfBranch: scopeOfBranch, focusBase: focusBase };
})();
if (typeof module !== 'undefined') module.exports = GSProc;
