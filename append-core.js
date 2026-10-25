/* Append-to-previous-report logic.
   Previous report is loaded as-is (ExcelJS), only NEW rows are appended at the bottom of each data tab.
   Old rows, column widths, alignment, Notes and Glossary tabs are never modified. */
(function (root) {
  'use strict';
  const RC = root.ReportCore;
  const SHEET = { findings: 'Findings', completed: 'Takedown Completed', pending: 'Takedown Pending', delist: 'Delisting', social: 'Major Social Platform Take down', summary: 'Summary' };

  const nh = (h) => String(h == null ? '' : h).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  const findSheet = (wb, name) => wb.worksheets.find((w) => nh(w.name) === nh(name)) || null;
  const clone = (o) => (o ? JSON.parse(JSON.stringify(o)) : o);

  function cellVal(cell) {
    let v = cell.value;
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map((t) => t.text).join('');
      if (v.result !== undefined) return v.result;
      if (v.text !== undefined) return v.text;
      if (v.hyperlink) return v.hyperlink;
      return null;
    }
    return v;
  }
  const str = (v) => (v == null ? '' : String(v).trim());
  /* every word starts with an uppercase letter (rest of the word is left as is: VK, YouTube ...) */
  const cap = (v) => (typeof v === 'string' ? v.replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase()) : v);
  const DELIST_FMT = '[hh]:mm:ss';

  function headerMap(ws) {
    const m = {}; ws.getRow(1).eachCell({ includeEmpty: false }, (c, n) => { const k = nh(cellVal(c)); if (k && !(k in m)) m[k] = n; });
    return m;
  }
  function lastDataRow(ws, srCol) {
    for (let r = ws.rowCount; r >= 2; r--) { const v = cellVal(ws.getRow(r).getCell(srCol)); if (v !== null && v !== '') return r; }
    return 1;
  }

  /* ---- domain helpers ---- */
  function hostOf(u) {
    const s = str(u); const m = s.match(/^[a-z][a-z0-9+.-]*:\/\/([^\/?#]+)/i);
    let h = m ? m[1] : s.split(/[\/?#]/)[0];
    return h.replace(/^.*@/, '').replace(/:\d+$/, '').toLowerCase().replace(/^www\./, '');
  }
  function baseOf(h) {
    if (/^\d+(\.\d+){3}$/.test(h)) return h;
    const p = h.split('.'); if (p.length <= 2) return h;
    const n = p[p.length - 1].length === 2 && /^(co|com|org|net|gov|edu|ac|or|ne|go)$/.test(p[p.length - 2]) ? 3 : 2;
    return p.slice(-n).join('.');
  }

  /* ---- load previous report + index it ---- */
  async function loadPrevious(ExcelJS, buf) {
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
    const info = { wb, keys: { findings: new Set(), completed: new Set(), pending: new Set(), delist: new Set(), social: new Set() }, domExact: {}, domBase: {}, findCountry: {}, pendingUrls: [], completedUrls: new Set(), pendingDolby: {}, findUrls: new Set(), missing: [], prevCounts: {} };
    for (const k of ['findings', 'completed', 'pending', 'delist', 'social']) {
      const ws = findSheet(wb, SHEET[k]);
      if (!ws) { info.missing.push(SHEET[k]); continue; }
      const H = headerMap(ws); const cUrl = H['url']; const cPlat = H['platform/source'];
      if (!cUrl) { info.missing.push(SHEET[k] + ' (URL column)'); continue; }
      const cCty = H['host country/domain country']; const cReason = H['reason takedown is pending:'] || H['reason takedown is pending']; const cHsp = H['hsp name']; const cIp = H['hsp ip address']; const cSrv = H['server location']; const cDol = H['dolby (y/n)'];
      let n = 0;
      for (let r = 2; r <= ws.rowCount; r++) {
        const row = ws.getRow(r); const url = str(cellVal(row.getCell(cUrl))); if (!url) continue; n++;
        const plat = cPlat ? cap(str(cellVal(row.getCell(cPlat)))) : '';
        info.keys[k].add((k === 'delist' || k === 'social' ? '' : plat) + '|' + url);
        if (k === 'findings' && cCty) { const cc = cellVal(row.getCell(cCty)); if (cc !== null && cc !== '') info.findCountry[plat + '|' + url] = cc; }
        if (k === 'findings') info.findUrls.add(normUrl(url));
        if (k === 'completed') info.completedUrls.add(url);
        if (k === 'pending') { info.pendingUrls.push(url); if (cDol) { const dv = cellVal(row.getCell(cDol)); if (dv !== null && dv !== '') info.pendingDolby[plat + '|' + url] = dv; } }
        if (k === 'pending') {
          const host = hostOf(url); if (!host) continue;
          const rec = { reason: cReason ? cellVal(row.getCell(cReason)) : null, hsp: cHsp ? cellVal(row.getCell(cHsp)) : null, ip: cIp ? cellVal(row.getCell(cIp)) : null, server: cSrv ? cellVal(row.getCell(cSrv)) : null };
          if (!(host in info.domExact)) info.domExact[host] = rec;          /* first placed wins */
          const b = baseOf(host); if (!(b in info.domBase)) info.domBase[b] = rec;
        }
      }
      info.prevCounts[k] = n;
    }
    return info;
  }

  /* ---- append rows to one sheet; each row = { 'Header Name': value } ---- */
  function appendRows(ws, rows, fallback, fmts) {
    if (!rows.length) return 0;
    const H = headerMap(ws); const srCol = H['sr. no'] || 1; const nCols = Math.max(ws.getRow(1).cellCount, ws.columnCount);
    const last = lastDataRow(ws, srCol); const tmpl = last >= 2 ? ws.getRow(last) : null;
    let sr = last >= 2 ? Number(cellVal(tmpl.getCell(srCol))) || (last - 1) : 0;
    let at = last >= 2 ? last : 1;
    for (const o of rows) {
      at++; sr++; const row = ws.getRow(at);
      const vals = {}; Object.keys(o).forEach((k) => { vals[nh(k)] = o[k]; }); vals['sr. no'] = sr;
      for (let c = 1; c <= nCols; c++) {
        const cell = row.getCell(c);
        if (tmpl) cell.style = clone(tmpl.getCell(c).style); else cell.style = clone(fallback.cellStyle(c));
      }
      if (tmpl && tmpl.height) row.height = tmpl.height;
      for (const key of Object.keys(vals)) {
        const col = H[key]; if (!col) continue; const v = vals[key];
        const cell = row.getCell(col);
        cell.value = v === undefined ? null : v;
        if (v instanceof Date && (!cell.numFmt || cell.numFmt === 'General')) cell.numFmt = 'mm/dd/yyyy';
        if (fmts && fmts[key]) cell.numFmt = fmts[key];
      }
      row.commit && row.commit();
    }
    return rows.length;
  }

  const FB_BORDER = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  const fallbackFor = (leftCols) => ({ cellStyle: (c) => ({ font: { name: 'Calibri', size: 12 }, border: FB_BORDER, alignment: { horizontal: leftCols.includes(c) ? 'left' : 'center', vertical: 'middle' } }) });

  function countNumeric(ws, col) { let n = 0; for (let r = 2; r <= ws.rowCount; r++) { const v = cellVal(ws.getRow(r).getCell(col)); if (typeof v === 'number') n++; } return n; }

  /* ---- Summary: refresh values only (existing cell styles untouched) ---- */
  function refreshSummary(info, ctx) {
    const wb = info.wb; const sw = findSheet(wb, SHEET.summary); const totals = {};
    const { S, added, newComp, newPend, newFind } = ctx;
    const cnt = (k) => { const ws = findSheet(wb, SHEET[k]); return ws ? countNumeric(ws, (headerMap(ws)['sr. no'] || 1)) : 0; };
    totals.total = cnt('findings'); totals.comp = cnt('completed');
    /* open issues = rows in Takedown Pending whose URL is NOT already in Takedown Completed (old pending rows are kept, so subtract the ones now completed) */
    const doneUrls = new Set(info.completedUrls); newComp.forEach((x) => doneUrls.add(x.url));
    const pendUrls = info.pendingUrls.concat(newPend.map((x) => x.url));
    const stale = pendUrls.filter((u) => doneUrls.has(u)).length;
    totals.rawPending = cnt('pending'); totals.open = totals.rawPending - stale; totals.stalePending = stale;
    if (sw) {
      const H = headerMap(sw); const pad = (n) => String(n).padStart(2, '0');
      const colOf = (names, fb) => { for (const n of names) if (H[n]) return H[n]; return fb; };
      const set = (col, value, keepFormula) => {
        const c = sw.getRow(2).getCell(col); const cur = c.value;
        if (keepFormula && cur && typeof cur === 'object' && !(cur instanceof Date) && (cur.formula || cur.sharedFormula)) c.value = Object.assign({}, cur, { result: value }); else c.value = value;
      };
      /* new data  -> Total Infringements (+ since prior) | removal file -> Total Takedowns (+ since prior) | both -> Open issues + Percentage */
      const hasData = !ctx.removalOnly, hasRem = ctx.remMoved != null;
      const tdSince = hasRem ? ctx.remMoved : 0;   /* no removal file -> 00 (never keep the old number) */
      if (hasData) set(colOf(['total infringements'], 1), totals.total, true);
      if (hasRem) set(colOf(['total takedowns'], 3), totals.comp, true);
      set(colOf(['open issues remaining'], 5), totals.open, stale === 0);   /* formula kept only when it is exactly right */
      set(colOf(['current takedown percentage'], 6), totals.total ? totals.comp / totals.total : 0, true);
      const sinceTxt = S.since == null ? '' : ' (' + RC.fmtTime(S.since) + ' EST - ' + RC.dayLabel(S.since) + ')';   /* actual time typed in "Since prior report" */
      if (hasData) set(colOf(['total infringements since prior report (insert time)'], 2), pad(added.findings) + sinceTxt);
      set(colOf(['total takedowns since prior report (insert time)'], 4), pad(tdSince) + sinceTxt);
      /* the "New Added URLs" list no longer lives on Summary (URLs go to Findings) - drop it if an older report has it */
      const ucol = H['new added urls'];
      if (ucol) { if (ucol === sw.columnCount) sw.spliceColumns(ucol, 1); else for (let r = 1; r <= sw.rowCount; r++) { const cc = sw.getRow(r).getCell(ucol); cc.value = null; cc.style = {}; } }
    }
    wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
    return totals;
  }

  function appendToPrevious(info, res, opts) {
    const wb = info.wb; const refTs = opts.refTs; const D = (ts) => new Date(RC.dayStart(ts));
    const added = { findings: 0, completed: 0, pending: 0, delist: 0, social: 0 };
    const dom = { exact: 0, base: 0, fallback: 0 };
    const key = (x) => x.platform + '|' + x.url;

    /* Findings: a URL already present in the previous report is NOT added again */
    const seenNew = new Set();
    const newFind = res.findings.filter((x) => { const u = normUrl(x.url); if (info.findUrls.has(u) || seenNew.has(u)) return false; seenNew.add(u); return true; });
    const newComp = res.completedList.filter((x) => !info.keys.completed.has(key(x)));
    const newPend = res.pendingList.filter((x) => !info.keys.pending.has(key(x)) && !info.keys.completed.has(key(x)));
    const newDel = res.delist.filter((x) => !info.keys.delist.has('|' + x.url));
    const newSoc = res.socialRemoved.filter((x) => !info.keys.social.has('|' + x.url));

    const run = (k, list, mk, left, fmts) => {
      const ws = findSheet(wb, SHEET[k]); if (!ws) { if (list.length) info.missing.push(SHEET[k]); return; }
      added[k] = appendRows(ws, list.map(mk), fallbackFor(left), fmts);
    };

    run('findings', newFind, (x) => ({ 'Platform/Source': x.platform, Book: x.book, Title: x.title, URL: x.url, 'Coverage by Date - at EoD': RC.coverage(x.ident), 'Identification Date (IST)': D(x.ident), 'Identification Time (IST)': RC.fmtTime(x.ident), 'Channel/Domain': x.domain, 'Unique Subscribers': x.subs, Views: x.views, Downloads: x.downloads, 'File Size': x.fileSize, 'Dolby (Y/N)': null, 'Method of Verification': x.method, 'Host Country/Domain Country': x.country }), [4]);

    run('completed', newComp, (x) => ({ 'Platform/Source': x.platform, Book: x.book, Title: x.title, URL: x.url, 'Channel/Domain': x.domain, 'Coverage by Date - at EoD': RC.coverage(x.ident), 'Identification Date (IST)': D(x.ident), 'Identification Time (IST)': RC.fmtTime(x.ident), 'Takedown Initiated Date (IST)': D(x.initiated), 'Takedown Initiated Time (IST)': RC.fmtTime(x.initiated), 'Takedown Completed Date (IST)': D(x.completed), 'Takedown Completed Time (IST)': RC.fmtTime(x.completed), 'Takedown Duration (TAT)': Math.max(0, Math.floor((x.completed - x.initiated) / 3600e3)), 'Takedown Unit (TAT)': 'Hrs', 'File Size': x.fileSize, 'Dolby (Y/N)': info.pendingDolby[key(x)] !== undefined ? info.pendingDolby[key(x)] : null }), [4]);

    run('pending', newPend, (x) => {
      const host = hostOf(x.url); let p = info.domExact[host];
      if (p) dom.exact++; else { p = info.domBase[baseOf(host)]; if (p) dom.base++; }
      let reason, hsp, ip, server;
      if (p) { reason = p.reason; hsp = p.hsp; ip = p.ip; server = p.server; } else { dom.fallback++; reason = x.reason; hsp = x.hsp; ip = x.hspIp; server = x.server; }
      /* Findings sheet has priority: Server Location = Host Country/Domain Country of the same URL in Findings */
      if (x.blankReason) reason = null;   /* Facebook / Instagram / YouTube / X: Reason Takedown is Pending stays blank */
      server = info.findCountry[key(x)] !== undefined ? info.findCountry[key(x)] : x.country;
      return { 'Platform/Source': x.platform, Book: x.book, Title: x.title, URL: x.url, 'Coverage by Date - at EoD': RC.coverage(x.ident), 'Identification Date (IST)': D(x.ident), 'Identification Time (IST)': RC.fmtTime(x.ident), 'Takedown Initiated Date (IST)': D(x.initiated), 'Takedown Initiated Time (IST)': RC.fmtTime(x.initiated), 'Takedown Pending Duration': Math.max(1, Math.round((refTs - x.initiated) / 3600e3)), 'Takedown Pending Unit': 'Hrs', 'Reason Takedown is Pending:': reason, 'HSP Name': hsp, 'HSP IP Address': ip, 'Server Location': server, 'File Size': x.fileSize, 'Dolby (Y/N)': null };
    }, [4]);

    run('delist', newDel, (x) => ({ 'Source/Search Engine': 'Linking', 'Offending URL': '-', Book: x.book, Title: x.title, URL: x.url, 'Coverage by Date - at EoD': RC.coverage(x.ident), 'Delisting Request Date (IST)': D(x.delReq), 'Delisting Request Time (IST)': RC.fmtTime(x.delReq), 'Delisting Status': 'Delisted', 'Delisting Completion Time (IST)': RC.fmtTime(x.delComp), 'Delisting Completion Date (IST)': D(x.delComp), 'Time To Delist (HRS)': (Math.floor(x.delComp / 60000) - Math.floor(x.delReq / 60000)) / 1440 }), [5], { 'time to delist (hrs)': DELIST_FMT });

    run('social', newSoc, (x) => ({ 'Platform/Source': x.platform, Book: x.book, Title: x.title, URL: x.url, 'Coverage by Date - at EoD': RC.coverage(x.ident), 'Identification Date (IST)': D(x.ident), 'Identification Time (IST)': RC.fmtTime(x.ident), 'Channel/Domain': x.domain, 'Unique Subscribers': x.subs, Views: x.views, Downloads: x.downloads, 'File Size': x.fileSize, 'Method of Verification': x.method, 'Host Country/Domain Country': x.country, Status: 'Removed', 'Platform Name': x.platformName }), [4]);

    const totals = refreshSummary(info, { S: res.summary, added, newComp, newPend, newFind, remMoved: opts.remMoved == null ? null : opts.remMoved });
    return { added, dom, totals, ctx: { S: res.summary, added, newComp } };
  }

  /* removal is applied AFTER the new data was appended -> refresh Summary once more (new Pending rows that got completed are already gone from the sheet) */
  function finishSummary(info, ap, remMoved) {
    return refreshSummary(info, { S: ap.ctx.S, added: ap.ctx.added, newComp: ap.ctx.newComp, newPend: [], newFind: [], remMoved });
  }


  /* =====================================================================
     PENDING REMOVAL UPDATE
     Input  : Excel with only  URL | Removal Time (IST)
     Effect : every matching row is MOVED from "Takedown Pending" to "Takedown Completed"
              (Social rows -> "Major Social Platform Take down"), Dolby (Y/N) travels with it,
              Sr. No of the Pending tab is renumbered, Summary totals are refreshed.
     ===================================================================== */
  const SOCIAL_PLATFORM = 'Social Media & UGC';
  /* platform labels whose rows live in "Major Social Platform Take down" once removed */
  const SOCIAL_LABELS = new Set(['social media & ugc', 'facebook', 'instagram', 'market place', 'x', 'youtube']);
  const BOTH_LABELS = new Set(['social media & ugc', 'facebook', 'instagram', 'market place', 'x', 'youtube']);   /* these also get a row in Takedown Completed */
  const SOCIAL_NAMES = [['bilibili', 'BiliBili'], ['vk.', 'VK'], ['ximalaya', 'Ximalaya'], ['chomikuj', 'Chomikuj'], ['spotify', 'Spotify'], ['soundcloud', 'SoundCloud'], ['archive.org', 'Archive.org'], ['youtube', 'YouTube'], ['youtu.be', 'YouTube'], ['yumpu', 'Yumpu']];
  const normUrl = (u) => str(u).toLowerCase().replace(/^(https?:\/\/)www\./, '$1').replace(/#.*$/, '').replace(/\/+$/, '');

  /* "IST wall-clock" ms from a cell / text. Accepts Excel serial, Date, ISO, dd/mm/yyyy [hh:mm[:ss] [AM|PM]] */
  function parseRemovalTime(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : v.getTime();
    if (typeof v === 'number') return v > 1 ? Math.round((v - 25569) * 86400000) : null;
    const t = String(v).trim(); if (!t) return null;
    const iso = RC.toTs(t); if (iso !== null && /^\d{4}-\d{2}-\d{2}/.test(t)) return iso;
    const m = t.match(/^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})(?:[ T,]+(\d{1,2})(?::(\d{2}))?(?::(\d{2}))?\s*([AaPp])\.?[Mm]?\.?)?$/);
    if (!m) { const m2 = t.match(/^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/); if (!m2) return null; return fromParts(m2); }
    return fromParts(m);
    function fromParts(mm) {
      let y, mo, d; const a = +mm[1], b = +mm[2], c = +mm[3];
      if (mm[1].length === 4) { y = a; mo = b; d = c; }
      else { y = mm[3].length === 2 ? 2000 + c : c; if (b > 12) { mo = a; d = b; } else { d = a; mo = b; } }   /* IST style dd/mm unless clearly mm/dd */
      let h = mm[4] != null ? +mm[4] : 0; const mi = mm[5] != null ? +mm[5] : 0; const sc = mm[6] != null ? +mm[6] : 0;
      if (mm[7]) { const pm = /p/i.test(mm[7]); h = (h % 12) + (pm ? 12 : 0); }
      if (mo < 1 || mo > 12 || d < 1 || d > 31 || h > 23 || mi > 59 || sc > 59) return null;
      const ts = Date.UTC(y, mo - 1, d, h, mi, sc); return new Date(ts).getUTCDate() === d ? ts : null;
    }
  }

  /* read the uploaded removal Excel (SheetJS) -> [{ url, ts, rawTime, line }] */
  function parseRemovalFile(XLSX, buf) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    if (!aoa.length) throw new Error('Removal file khali hai.');
    const heads = aoa[0].map((h) => nh(h));
    let cu = heads.findIndex((h) => h === 'url' || h === 'urls' || h === 'infringing url' || h === 'source url'); if (cu < 0) cu = heads.findIndex((h) => h.includes('url'));
    let ct = heads.findIndex((h) => h.includes('removal') && h.includes('time')); if (ct < 0) ct = heads.findIndex((h) => h.includes('time') || h.includes('date'));
    if (cu < 0 || ct < 0) throw new Error('Removal file me "URL" aur "Removal Time (IST)" columns chahiye (template download karke use karo).');
    const rows = [];
    for (let r = 1; r < aoa.length; r++) {
      const url = str(aoa[r][cu]); if (!url) continue;
      rows.push({ url, ts: parseRemovalTime(aoa[r][ct]), rawTime: aoa[r][ct], line: r + 1 });
    }
    return rows;
  }

  /* cell helpers for reading "date + hh:mm AM" pairs written by this tool (or edited by hand in Excel) */
  function timeOfDay(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return v.getUTCHours() * 3600e3 + v.getUTCMinutes() * 60e3 + v.getUTCSeconds() * 1e3;
    if (typeof v === 'number') return Math.round((v % 1) * 86400000 / 60000) * 60000;
    const m = String(v).trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?$/); if (!m) return null;
    let h = +m[1]; if (m[4]) h = (h % 12) + (/p/i.test(m[4]) ? 12 : 0);
    return h * 3600e3 + +m[2] * 60e3 + (m[3] ? +m[3] * 1e3 : 0);
  }
  function dateOnly(v) {
    if (v instanceof Date) return RC.dayStart(v.getTime());
    if (typeof v === 'number') return RC.dayStart(Math.round((v - 25569) * 86400000));
    const t = parseRemovalTime(v); return t == null ? null : RC.dayStart(t);
  }
  const stamp = (d, t) => { const a = dateOnly(d), b = timeOfDay(t); return a == null ? null : a + (b || 0); };
  const platName = (url, plat) => { if (/^(facebook|instagram|market place|x|youtube)$/i.test(str(plat))) return str(plat); const u = url.toLowerCase(); for (const [k, n] of SOCIAL_NAMES) if (u.includes(k)) return cap(n); return '-'; };

  function applyRemovalUpdate(info, entries) {
    const wb = info.wb; const D = (ts) => new Date(RC.dayStart(ts));
    const out = { total: entries.length, movedCompleted: 0, movedSocial: 0, moved: 0, dolbyCarried: 0, notFound: [], alreadyDone: [], badTime: [], dupes: 0, clamped: 0 };
    const wsP = findSheet(wb, SHEET.pending), wsC = findSheet(wb, SHEET.completed), wsS = findSheet(wb, SHEET.social), wsF = findSheet(wb, SHEET.findings);
    if (!wsP) throw new Error('Previous report me "Takedown Pending" tab nahi mila.');
    if (!wsC) throw new Error('Previous report me "Takedown Completed" tab nahi mila.');
    const HP = headerMap(wsP); const HF = wsF ? headerMap(wsF) : {};
    const P = (name) => HP[nh(name)];
    if (!P('URL')) throw new Error('Takedown Pending tab me URL column nahi mila.');

    /* index Pending rows by URL */
    const idx = new Map();
    for (let r = 2; r <= wsP.rowCount; r++) {
      const u = str(cellVal(wsP.getRow(r).getCell(P('URL')))); if (!u) continue;
      const k = normUrl(u); if (!idx.has(k)) idx.set(k, []); idx.get(k).push(r);
    }
    /* index Findings rows (Channel/Domain, subscribers, views ...) */
    const findIdx = new Map();
    if (wsF && HF['url']) for (let r = 2; r <= wsF.rowCount; r++) { const row = wsF.getRow(r); const u = str(cellVal(row.getCell(HF['url']))); if (u) findIdx.set(str(cellVal(row.getCell(HF['platform/source']))) + '|' + u, r); }
    const doneNorm = new Set(Array.from(info.completedUrls).map(normUrl));

    const compRows = [], socRows = [], removeAt = []; const seen = new Set();
    for (const e of entries) {
      const k = normUrl(e.url);
      if (seen.has(k)) { out.dupes++; continue; } seen.add(k);
      if (e.ts == null) { out.badTime.push(e); continue; }
      const rows = idx.get(k);
      if (!rows) { (doneNorm.has(k) ? out.alreadyDone : out.notFound).push(e.url); continue; }
      for (const r of rows) {
        const row = wsP.getRow(r); const g = (n) => (P(n) ? cellVal(row.getCell(P(n))) : null);
        const plat = str(g('Platform/Source')); const url = str(g('URL'));
        const initiated = stamp(g('Takedown Initiated Date (IST)'), g('Takedown Initiated Time (IST)'));
        let done = e.ts; if (initiated != null && done < initiated) { done = initiated; out.clamped++; }
        const dolby = g('Dolby (Y/N)'); if (dolby !== null && dolby !== '') out.dolbyCarried++;
        const fr = findIdx.get(plat + '|' + url); const fg = (n) => (fr && HF[nh(n)] ? cellVal(wsF.getRow(fr).getCell(HF[nh(n)])) : null);
        const channel = fg('Channel/Domain') || hostOf(url);
        if (SOCIAL_LABELS.has(plat.toLowerCase()) && wsS) {
          socRows.push({ 'Platform/Source': plat, Book: g('Book'), Title: g('Title'), URL: url, 'Coverage by Date - at EoD': g('Coverage by Date - at EoD'), 'Identification Date (IST)': g('Identification Date (IST)'), 'Identification Time (IST)': g('Identification Time (IST)'), 'Channel/Domain': channel, 'Unique Subscribers': fg('Unique Subscribers') ?? '-', Views: fg('Views') ?? '-', Downloads: fg('Downloads') ?? '-', 'File Size': g('File Size'), 'Method of Verification': fg('Method of Verification') || 'Meta Data', 'Host Country/Domain Country': fg('Host Country/Domain Country') || '-', Status: 'Removed', 'Platform Name': platName(url, plat) });
          out.movedSocial++;
        }
        if (!(SOCIAL_LABELS.has(plat.toLowerCase()) && wsS) || BOTH_LABELS.has(plat.toLowerCase())) {
          compRows.push({ 'Platform/Source': plat, Book: g('Book'), Title: g('Title'), URL: url, 'Channel/Domain': channel, 'Coverage by Date - at EoD': g('Coverage by Date - at EoD'), 'Identification Date (IST)': g('Identification Date (IST)'), 'Identification Time (IST)': g('Identification Time (IST)'), 'Takedown Initiated Date (IST)': g('Takedown Initiated Date (IST)'), 'Takedown Initiated Time (IST)': g('Takedown Initiated Time (IST)'), 'Takedown Completed Date (IST)': D(done), 'Takedown Completed Time (IST)': RC.fmtTime(done), 'Takedown Duration (TAT)': initiated == null ? 0 : Math.max(0, Math.floor((done - initiated) / 3600e3)), 'Takedown Unit (TAT)': 'Hrs', 'File Size': g('File Size'), 'Dolby (Y/N)': dolby === '' ? null : dolby });
          out.movedCompleted++;
        }
        removeAt.push({ r, key: plat + '|' + url, url });
      }
    }
    out.moved = removeAt.length;
    if (!removeAt.length) return out;

    appendRows(wsC, compRows, fallbackFor([4]));
    if (wsS) appendRows(wsS, socRows, fallbackFor([4]));

    /* remove moved rows from Pending (bottom -> top) and renumber Sr. No */
    removeAt.map((x) => x.r).sort((a, b) => b - a).forEach((r) => wsP.spliceRows(r, 1));
    const sc = P('Sr. No') || 1; let n = 0;
    for (let r = 2; r <= wsP.rowCount; r++) { const row = wsP.getRow(r); if (str(cellVal(row.getCell(P('URL'))))) { n++; row.getCell(sc).value = n; } }

    /* keep the in-memory index in sync so the normal append step does not re-add / double count */
    const movedUrls = new Set(removeAt.map((x) => x.url));
    removeAt.forEach((x) => { info.keys.pending.delete(x.key); delete info.pendingDolby[x.key]; });
    info.pendingUrls = info.pendingUrls.filter((u) => !movedUrls.has(u));
    compRows.forEach((o) => { info.keys.completed.add(o['Platform/Source'] + '|' + o.URL); info.completedUrls.add(o.URL); });
    socRows.forEach((o) => info.keys.social.add('|' + o.URL));
    return out;
  }

  /* Summary totals only (used when there are no new data files, just a removal update) */
  function refreshTotalsOnly(info, o) {
    o = o || {};
    return refreshSummary(info, { S: { since: o.since == null ? null : o.since }, added: { findings: 0, completed: 0 }, newComp: [], newPend: [], newFind: [], removalOnly: true, remMoved: o.remMoved == null ? 0 : o.remMoved });
  }

  /* =====================================================================
     FINAL PASS on the whole workbook (old + new rows), run once before saving
       1. Takedown Pending  -> "Takedown Pending Duration" recalculated up to the report time (rounded hrs, never 0 -> min 1)
       2. Delisting         -> "Time To Delist (HRS)" column formatted [hh]:mm:ss
       3. Platform columns  -> every word starts with an uppercase letter
     ===================================================================== */
  function finalizeReport(info, refTs) {
    const wb = info.wb; const out = { durations: 0, delistFmt: 0, platform: 0 };

    const wsP = findSheet(wb, SHEET.pending);
    if (wsP) {
      const H = headerMap(wsP); const cD = H['takedown pending duration']; const cDt = H['takedown initiated date (ist)']; const cTm = H['takedown initiated time (ist)'];
      if (cD && cDt) for (let r = 2; r <= wsP.rowCount; r++) {
        const row = wsP.getRow(r); const ini = stamp(cellVal(row.getCell(cDt)), cTm ? cellVal(row.getCell(cTm)) : null);
        if (ini == null) continue;
        row.getCell(cD).value = Math.max(1, Math.round((refTs - ini) / 3600e3)); out.durations++;
      }
    }

    const wsD = findSheet(wb, SHEET.delist);
    if (wsD) {
      const c = headerMap(wsD)['time to delist (hrs)'];
      if (c) for (let r = 2; r <= wsD.rowCount; r++) {
        const cell = wsD.getRow(r).getCell(c); const v = cell.value; if (v === null || v === undefined || v === '') continue;
        if (typeof v === 'string') { const t = timeOfDay(v); if (t == null) continue; cell.value = t / 86400000; }
        cell.numFmt = DELIST_FMT; out.delistFmt++;
      }
    }

    for (const k of ['findings', 'completed', 'pending', 'social']) {
      const ws = findSheet(wb, SHEET[k]); if (!ws) continue; const H = headerMap(ws);
      for (const col of [H['platform/source'], k === 'social' ? H['platform name'] : null]) {
        if (!col) continue;
        for (let r = 2; r <= ws.rowCount; r++) {
          const cell = ws.getRow(r).getCell(col); const v = cell.value;
          if (typeof v === 'string') { const n = cap(v); if (n !== v) { cell.value = n; out.platform++; } }
        }
      }
    }
    return out;
  }

  /* blank template (ExcelJS) the user fills and uploads back */
  function buildRemovalTemplate(ExcelJS) {
    const wb = new ExcelJS.Workbook(); wb.creator = 'Report Generator';
    const ws = wb.addWorksheet('Removal Update', { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = [{ width: 110 }, { width: 28 }];
    ws.getColumn(1).numFmt = '@'; ws.getColumn(2).numFmt = 'dd-mm-yyyy hh:mm AM/PM';
    const h = ws.addRow(['URL', 'Removal Time (IST)']); h.height = 20;
    h.eachCell((c) => { c.numFmt = 'General'; c.font = { name: 'Calibri', size: 12, bold: true }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA6C9EC' } }; c.border = FB_BORDER; c.alignment = { horizontal: 'center', vertical: 'middle' }; });
    const ins = wb.addWorksheet('Instructions'); ins.getColumn(1).width = 120;
    ['How to use', '1. "Removal Update" sheet me column A me Takedown Pending wala URL paste karo.', '2. Column B me us URL ka Removal Time (IST) daalo - e.g. 15-10-2026 02:30 PM (Excel date-time cell ya text dono chalte hain).', '3. Sirf yehi 2 columns chahiye - baaki details report se apne aap aa jayengi (Dolby bhi).', '4. Is file ko Report Generator me "Pending removal update" box me upload karo.', 'Note: Jo URL Takedown Pending me nahi milega, wo skip hoga aur log me dikhega.']
      .forEach((t, i) => { const c = ins.addRow([t]).getCell(1); c.font = { name: 'Calibri', size: 12, bold: i === 0 }; });
    return wb;
  }

  RC._h = { nh, findSheet, headerMap, cellVal, str, cap, appendRows, fallbackFor, hostOf, baseOf, normUrl, timeOfDay, dateOnly };
  RC.finishSummary = finishSummary; RC.finalizeReport = finalizeReport; RC.loadPrevious = loadPrevious; RC.appendToPrevious = appendToPrevious; RC._hostOf = hostOf; RC._baseOf = baseOf;
  RC.parseRemovalTime = parseRemovalTime; RC.parseRemovalFile = parseRemovalFile; RC.applyRemovalUpdate = applyRemovalUpdate; RC.refreshTotalsOnly = refreshTotalsOnly; RC.buildRemovalTemplate = buildRemovalTemplate;
})(typeof self !== 'undefined' ? self : globalThis);
