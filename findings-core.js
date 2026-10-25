/* Findings report: ONE sheet ("Findings", 16 columns, EST) built only from the uploaded data files - no previous report needed.
   Layout copied from the sample "..._Findings_<date>.xlsx": logo on top (A1:P4), header on row 5, data from row 6, frozen panes, zoom 70.
   A row is "Takedown Completed = Y" when the data file itself says it is removed (Removal Status = Dead etc.) or when the optional removal file lists its URL.
   IST -> EST = minus 10h 30m (same fixed shift as the Comet report). */
(function (root) {
  'use strict';
  const RC = root.ReportCore; const H = RC._h;
  const SHIFT = 10.5 * 3600e3;
  const pad = (n) => String(n).padStart(2, '0');
  const dayD = (ts) => new Date(RC.dayStart(ts));
  const tsec = (ts) => {
    const m = ((ts % 86400000) + 86400000) % 86400000; let h = Math.floor(m / 3600000);
    const mi = Math.floor((m % 3600000) / 60000), s = Math.floor((m % 60000) / 1000); const ap = h >= 12 ? 'PM' : 'AM'; h = h % 12 || 12;
    return pad(h) + ':' + pad(mi) + ':' + pad(s) + ' ' + ap;
  };
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const HEAD = ['Sr. No', 'Platform', 'Title', 'URL', 'Infringement Type', 'Identification Date EST', 'Identification Time EST', 'Domain', 'Channel URL', 'Subscriber', 'Views',
    'Takedown Initiated Date (EST)', 'Takedown Initiated Time (EST)', 'Takedown Completed (Y/N)', 'Takedown Completion Date (EST)', 'Takedown Completion Time (EST)'];
  const WIDTH = [11.67, 26.08, 51.58, 21.67, 20.33, 20.42, 34.83, 24, 34, 17.42, 26.75, 26.83, 24.75, 29.5, 29.58, 30];
  const LEFT = [3, 4];                                   /* Title, URL */
  const DATE_FMT = 'yyyy/mm/dd', NUM_FMT = '0_);[Red]\\(0\\)';
  const HEAD_FILL = 'FF163E64';                          /* theme dark-blue (+10%) of the sample header */
  const BORDER = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

  /* file name like "101626_Findings_15_October.xlsx": report date (MMDDYY) + data date (the "since" day, else report day) */
  function findingsFileName(refTs, sinceTs) {
    const r = new Date(refTs), d = new Date(sinceTs == null ? refTs : sinceTs);
    return pad(r.getUTCMonth() + 1) + pad(r.getUTCDate()) + String(r.getUTCFullYear()).slice(-2) + '_Findings_' + d.getUTCDate() + '_' + MONTHS[d.getUTCMonth()] + '.xlsx';
  }

  /* optional removal entries [{url, ts(IST wall clock)}] complete matching rows that are still pending */
  function applyRemoval(recs, entries) {
    const out = { total: entries.length, moved: 0, notFound: [], alreadyDone: [], badTime: [], dupes: 0, clamped: 0 };
    const idx = new Map(); recs.forEach((x) => { const k = H.normUrl(x.url); if (!idx.has(k)) idx.set(k, []); idx.get(k).push(x); });
    const seen = new Set();
    for (const e of entries) {
      const k = H.normUrl(e.url);
      if (seen.has(k)) { out.dupes++; continue; } seen.add(k);
      if (e.ts == null) { out.badTime.push(e); continue; }
      const rows = idx.get(k); if (!rows) { out.notFound.push(e.url); continue; }
      let did = false;
      for (const x of rows) {
        if (x.done) continue;
        let t = e.ts; if (t < x.initiated) { t = x.initiated; out.clamped++; }
        x.done = true; x.completed = t; did = true; out.moved++;
      }
      if (!did) out.alreadyDone.push(e.url);
    }
    return out;
  }

  function dataRow(x, n) {
    /* Domain vs Channel URL: web rows carry a domain, social / telegram rows a channel link */
    const dm = x.domain == null || x.domain === '' ? '-' : String(x.domain);
    const web = x.kind === 'Linking' || x.kind === 'Host' || /^market place$/i.test(x.platform);
    let domain = '-', channel = '-';
    if (dm !== '-') { if (web) domain = dm; else if (/^https?:/i.test(dm)) channel = dm; else domain = dm; }
    const nv = (v) => (typeof v === 'number' ? v : (v === null || v === undefined || v === '' || v === '-' ? '-' : (isFinite(Number(v)) ? Number(v) : String(v))));
    const ini = x.initiated - SHIFT, comp = x.done && x.completed != null ? x.completed - SHIFT : null;
    return [n, H.cap(x.platform), x.title, x.url, x.inf || '-', dayD(x.ident - SHIFT), tsec(x.ident - SHIFT), domain, channel, nv(x.subs), nv(x.views),
      dayD(ini), tsec(ini), comp == null ? 'N' : 'Y', comp == null ? null : dayD(comp), comp == null ? null : tsec(comp)];
  }

  function buildFindings(ExcelJS, res, opts) {
    const o = opts || {};
    const recs = res.findings.slice();
    const rem = o.removal ? applyRemoval(recs, o.removal) : null;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Findings', { views: [{ state: 'frozen', ySplit: 5, topLeftCell: 'A6', activeCell: 'A6', zoomScale: 70, zoomScaleNormal: 70 }] });
    ws.properties.defaultRowHeight = 15; ws.properties.defaultColWidth = 11.08;
    WIDTH.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    ws.mergeCells('A1:P4');
    if (root.FINDINGS_LOGO) {
      const id = wb.addImage({ base64: root.FINDINGS_LOGO.replace(/^data:image\/png;base64,/, ''), extension: 'png' });
      ws.addImage(id, { tl: { col: 0.01, row: 0.1 }, ext: { width: 176, height: 52.5 } });
    }
    const hr = ws.getRow(5);
    HEAD.forEach((h, i) => {
      const c = hr.getCell(i + 1); c.value = h;
      c.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEAD_FILL } };
      c.alignment = { horizontal: 'center', vertical: 'middle' }; c.border = BORDER;
    });
    hr.commit();
    let n = 0, comp = 0;
    for (const x of recs) {
      n++; const vals = dataRow(x, n); if (vals[13] === 'Y') comp++;
      const row = ws.getRow(5 + n); row.height = 15.6;
      vals.forEach((v, i) => {
        const c = row.getCell(i + 1); c.value = v;
        c.font = { name: 'Calibri', size: i === 1 ? 12 : 11 }; c.border = BORDER;
        c.alignment = LEFT.includes(i + 1) ? { horizontal: 'left' } : { horizontal: 'center' };
        if (v instanceof Date) c.numFmt = DATE_FMT; else if (typeof v === 'number' && (i === 9 || i === 10)) c.numFmt = NUM_FMT;
      });
      row.commit();
    }
    const byPlat = {}; recs.forEach((x) => { const p = H.cap(x.platform); byPlat[p] = (byPlat[p] || 0) + 1; });
    return { wb, rem, stats: { total: n, completed: comp, pending: n - comp, byPlat } };
  }

  RC.buildFindings = buildFindings; RC.findingsFileName = findingsFileName;
})(typeof self !== 'undefined' ? self : globalThis);
