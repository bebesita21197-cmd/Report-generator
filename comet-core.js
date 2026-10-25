/* Comet report (single "List" sheet, one row per URL, EST columns).
   Previous Comet report is loaded as-is (ExcelJS); only NEW rows are appended at the bottom and
   removal-update URLs are completed in place. Old rows / widths / styles are never touched otherwise.
   IST -> EST = minus 10 hours 30 minutes (fixed). */
(function (root) {
  'use strict';
  const RC = root.ReportCore; const H = RC._h;
  const SHIFT = 10.5 * 3600e3;
  const pad = (n) => String(n).padStart(2, '0');
  const dayD = (ts) => new Date(RC.dayStart(ts));
  /* "hh:mm:ss AM" text from an IST-wall-clock / EST-wall-clock timestamp */
  const tsec = (ts) => {
    const m = ((ts % 86400000) + 86400000) % 86400000; let h = Math.floor(m / 3600000);
    const mi = Math.floor((m % 3600000) / 60000), s = Math.floor((m % 60000) / 1000); const ap = h >= 12 ? 'PM' : 'AM'; h = h % 12 || 12;
    return pad(h) + ':' + pad(mi) + ':' + pad(s) + ' ' + ap;
  };
  const col = (info, name) => info.M[H.nh(name)];
  const REQUIRED = ['Sr. No', 'Source', 'URL', 'Takedown Initiated Date (EST)', 'Takedown Initiated Time (EST)', 'Takedown Pending  (Y/N)', 'Takedown Completed (Y/N)', 'Takedown Completion Date (EST)', 'Takedown Completion Time (EST)', 'TAT Hours', 'Reason Pending'];

  async function loadComet(ExcelJS, buf) {
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
    const ws = H.findSheet(wb, 'List') || wb.worksheets[0];
    if (!ws) throw new Error('Comet file me koi sheet nahi mili.');
    const info = { wb, ws, M: H.headerMap(ws), urls: new Map(), domExact: {}, domBase: {}, counts: { total: 0, pending: 0 } };
    const miss = REQUIRED.filter((n) => !col(info, n));
    if (miss.length) throw new Error('Ye Comet report ka format nahi lagta - columns missing: ' + miss.join(', '));
    const cUrl = col(info, 'URL'), cPend = col(info, 'Takedown Pending  (Y/N)'), cRes = col(info, 'Reason Pending'), cHsp = col(info, 'HSP Name'), cIp = col(info, 'HSP IP Address');
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r); const url = H.str(H.cellVal(row.getCell(cUrl))); if (!url) continue;
      const k = H.normUrl(url); if (!info.urls.has(k)) info.urls.set(k, []); info.urls.get(k).push(r);
      info.counts.total++;
      if (H.str(H.cellVal(row.getCell(cPend))).toUpperCase() === 'Y') {
        info.counts.pending++;
        const host = H.hostOf(url); if (!host) continue;
        const rec = { reason: cRes ? H.cellVal(row.getCell(cRes)) : null, hsp: cHsp ? H.cellVal(row.getCell(cHsp)) : null, ip: cIp ? H.cellVal(row.getCell(cIp)) : null };
        if (!(host in info.domExact)) info.domExact[host] = rec;           /* first placed wins (same rule as Live report) */
        const b = H.baseOf(host); if (!(b in info.domBase)) info.domBase[b] = rec;
      }
    }
    return info;
  }

  /* ---- new URLs -> new rows (all Pending: removal file completes them later) ---- */
  function cometRow(info, x, dom) {
    const host = H.hostOf(x.url); let p = info.domExact[host]; if (p) dom.exact++; else { p = info.domBase[H.baseOf(host)]; if (p) dom.base++; else dom.fallback++; }
    const dash = (v) => (v === null || v === undefined || v === '' ? '-' : v);
    const reason = x.blankReason ? '-' : dash(p ? p.reason : x.reason);
    const hsp = dash(p ? p.hsp : x.hsp), ip = dash(p ? p.ip : x.hspIp);
    /* Domain vs Channel Name: web / marketplace rows carry a domain, social / telegram rows a channel URL */
    const dm = x.domain == null || x.domain === '' ? '-' : String(x.domain);
    const web = x.kind === 'Linking' || x.kind === 'Host' || /^market place$/i.test(x.platform);
    let domain = '-', channel = '-';
    if (dm !== '-') { if (web) domain = dm; else if (/^https?:/i.test(dm)) channel = dm; else domain = dm; }
    /* File Size + Unit: "474.53 MB" -> 474.53 | MB ; text (Streaming / Signpost / Paid) -> text | - */
    const fsz = String(x.fileSize == null ? '' : x.fileSize).trim(); const m = fsz.match(/^([\d.,]+)\s*(KB|MB|GB|TB)s?$/i);
    const views = typeof x.views === 'number' ? x.views : (x.views === null || x.views === undefined || x.views === '' || x.views === '-' ? '-' : (isFinite(Number(x.views)) ? Number(x.views) : String(x.views)));
    const ident = x.ident, ini = x.initiated;
    return {
      'Source': H.cap(x.platform), 'Book #': x.book, 'Title': x.title, 'URL': x.url, 'Infringement Type': x.inf || '-',
      'Report Date': dayD(ident), 'Identification Date IST': dayD(ident), 'Identification Time IST': tsec(ident),
      'Identification Date EST': dayD(ident - SHIFT), 'Identification Time EST': tsec(ident - SHIFT),
      'Domain': domain, 'Channel Name': channel, 'Views': views, 'Downloads': '-',
      'File Size': m ? Number(m[1].replace(/,/g, '')) : (fsz || '-'), 'Unit': m ? m[2].toUpperCase() : '-',
      'Dolby (Y/N)': null, 'Host Country': dash(x.country),
      'Takedown Initiated (Y/N)': 'Y', 'Takedown Initiated Date (EST)': dayD(ini - SHIFT), 'Takedown Initiated Time (EST)': tsec(ini - SHIFT),
      'Takedown Pending  (Y/N)': 'Y', 'Reason Pending': reason, 'HSP Name': hsp, 'HSP IP Address': ip,
      'Takedown Completed (Y/N)': 'N', 'Takedown Completion Date (EST)': '-', 'Takedown Completion Time (EST)': '-', 'TAT Hours': '-',
      'Counter Notice (Y/N)': 'N', 'Reason For Counter Notice': '-',
    };
  }

  function appendComet(info, res) {
    const seen = new Set(); const dom = { exact: 0, base: 0, fallback: 0 };
    const fresh = res.findings.filter((x) => { const k = H.normUrl(x.url); if (info.urls.has(k) || seen.has(k)) return false; seen.add(k); return true; });
    const rows = fresh.map((x) => cometRow(info, x, dom));
    H.appendRows(info.ws, rows, H.fallbackFor([4, 5]), { 'views': 'General', 'file size': 'General' });
    info.counts.total += rows.length; info.counts.pending += rows.length;
    /* register the new rows so a removal update applied afterwards can find them */
    const cUrl = col(info, 'URL');
    for (let r = 2; r <= info.ws.rowCount; r++) { const u = H.str(H.cellVal(info.ws.getRow(r).getCell(cUrl))); if (!u) continue; const k = H.normUrl(u); const l = info.urls.get(k) || []; if (!l.includes(r)) { l.push(r); info.urls.set(k, l); } }
    return { added: rows.length, skipped: res.findings.length - rows.length, dom };
  }

  /* ---- removal file (URL | Removal Time IST) -> Pending Y becomes Completed Y ---- */
  function applyCometRemoval(info, entries) {
    const ws = info.ws; const out = { total: entries.length, moved: 0, notFound: [], alreadyDone: [], badTime: [], dupes: 0, clamped: 0 };
    const cP = col(info, 'Takedown Pending  (Y/N)'), cC = col(info, 'Takedown Completed (Y/N)'), cD = col(info, 'Takedown Completion Date (EST)'), cT = col(info, 'Takedown Completion Time (EST)'), cTat = col(info, 'TAT Hours');
    const cID = col(info, 'Takedown Initiated Date (EST)'), cIT = col(info, 'Takedown Initiated Time (EST)'), cRes = col(info, 'Reason Pending');
    const seen = new Set();
    for (const e of entries) {
      const k = H.normUrl(e.url);
      if (seen.has(k)) { out.dupes++; continue; } seen.add(k);
      if (e.ts == null) { out.badTime.push(e); continue; }
      const rows = info.urls.get(k); if (!rows) { out.notFound.push(e.url); continue; }
      let did = false;
      for (const r of rows) {
        const row = ws.getRow(r);
        if (H.str(H.cellVal(row.getCell(cP))).toUpperCase() !== 'Y') continue;
        const a = H.dateOnly(H.cellVal(row.getCell(cID))), b = H.timeOfDay(H.cellVal(row.getCell(cIT)));
        const ini = a == null ? null : a + (b || 0);
        let done = e.ts - SHIFT;                                   /* IST -> EST */
        if (ini != null && done < ini) { done = ini; out.clamped++; }
        const mins = ini == null ? 0 : Math.floor((done - ini) / 60000);
        row.getCell(cP).value = 'N'; row.getCell(cC).value = 'Y';
        if (cRes) row.getCell(cRes).value = '-';                  /* takedown complete -> nothing is pending any more */
        const dc = row.getCell(cD); dc.value = dayD(done); dc.numFmt = 'yyyy/mm/dd';
        row.getCell(cT).value = tsec(done);
        row.getCell(cTat).value = mins < 60 ? '0:' + pad(mins) : Math.floor(mins / 60);     /* under 1 hour -> 0:MM like the existing file */
        did = true; out.moved++; info.counts.pending--;
      }
      if (!did) out.alreadyDone.push(e.url);
    }
    return out;
  }

  /* every row whose Takedown Completed = Y gets Reason Pending = "-" (new completions + rows that were already completed in the previous file) */
  function dashCompletedReasons(info) {
    const ws = info.ws, cU = col(info, 'URL'), cC = col(info, 'Takedown Completed (Y/N)'), cRes = col(info, 'Reason Pending'); let n = 0;
    if (!cRes) return 0;
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r); if (!H.str(H.cellVal(row.getCell(cU)))) continue;
      if (H.str(H.cellVal(row.getCell(cC))).toUpperCase() !== 'Y') continue;
      const cell = row.getCell(cRes); if (H.str(H.cellVal(cell)) !== '-') { cell.value = '-'; n++; }
    }
    return n;
  }

  RC.COMET_REQUIRED = REQUIRED; RC.dashCompletedReasons = dashCompletedReasons;
  RC.loadComet = loadComet; RC.appendComet = appendComet; RC.applyCometRemoval = applyCometRemoval;
})(typeof self !== 'undefined' ? self : globalThis);
