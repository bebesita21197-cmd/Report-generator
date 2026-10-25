/* Report core: parsing + business rules + workbook writer. Works in browser (global ReportCore) and Node. */
(function (root) {
  'use strict';
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const BOOKS = [[/sorcerer|philosopher/i,'Book 1'],[/chamber/i,'Book 2'],[/prisoner/i,'Book 3'],[/goblet/i,'Book 4'],[/phoenix/i,'Book 5'],[/half[- ]?blood/i,'Book 6'],[/deathly/i,'Book 7']];
  const PLATFORM_RANK = { Linking: 0, Host: 1, Telegram: 2, 'Social Media & UGC': 3 };
  const PLATFORM_NAMES = { 'bilibili': 'BiliBili', 'vk.ru': 'VK', 'vk.com': 'VK', 'ximalaya.com': 'Ximalaya', 'chomikuj': 'Chomikuj', 'spotify': 'Spotify', 'archive.org': 'Archive.org', 'soundcloud.com': 'SoundCloud', 'youtube': 'YouTube', 'yumpu.com': 'Yumpu' };
  const PLATFORM_COUNTRY = { 'BiliBili': 'China', 'VK': 'Russia', 'Ximalaya': 'China', 'Chomikuj': 'Poland', 'Spotify': 'Sweden', 'SoundCloud': 'Germany', 'Archive.org': 'United States', 'YouTube': 'United States', 'Yumpu': 'Switzerland' };

  /* ---------- New-format platform exports (Facebook / Instagram / Market Place / YouTube / Telegram / Twitter) ----------
     Each platform file is detected by its header signature. Edit the column names here if an export ever changes.
     ident   : columns tried in order for Identification date/time ("Discovery Done At" first, "URL Upload Date" only if that is empty/missing)
     init    : Takedown Initiated time        removal : Takedown Completed time (only used when Removal Status = Dead)
     kind    : 'Social' -> removed rows go to BOTH "Major Social Platform Take down" and "Takedown Completed";  'Telegram' -> stay in Takedown Completed/Pending */
  const PLATFORM_SPECS = {
    facebook: { label: 'Facebook', reason: '', kind: 'Social', sig: ['Video URL', 'Profile URL', 'Subscriber Count'], url: 'Video URL', channel: (r) => r['Profile URL'], subs: 'Subscriber Count', views: 'Views', ident: ['Discovery Done At', 'URL Upload Date'], init: ['Enforcement Time', 'Enforcement Done At'], status: 'Removal Status', removal: ['Takedown Time', 'Removal Time'], country: 'United States' },
    instagram: { label: 'Instagram', reason: '', kind: 'Social', sig: ['Video URL', 'Profile URL', 'URLUploadDate'], url: 'Video URL', channel: (r) => r['Profile URL'], subs: 'Followers Count', views: 'Views', ident: ['Discovery Done At', 'URLUploadDate', 'URL Upload Date'], init: ['Enforcement Time', 'Enforcement Done At'], status: 'Removal Status', removal: ['Takedown Time', 'Removal Time'], country: 'United States' },
    marketplace: { label: 'Market Place', kind: 'Social', sig: ['Seller Name', 'Shop URL'], url: 'Source URL', size: 'Paid', channel: (r) => r['Domain'] || r['Shop URL'] || r['Seller Name'], subs: null, views: null, ident: ['Discovery Done At'], init: ['Enforcement Done At'], status: 'Removal Status', removal: ['Removal Time'], country: (r) => { const d = String(r['Domain'] || r['Source URL'] || '').toLowerCase(); if (/ebay|amazon|etsy|walmart/.test(d)) return 'United States'; if (/aliexpress|alibaba|temu|dhgate|taobao/.test(d)) return 'China'; return r['Seller Country'] || '-'; } },
    youtube: { label: 'YouTube', reason: '', kind: 'Social', sig: ['Video Id', 'Channel Id'], url: 'Source URL', channel: (r) => (r['Channel Id'] ? 'https://www.youtube.com/channel/' + r['Channel Id'] : r['Channel Name']), subs: 'Subscriber Count', views: 'View Count', ident: ['Discovery Done At', 'URL Upload Date'], init: ['Enforcement Done At'], status: 'Removal Status', removal: ['Removal Time'], country: 'United States' },
    telegram: { label: 'Telegram', kind: 'Telegram', sig: ['Channel URL', 'GroupType'], url: 'Source URL', channel: (r) => r['Channel URL'], subs: 'Subscribers', views: 'Views', ident: ['Discovery Done At', 'URL Upload Date'], init: ['Enforcement Done At'], status: 'Removal Status', removal: ['Removal Time'], country: 'Russia' },
    twitter: { label: 'X', reason: '', kind: 'Social', sig: ['source_url_link', 'ProfileURL'], url: 'source_url_link', channel: (r) => r['ProfileURL'], subs: 'follower_count', views: 'view_count', ident: ['Discovery Done At'], init: ['Enforcement Time', 'Enforcement Done At'], status: 'Removal Status', removal: ['Removal Time'], country: 'United States' },
  };

  const infOf = (r) => { for (const k of ['Infringement Type', 'Infringment Type', 'InfringementTypeName']) { const v = r[k]; if (v != null && String(v).trim() !== '') return String(v).trim(); } return ''; };
  /* file size given in the raw platform file (if any) - otherwise "Streaming" */
  /* "1.5 GBs" -> "1.5 GB", "20 MBs" -> "20 MB" (also KBs / TBs) - only for pure size strings */
  const SIZE_S_RE = /^(\s*[\d.,]+\s*)(KB|MB|GB|TB)s(\s*)$/i;
  const fixSize = (v) => (typeof v === 'string' && SIZE_S_RE.test(v) ? v.replace(SIZE_S_RE, (m, a, u, z) => a + u.toUpperCase() + z).trim() : v);
  const sizeOf = (r) => { for (const k of Object.keys(r)) { if (/^(file ?size|size)$/i.test(k.trim())) { const v = r[k]; if (v != null && String(v).trim() !== '' && String(v).trim() !== '-') return fixSize(String(v).trim()); } } return ''; };
  const norm = (h) => String(h == null ? '' : h).replace(/\u00a0/g, ' ').trim();
  const blank = (v) => v === null || v === undefined || (typeof v === 'string' && (v.trim() === '' || v.trim() === '-'));
  const pad = (n) => String(n).padStart(2, '0');
  const cap = (v) => (typeof v === 'string' ? v.replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase()) : v);

  /* ---------- date helpers (all values are "IST wall-clock", stored as UTC-based ms so nothing shifts) ---------- */
  function toTs(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return v.getTime();
    if (typeof v === 'number') return Math.round((v - 25569) * 86400000);
    const s = String(v).trim();
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/);
    if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0));
    const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m2) return Date.UTC(+m2[1], +m2[2] - 1, +m2[3]);
    return null;
  }
  const dayStart = (ts) => Math.floor(ts / 86400000) * 86400000;
  function fmtTime(ts) {
    const d = new Date(ts); let h = d.getUTCHours(); const ap = h >= 12 ? 'PM' : 'AM'; h = h % 12 || 12;
    return pad(h) + ':' + pad(d.getUTCMinutes()) + ' ' + ap;
  }
  function coverage(ts) { const d = new Date(ts); return pad(d.getUTCDate()) + ' ' + MONTHS[d.getUTCMonth()] + ': Coverage'; }
  const dayLabel = (ts) => { const d = new Date(ts); return d.getUTCDate() + ' ' + MONTHS[d.getUTCMonth()]; };

  /* ---------- parsing ---------- */
  function parseWorkbook(XLSX, buf, keyword) {
    const wb = XLSX.read(buf, { type: 'array', dense: true, cellDates: false, cellNF: false, cellText: false, cellStyles: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    if (!aoa.length) throw new Error('Empty sheet');
    let headers = aoa[0].map(norm);
    if (!headers.includes('Asset Name') && headers.includes('Asset')) headers = headers.map((h) => (h === 'Asset' ? 'Asset Name' : h));   /* Facebook export calls it "Asset" */
    const idx = {}; headers.forEach((h, i) => { if (!(h in idx)) idx[h] = i; });
    const specId = Object.keys(PLATFORM_SPECS).find((k) => PLATFORM_SPECS[k].sig.every((h) => h in idx));   /* must be checked first: Facebook also has Video Title + Removal Status */
    const isSocial = !specId && 'Video Title' in idx && 'Removal Status' in idx;
    let role;
    if (specId) role = 'platform';
    else if (isSocial) role = 'social';
    else {
      let n = 0, withSrc = 0; const si = idx['Source URL'];
      for (let r = 1; r < Math.min(aoa.length, 600); r++) { n++; if (!blank(aoa[r][si])) withSrc++; }
      role = n && withSrc / n > 0.5 ? 'source' : 'infringing';
    }
    const ai = idx['Asset Name']; const kw = (keyword || '').trim().toLowerCase();
    const rows = []; let total = 0;
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r]; total++;
      const asset = String(row[ai] == null ? '' : row[ai]);
      if (kw && !asset.toLowerCase().includes(kw)) continue;
      const o = specId ? { __spec: specId } : {}; for (let c = 0; c < headers.length; c++) { const v = row[c]; if (v !== null && v !== undefined && v !== '') o[headers[c]] = v; }
      rows.push(o);
    }
    return { role, rows, total, headers, spec: specId ? PLATFORM_SPECS[specId].label : null };
  }

  /* ---------- helpers for business rules ---------- */
  const bookOf = (title) => { for (const [re, b] of BOOKS) if (re.test(title)) return b; return '-'; };
  const DOLBY_RE = /atmos|dolby|eac3|e-ac-3|ddp|dd\+/i;
  const txt = (v) => (blank(v) ? '' : String(v).trim());
  function fileSize(r) { const s = txt(r['Note 1']) || txt(r['File Size']); return s ? fixSize(s) : 'Streaming'; }
  function normalizeUrl(u) { return String(u).trim(); }
  const isDead = (s) => /^dead$/i.test(txt(s));

  /* rows of the selected data that have NO Enforcement time, or whose Dolby column (any header containing "dolby") exists but is blank */
  function checkMissing(data) {
    const out = [];
    const dolbyBlank = (r) => { const ks = Object.keys(r).filter((k) => /dolby/i.test(k)); return ks.length > 0 && ks.every((k) => blank(r[k])); };
    const add = (label, url, r, enfKeys) => {
      const miss = []; if (!enfKeys.some((k) => toTs(r[k]) != null)) miss.push('Enforcement time'); if (dolbyBlank(r)) miss.push('Dolby status');
      if (miss.length) out.push({ label, url: String(url == null ? '' : url), miss });
    };
    (data.infringing || []).forEach((r) => { const u = txt(r['Infringing URL']); if (u) add('Infringing', u, r, ['Enforcement Done At (Infringing)']); });
    (data.source || []).forEach((r) => { const u = txt(r['Source URL']); if (u) add('Source', u, r, ['Enforcement Done At (Source)']); });
    (data.social || []).forEach((r) => { const u = txt(r['Source URL']); if (u) add('Social', u, r, ['Enforcement Done At']); });
    (data.platform || []).forEach((r) => { const sp = PLATFORM_SPECS[r.__spec]; if (!sp) return; const u = txt(r[sp.url]); if (u) add(sp.label, u, r, sp.init); });
    return out;
  }

  function build(data, opts) {
    const o = Object.assign({ keyword: 'Harry Potter', refTs: Date.now() + 5.5 * 3600e3, sinceTs: null, rng: Math.random }, opts || {});
    const warn = []; const stats = { randomDelist: 0, dedupDropped: 0 };
    const recs = []; const seen = new Set();

    /* country by domain (fallback when a row has no country) */
    const domCountry = {};
    const note = (dom, c) => { if (!dom || !c) return; (domCountry[dom] = domCountry[dom] || {}); domCountry[dom][c] = (domCountry[dom][c] || 0) + 1; };
    const cleanCountry = (r, sfx) => { const hq = txt(r['Domain IP Host Headqauter (' + sfx + ')']); if (hq) return hq; const c = txt(r['Country']); return c && !/^global$/i.test(c) ? c : ''; };
    (data.infringing || []).forEach((r) => note(txt(r['Infringing Domain']), cleanCountry(r, 'Infringing')));
    (data.source || []).forEach((r) => note(txt(r['Source Domain']), cleanCountry(r, 'Source')));
    /* domain level lookups: escalation flag, HSP, IP, server location (rows of same domain share them) */
    const domInfo = {};
    const learn = (dom, r, sfx) => {
      if (!dom) return; const d = (domInfo[dom] = domInfo[dom] || { esc: false, hsp: '', ip: '', hq: '' });
      if (!blank(r['Escalation Notice Sent Date (' + sfx + ')'])) d.esc = true;
      d.hsp = d.hsp || txt(r['Host Service Provider Name (' + sfx + ')']);
      d.ip = d.ip || txt(r['Domain IP Address (' + sfx + ')']);
      d.hq = d.hq || txt(r['Domain IP Host Headqauter (' + sfx + ')']);
    };
    (data.infringing || []).forEach((r) => learn(txt(r['Infringing Domain']), r, 'Infringing'));
    (data.source || []).forEach((r) => learn(txt(r['Source Domain']), r, 'Source'));
    const topCountry = (dom) => { const m = domCountry[dom]; if (!m) return ''; return Object.entries(m).sort((a, b) => b[1] - a[1])[0][0]; };

    function addWeb(r, kind) {
      const sfx = kind === 'Linking' ? 'Infringing' : 'Source';
      const url = normalizeUrl(kind === 'Linking' ? r['Infringing URL'] : r['Source URL']);
      if (!url) return;
      const key = kind + '|' + url; if (seen.has(key)) { stats.dedupDropped++; return; }
      seen.add(key);
      const title = txt(r['Asset Name']);
      const ident = toTs(r['URL Upload Date']) ?? toTs(r['Discovery Done At']);
      if (ident == null) { warn.push('Identification time missing: ' + url); return; }
      let platform = kind;
      if (/(^|\.)t\.me$/i.test(txt(r[kind === 'Linking' ? 'Infringing Domain' : 'Source Domain'])) || /^https?:\/\/t\.me\//i.test(url)) platform = 'Telegram';
      if (r.__platform) platform = r.__platform;      /* platform chosen for the uploaded file (overrides auto-detect, file name is never used) */
      const dom = txt(r[kind === 'Linking' ? 'Infringing Domain' : 'Source Domain']);
      const status = r['DMCA Removal Status (' + sfx + ')'];
      const initiated = toTs(r['Enforcement Done At (' + sfx + ')']) ?? ident;
      let completed = toTs(r['DMCA Removal Done At (' + sfx + ')']) ?? toTs(r['DMCA Removal Updated At (' + sfx + ')']);
      const done = isDead(status);
      if (done && completed == null) { warn.push('Dead but no removal time, used initiated time: ' + url); completed = initiated; }
      if (done && completed < initiated) completed = initiated;
      const di = domInfo[dom] || {}; const escal = !!di.esc;
      recs.push({
        platform, kind, inf: infOf(r), book: bookOf(title), title, url, ident, domain: dom, subs: '-', views: '-', downloads: '-',
        fileSize: fileSize(r), dolbyHit: DOLBY_RE.test(url + ' ' + title + ' ' + txt(r['Matched Title'])), hash: txt(r['Note 2']).toLowerCase(),
        method: 'Downloaded & Validated', country: cleanCountry(r, sfx) || (di.hq) || 'United States',
        initiated, completed: done ? completed : null, done,
        reason: escal ? 'Non-Compliance from Domain owner -  Escalated to HSP' : '-',
        hsp: txt(r['Host Service Provider Name (' + sfx + ')']) || di.hsp || '-', hspIp: txt(r['Domain IP Address (' + sfx + ')']) || di.ip || '-',
        server: txt(r['Domain IP Host Headqauter (' + sfx + ')']) || di.hq || '-',
        delistDone: kind === 'Linking' ? toTs(r['Delisting Done At (Infringing)']) : null,
      });
      if (platform === 'Telegram') { const l = recs[recs.length - 1]; l.method = 'Meta Data'; l.reason = 'Current TAT 72+ Hrs'; l.fileSize = 'Streaming'; }
    }
    (data.infringing || []).forEach((r) => addWeb(r, 'Linking'));
    (data.source || []).forEach((r) => addWeb(r, 'Host'));

    /* Dolby: text match, shared across identical file hashes */
    const hashY = new Set(); recs.forEach((x) => { if (x.dolbyHit && x.hash) hashY.add(x.hash); });
    recs.forEach((x) => { x.dolby = x.dolbyHit || (x.hash && hashY.has(x.hash)) ? 'Y' : 'N'; });

    /* Social */
    (data.social || []).forEach((r) => {
      const url = normalizeUrl(r['Source URL']); if (!url) return;
      const key = 'S|' + url; if (seen.has(key)) { stats.dedupDropped++; return; } seen.add(key);
      const title = txt(r['Asset Name']);
      const ident = toTs(r['Discovery Done At']); if (ident == null) { warn.push('Social: identification time missing ' + url); return; }
      const domRaw = txt(r['Domain']); const pn = cap(PLATFORM_NAMES[domRaw.toLowerCase()] || domRaw || '-');
      const done = isDead(r['Removal Status']);
      const initiated = toTs(r['Enforcement Done At']) ?? ident;
      let completed = toTs(r['Removal Time']); if (done && completed == null) completed = initiated;
      const cty = txt(r['Country']);
      recs.push({
        platform: r.__platform || 'Social Media & UGC', kind: 'Social', inf: infOf(r), book: bookOf(title), title, url, ident, domain: txt(r['Channel/Profile URL']) || '-',
        subs: r['Subscriber Count'] ?? '-', views: r['View Count'] ?? '-', downloads: '-', fileSize: 'Streaming', dolby: 'N',
        method: 'Meta Data', country: PLATFORM_COUNTRY[pn] || (cty && !/^global$/i.test(cty) ? cty : '-'), platformName: r.__platform && r.__platform !== 'Social Media & UGC' ? r.__platform : pn,
        initiated, completed: done ? completed : null, done, reason: 'Current TAT 72+ Hrs', hsp: '-', hspIp: '-', server: '-',
      });
    });

    /* New-format platform files: Facebook / YouTube / Telegram / Twitter(X) */
    (data.platform || []).forEach((r) => {
      const sp = PLATFORM_SPECS[r.__spec]; if (!sp) return;
      const url = normalizeUrl(r[sp.url] == null ? '' : r[sp.url]); if (!url) return;
      const key = 'S|' + url; if (seen.has(key)) { stats.dedupDropped++; return; } seen.add(key);
      const pick = (names) => { for (const n of names) { const t = toTs(r[n]); if (t != null) return t; } return null; };
      const title = txt(r['Asset Name']);
      const ident = pick(sp.ident); if (ident == null) { warn.push(sp.label + ': identification time missing ' + url); return; }
      const done = isDead(r[sp.status]);
      const initiated = pick(sp.init) ?? ident;
      let completed = pick(sp.removal); if (done && completed == null) completed = initiated; if (done && completed < initiated) completed = initiated;
      const ch = sp.channel(r); const dom = blank(ch) ? '-' : String(ch).trim();
      const sel = r.__platform && r.__platform !== 'Social Media & UGC' ? r.__platform : sp.label;
      recs.push({
        platform: r.__platform || sp.label, kind: sp.kind, inf: infOf(r), book: bookOf(title), title, url, ident, domain: dom,
        subs: r[sp.subs] ?? '-', views: r[sp.views] ?? '-', downloads: '-', fileSize: sp.size || sizeOf(r) || 'Streaming', dolby: null, both: sp.kind === 'Social',
        method: 'Meta Data', country: typeof sp.country === 'function' ? sp.country(r) : sp.country, platformName: sel,
        initiated, completed: done ? completed : null, done, reason: sp.reason === undefined ? 'Current TAT 72+ Hrs' : sp.reason, blankReason: sp.reason === '', hsp: '-', hspIp: '-', server: '-',
      });
    });

    /* New data is never classified active/dead here: every new URL is Pending.
       Takedown Completed / Major Social Platform Take down are filled ONLY by the removal-update file. */
    if (o.allPending !== false) recs.forEach((x) => { x.done = false; x.completed = null; });

    /* order + serials */
    recs.sort((a, b) => (dayStart(a.ident) - dayStart(b.ident)) || ((PLATFORM_RANK[a.platform] ?? 9) - (PLATFORM_RANK[b.platform] ?? 9)) || String(a.platform).localeCompare(String(b.platform)) || (a.ident - b.ident));
    const findings = recs;
    const completedList = findings.filter((x) => x.done && (x.kind !== 'Social' || x.both));
    const pendingList = findings.filter((x) => !x.done);
    const socialRemoved = findings.filter((x) => x.kind === 'Social' && x.done);

    /* delisting: only Linking, always "Delisted" */
    const delist = findings.filter((x) => x.kind === 'Linking').map((x) => {
      const req = x.initiated;
      let comp = x.delistDone != null ? Math.floor(x.delistDone / 60000) * 60000 : null;
      if (comp == null || comp < req) { comp = Math.floor(req / 60000) * 60000 + Math.round(10 + o.rng() * 10) * 60000; stats.randomDelist++; }
      return Object.assign({}, x, { delReq: req, delComp: comp });
    });

    const total = findings.length, comp = completedList.length, open = pendingList.length;
    const since = o.sinceTs;
    const newFind = since == null ? null : findings.filter((x) => x.ident >= since).length;
    const newTd = since == null ? null : completedList.filter((x) => x.completed >= since).length;
    return { findings, completedList, pendingList, socialRemoved, delist, summary: { total, comp, open, newFind, newTd, since }, warn, stats, opts: o };
  }

  /* ---------- workbook writer (ExcelJS) ---------- */
  const BORDER = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  const HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA6C9EC' } };
  const SUM_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB5E6A2' } };
  const GLOSS_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E2F3' } };
  const FONT = (b) => ({ name: 'Calibri', size: 12, bold: !!b });

  function addTable(wb, name, headers, widths, rows, o) {
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1, zoomScale: o.zoom || 70, showGridLines: true }], properties: { tabColor: o.tab ? { argb: o.tab } : undefined } });
    ws.columns = headers.map((h, i) => ({ width: widths[i] || 18 }));
    const hr = ws.addRow(headers); hr.height = o.headHeight || 18;
    hr.eachCell((c) => { c.font = FONT(true); c.fill = HEAD_FILL; c.border = BORDER; c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
    const fmts = o.fmts || {}; const left = o.left || [];
    for (const r of rows) {
      const row = ws.addRow(r);
      row.eachCell({ includeEmpty: true }, (c, n) => {
        c.font = FONT(false); c.border = BORDER;
        c.alignment = { horizontal: left.includes(n) ? 'left' : 'center', vertical: 'middle' };
        if (fmts[n]) c.numFmt = fmts[n];
      });
    }
    if (o.filter) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } };
    return ws;
  }

  function writeWorkbook(ExcelJS, res) {
    const { findings, completedList, pendingList, socialRemoved, delist, summary: S, opts } = res;
    const wb = new ExcelJS.Workbook(); wb.creator = 'Report Generator'; wb.created = new Date();
    const D = (ts) => new Date(dayStart(ts));
    const DATEF = 'mm/dd/yyyy';
    wb.addWorksheet('Notes', { views: [{ zoomScale: 70 }] }).getColumn(1).width = 12.5;

    /* Summary (formulas + cached results) */
    const ws = wb.addWorksheet('Summary', { views: [{ zoomScale: 85 }] });
    ws.columns = [17.7, 44.3, 15.3, 41.8, 20.4, 27.0].map((w) => ({ width: w }));
    const sinceTxt = S.since == null ? '-' : fmtTime(S.since) + ' EST - ' + dayLabel(S.since);
    const h = ws.addRow(['Total Infringements', 'Total infringements since prior report (insert time)', 'Total Takedowns', 'Total takedowns since prior report (insert time)', 'Open issues remaining', 'Current Takedown Percentage']);
    h.height = 20.25; h.eachCell((c) => { c.font = FONT(true); c.fill = SUM_FILL; c.border = BORDER; c.alignment = { horizontal: 'center', vertical: 'middle' }; });
    const pct = S.total ? S.comp / S.total : 0;
    const v = ws.addRow([
      { formula: 'COUNT(Findings!A:A)', result: S.total },
      S.since == null ? '-' : pad(S.newFind) + ' (' + sinceTxt + ')',
      { formula: "COUNT('Takedown Completed'!A:A)", result: S.comp },
      S.since == null ? '-' : pad(S.newTd) + ' (' + sinceTxt + ')',
      { formula: "COUNT('Takedown Pending'!A:A)", result: S.open },
      { formula: 'IF(A2=0,0,C2/A2)', result: pct },
    ]);
    v.height = 15.75; v.eachCell((c, n) => { c.font = FONT(true); c.border = BORDER; c.alignment = { horizontal: 'center', vertical: 'middle' }; if (n === 6) c.numFmt = '0%'; });

    /* Findings */
    addTable(wb, 'Findings',
      ['Sr. No', 'Platform/Source', 'Book', 'Title', 'URL', 'Coverage by Date - at EoD', 'Identification Date (IST)', 'Identification Time (IST)', 'Channel/Domain', 'Unique Subscribers', 'Views', 'Downloads', 'File Size', 'Dolby (Y/N)', 'Method of Verification', 'Host Country/Domain Country'],
      [7.5, 21.6, 16.6, 48.6, 87.7, 31.1, 25.6, 24, 54.1, 21.5, 9, 11.6, 14.6, 12, 25.1, 30.6],
      findings.map((x, i) => [i + 1, x.platform, x.book, x.title, x.url, coverage(x.ident), D(x.ident), fmtTime(x.ident), x.domain, x.subs, x.views, x.downloads, x.fileSize, x.dolby, x.method, x.country]),
      { tab: 'FF00B0F0', zoom: 85, filter: true, fmts: { 7: DATEF }, left: [4], headHeight: 36 });

    /* Takedown Completed */
    addTable(wb, 'Takedown Completed',
      ['Sr. No', 'Platform/Source', 'Book', 'Title', 'URL', 'Channel/Domain', 'Coverage by Date - at EoD', 'Identification Date (IST)', 'Identification Time (IST)', 'Takedown Initiated Date (IST)', 'Takedown Initiated Time (IST)', 'Takedown Completed Date (IST)', 'Takedown Completed Time (IST)', 'Takedown Duration (TAT)', 'Takedown Unit\u00a0(TAT)', 'File Size', 'Dolby (Y/N)'],
      [7.1, 17.6, 16.6, 48.6, 106, 35.1, 35.6, 22.6, 24.1, 26.6, 30.5, 28.6, 31.6, 23.6, 18, 18.2, 18.2],
      completedList.map((x, i) => [i + 1, x.platform, x.book, x.title, x.url, x.domain, coverage(x.ident), D(x.ident), fmtTime(x.ident), D(x.initiated), fmtTime(x.initiated), D(x.completed), fmtTime(x.completed), Math.max(0, Math.floor((x.completed - x.initiated) / 3600e3)), 'Hrs', x.fileSize, x.dolby]),
      { zoom: 70, fmts: { 8: DATEF, 10: DATEF, 12: DATEF }, left: [4], headHeight: 36 });

    /* Takedown Pending */
    addTable(wb, 'Takedown Pending',
      ['Sr. No', 'Platform/Source', 'Book', 'Title', 'URL', 'Coverage by Date - at EoD', 'Identification Date (IST)', 'Identification Time (IST)', 'Takedown Initiated Date (IST)', 'Takedown Initiated Time (IST)', 'Takedown Pending Duration', 'Takedown Pending Unit', 'Reason Takedown is Pending:', 'HSP Name', 'HSP IP Address', 'Server Location', 'File Size', 'Dolby (Y/N)'],
      [6.1, 17.6, 11.1, 51.5, 105.5, 23.1, 21.1, 20, 28.7, 27.1, 27.7, 23.5, 48.6, 48.1, 28.1, 18.5, 16, 18.2],
      pendingList.map((x, i) => [i + 1, x.platform, x.book, x.title, x.url, coverage(x.ident), D(x.ident), fmtTime(x.ident), D(x.initiated), fmtTime(x.initiated), Math.max(1, Math.round((opts.refTs - x.initiated) / 3600e3)), 'Hrs', x.reason, x.hsp, x.hspIp, x.server, x.fileSize, x.dolby]),
      { tab: 'FFFFFF00', zoom: 70, fmts: { 7: DATEF, 9: DATEF }, left: [4], headHeight: 36 });

    /* Delisting */
    addTable(wb, 'Delisting',
      ['Sr. No', 'Source/Search Engine', 'Offending URL', 'Book', 'Title', 'URL', 'Coverage by Date - at EoD', 'Delisting Request Date (IST)', 'Delisting Request Time (IST)', 'Delisting Status', 'Delisting Completion Time (IST)', 'Delisting Completion Date (IST)', 'Time To Delist (HRS)'],
      [7.6, 20.6, 23.6, 14, 46.1, 69.2, 27.5, 37.1, 31.6, 22.6, 31.6, 27.5, 18.2],
      delist.map((x, i) => [i + 1, 'Linking', '-', x.book, x.title, x.url, coverage(x.ident), D(x.delReq), fmtTime(x.delReq), 'Delisted', fmtTime(x.delComp), D(x.delComp), (Math.floor(x.delComp / 60000) - Math.floor(x.delReq / 60000)) / 1440]),
      { zoom: 70, fmts: { 8: DATEF, 12: DATEF, 13: '[hh]:mm:ss' }, left: [5], headHeight: 36 });

    /* Major Social Platform Take down */
    addTable(wb, 'Major Social Platform Take down',
      ['Sr. No', 'Platform/Source', 'Book', 'Title', 'URL', 'Coverage by Date - at EoD', 'Identification Date (IST)', 'Identification Time (IST)', 'Channel/Domain', 'Unique Subscribers', 'Views', 'Downloads', 'File Size', 'Method of Verification', 'Host Country/Domain Country', 'Status', 'Platform Name'],
      [7, 22.6, 16.6, 48.6, 43.2, 24.6, 25.1, 26.1, 44.6, 13.1, 9.2, 16.6, 12, 21.6, 21.1, 21, 16.1],
      socialRemoved.map((x, i) => [i + 1, x.platform, x.book, x.title, x.url, coverage(x.ident), D(x.ident), fmtTime(x.ident), x.domain, x.subs, x.views, x.downloads, x.fileSize, x.method, x.country, 'Removed', x.platformName]),
      { zoom: 70, fmts: { 7: DATEF }, left: [4], headHeight: 31.5 });

    /* Glossary */
    const G = root.GLOSSARY; const gs = wb.addWorksheet('Glossary', { views: [{ zoomScale: 70 }] });
    gs.columns = [{ width: 54.1 }, { width: 47 }, { width: 130.1 }];
    G.rows.forEach((r, i) => {
      const row = gs.addRow(r.map((c) => (typeof c === 'string' ? c.replace(/\(Harry Potter\)/g, '(' + (opts.keyword || 'Harry Potter') + ')') : c)));
      row.eachCell({ includeEmpty: true }, (c, n) => {
        c.border = BORDER;
        if (i === 0) { c.font = { name: 'Calibri', size: 11, bold: true }; c.fill = GLOSS_FILL; c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; }
        else { c.font = { name: 'Calibri', size: 11, bold: n < 3 && n !== 3 && n <= 2 }; c.alignment = n === 3 ? { wrapText: true, vertical: 'middle' } : { horizontal: 'center', vertical: 'middle', wrapText: true }; }
      });
    });
    G.merges.forEach((m) => gs.mergeCells(m));
    gs.getRow(1).height = 33;
    return wb;
  }

  /* Comet report file name: MMDD_hhmmPM_New_Comet_Data.xlsx (same time stamp as the Live report) */
  function cometFileName(refTs) {
    const d = new Date(refTs); let h = d.getUTCHours(); const ap = h >= 12 ? 'PM' : 'AM'; h = h % 12 || 12;
    return pad(d.getUTCMonth() + 1) + pad(d.getUTCDate()) + '_' + pad(h) + pad(d.getUTCMinutes()) + ap + '_New_Comet_Data.xlsx';
  }
  function fileName(refTs, keyword) {
    const d = new Date(refTs); let h = d.getUTCHours(); const ap = h >= 12 ? 'PM' : 'AM'; h = h % 12 || 12;
    return pad(d.getUTCMonth() + 1) + pad(d.getUTCDate()) + '_' + pad(h) + pad(d.getUTCMinutes()) + ap + '_IST_' + (keyword || 'Report').replace(/[^\w]+/g, '') + '_LiveReport.xlsx';
  }

  /* Run on every workbook right before it is written. ExcelJS writes NaN / Infinity / invalid dates as <v>NaN</v> and
     keeps control characters / >32767-char strings, and Excel then shows "found a problem with some content" and "repairs"
     (= wipes) cells. Clean all of that + normalise "GBs/MBs" -> "GB/MB". */
  function sanitizeWorkbook(wb) {
    const BAD = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/g; let fixed = 0;
    const clean = (v) => {
      if (typeof v === 'number') return isFinite(v) ? v : (fixed++, '-');
      if (v instanceof Date) return isNaN(v.getTime()) ? (fixed++, null) : v;
      if (typeof v === 'string') { let t = v; if (BAD.test(t)) { t = t.replace(BAD, ''); fixed++; } t = fixSize(t); if (t.length > 32767) { t = t.slice(0, 32767); fixed++; } return t; }
      return v;
    };
    wb.eachSheet((ws) => {
      ws.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (c) => {
          const v = c.value; if (v === null || v === undefined) return;
          if (typeof v === 'object' && !(v instanceof Date)) {
            if (v.formula !== undefined || v.sharedFormula !== undefined) { if (typeof v.result === 'number' && !isFinite(v.result)) { c.value = Object.assign({}, v, { result: 0 }); fixed++; } else if (v.result instanceof Date && isNaN(v.result)) { c.value = Object.assign({}, v, { result: 0 }); fixed++; } }
            else if (v.richText) { v.richText.forEach((t) => { if (typeof t.text === 'string') t.text = clean(t.text); }); }
            else if (typeof v.text === 'string') { v.text = clean(v.text); }
            else if (v.error === undefined && !v.hyperlink) { c.value = null; fixed++; }   /* stray object -> would be written as garbage */
            return;
          }
          const n = clean(v); if (n !== v) c.value = n;
        });
      });
    });
    return fixed;
  }

  root.ReportCore = { checkMissing, sanitizeWorkbook, fixSize, PLATFORM_SPECS, parseWorkbook, build, writeWorkbook, fileName, cometFileName, toTs, dayStart, dayLabel, coverage, fmtTime };
  if (typeof module !== 'undefined' && module.exports) module.exports = root.ReportCore;
})(typeof self !== 'undefined' ? self : globalThis);
