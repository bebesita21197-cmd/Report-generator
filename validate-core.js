/* File format checks for every Browse / drop. Each check throws an Error (Hinglish message) when the file is not in the expected format.
   Nothing here changes any report logic - it only decides whether a picked file is accepted. Works in browser and Node (needs SheetJS passed in). */
(function (root) {
  'use strict';
  const nh = (h) => String(h == null ? '' : h).replace(/ /g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  const LIVE_TABS = ['Findings', 'Takedown Pending', 'Takedown Completed'];

  /* what each platform button expects (shown in the error so the person knows what to pick) */
  const EXPECT = {
    'Linking': 'Infringing URL, Infringing Domain, Asset Name',
    'Host': 'Source URL, Source Domain, Asset Name',
    'Social Media & UGC': 'Source URL, Video Title, Removal Status, Channel/Profile URL, Asset Name',
  };

  /* ---- bytes: real .xlsx is a ZIP ("PK"), old .xls is an OLE container ---- */
  function magic(buf) {
    const b = new Uint8Array(buf, 0, Math.min(8, buf.byteLength || 0));
    if (b.length >= 4 && b[0] === 0x50 && b[1] === 0x4b && (b[2] === 3 || b[2] === 5 || b[2] === 7)) return 'zip';
    if (b.length >= 8 && b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0) return 'ole';
    return 'other';
  }
  /* name check: allowXls=true -> .xlsx or .xls, otherwise only .xlsx */
  function checkName(name, allowXls) {
    const ok = allowXls ? /\.xlsx?$/i.test(name) : /\.xlsx$/i.test(name);
    if (!ok) { const m = String(name).match(/\.([^.]+)$/); throw new Error('Galat file type' + (m ? ' (.' + m[1] + ')' : '') + ' - sirf ' + (allowXls ? '.xlsx / .xls' : '.xlsx') + ' Excel file chalegi.'); }
  }
  function checkBytes(buf, name, allowXls) {
    if (!buf || !buf.byteLength) throw new Error('File khali hai (0 KB).');
    const m = magic(buf);
    if (m === 'zip' || (allowXls && m === 'ole')) return m;
    throw new Error(m === 'ole'
      ? 'Ye purani .xls ya password-protected file hai, asli .xlsx nahi. Excel me "Save As -> .xlsx" karke use karo.'
      : 'Ye asli Excel file nahi hai (extension ' + (/\.xlsx?$/i.test(name) ? '.xlsx' : '') + ' hai lekin andar ka content Excel ka nahi - CSV/text ya corrupt file ho sakti hai).');
  }
  function readTop(XLSX, buf) {                       /* first row of every sheet only */
    try { return XLSX.read(buf, { type: 'array', dense: true, sheetRows: 1, cellDates: false, cellNF: false, cellText: false, cellStyles: false }); }
    catch (e) { throw new Error('Excel file khul nahi paayi (corrupt ya galat format): ' + e.message); }
  }
  const headersOf = (XLSX, ws) => { const a = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false }); return a.length ? a[0].map(nh) : []; };

  /* ---- Previous Live report: Findings + Takedown Pending + Takedown Completed tabs, each with a URL column ---- */
  function validateLive(XLSX, buf) {
    const wb = readTop(XLSX, buf); const names = wb.SheetNames; const low = names.map(nh);
    const missing = LIVE_TABS.filter((t) => !low.includes(nh(t)));
    if (missing.length) {
      const hint = low.includes('list') ? ' Ye file Comet report lag rahi hai (List sheet) - isse "Previous Comet Report" me daalo.' : '';
      throw new Error('Ye Live report ka format nahi hai - tabs missing: ' + missing.join(', ') + '. File me tabs mile: ' + names.join(', ') + '.' + hint);
    }
    const noUrl = LIVE_TABS.filter((t) => !headersOf(XLSX, wb.Sheets[names[low.indexOf(nh(t))]]).includes('url'));
    if (noUrl.length) throw new Error('Live report ke in tabs me "URL" column nahi mila: ' + noUrl.join(', ') + '.');
  }

  /* ---- Previous Comet report: "List" sheet (or first sheet) with all Comet columns ---- */
  function validateComet(XLSX, buf) {
    const wb = readTop(XLSX, buf); const names = wb.SheetNames; const low = names.map(nh);
    const i = low.indexOf('list'); const ws = wb.Sheets[names[i >= 0 ? i : 0]];
    if (!ws) throw new Error('Comet file me koi sheet nahi mili.');
    const heads = headersOf(XLSX, ws); const req = root.ReportCore && root.ReportCore.COMET_REQUIRED ? root.ReportCore.COMET_REQUIRED : [];
    const miss = req.filter((n) => !heads.includes(nh(n)));
    if (miss.length) {
      const hint = LIVE_TABS.every((t) => low.includes(nh(t))) ? ' Ye file Live report lag rahi hai - isse "Previous Live Report" me daalo.' : '';
      throw new Error('Ye Comet report ka format nahi hai - columns missing: ' + miss.join(', ') + '.' + hint);
    }
  }

  /* ---- Removal update file: URL + Removal Time columns and at least one URL row ---- */
  function validateRemoval(XLSX, buf) {
    let aoa;
    try { const wb = XLSX.read(buf, { type: 'array', cellDates: false }); aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null, blankrows: false }); }
    catch (e) { throw new Error('Excel file khul nahi paayi (corrupt ya galat format): ' + e.message); }
    if (!aoa.length) throw new Error('Removal file khali hai.');
    const heads = aoa[0].map(nh);
    let cu = heads.findIndex((h) => h === 'url' || h === 'urls' || h === 'infringing url' || h === 'source url'); if (cu < 0) cu = heads.findIndex((h) => h.includes('url'));
    let ct = heads.findIndex((h) => h.includes('removal') && h.includes('time')); if (ct < 0) ct = heads.findIndex((h) => h.includes('time') || h.includes('date'));
    if (cu < 0 || ct < 0) throw new Error('Removal file ka format galat hai - 2 columns chahiye: "URL" aur "Removal Time (IST)" (' + (cu < 0 ? 'URL column nahi mila' : 'Removal Time column nahi mila') + '). Template download karke use karo.');
    if (!aoa.slice(1).some((r) => r[cu] != null && String(r[cu]).trim() !== '')) throw new Error('Removal file me ek bhi URL row nahi hai.');
  }

  /* ---- New data file for a platform button (parsed = ReportCore.parseWorkbook result) ---- */
  function describe(parsed) {
    if (parsed.spec) return parsed.spec + ' file';
    if (parsed.role === 'social') return 'Social Media & UGC file';
    const h = parsed.headers || [];
    if (h.includes('Infringing URL')) return 'Linking file';
    if (h.includes('Source URL')) return 'Host file';
    return 'koi pehchaani hui platform file nahi';
  }
  function validateNewData(parsed, plat) {
    const RC = root.ReportCore; const h = parsed.headers || []; const has = (c) => h.includes(c);
    const wrong = (exp) => new Error('"' + plat + '" ka format nahi hai' + (exp ? ' (chahiye: ' + exp + ')' : '') + '. Ye ' + describe(parsed) + ' lag rahi hai' + (parsed.spec || parsed.role === 'social' || has('Infringing URL') || has('Source URL') ? ' - uske sahi button me daalo.' : '.'));
    const spec = RC && RC.PLATFORM_SPECS ? Object.values(RC.PLATFORM_SPECS).find((s) => s.label === plat) : null;
    if (!parsed.rows || !parsed.rows.length) throw new Error('File me data rows nahi hain.');
    if (spec) {
      if (parsed.spec !== plat) throw wrong('columns: ' + spec.sig.join(', '));
      if (!has(spec.url)) throw new Error('"' + plat + '" file me URL column "' + spec.url + '" nahi mila.');
    } else if (plat === 'Linking') {
      if (parsed.spec || parsed.role !== 'infringing' || !has('Infringing URL')) throw wrong(EXPECT.Linking);
    } else if (plat === 'Host') {
      if (parsed.spec || parsed.role !== 'source' || !has('Source URL')) throw wrong(EXPECT.Host);
    } else if (plat === 'Social Media & UGC') {
      if (parsed.spec || parsed.role !== 'social' || !has('Source URL')) throw wrong(EXPECT['Social Media & UGC']);
    } else {                                                           /* "+ Other..." platform: any one of the known layouts */
      const known = parsed.spec || parsed.role === 'social' ? true : (has('Infringing URL') || has('Source URL'));
      if (!known) throw new Error('"' + plat + '" file ka format pehchana nahi gaya - kisi bhi supported format me "Infringing URL" / "Source URL" jaisa URL column chahiye.');
    }
  }

  root.FileCheck = { magic, checkName, checkBytes, validateLive, validateComet, validateRemoval, validateNewData };
})(typeof self !== 'undefined' ? self : globalThis);
