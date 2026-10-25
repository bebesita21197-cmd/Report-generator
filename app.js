(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const files = []; // File objects
  let prevFile = null;
  let cometFile = null;
  let remFile = null;
  const fmtMB = (n) => (n / 1048576).toFixed(1) + ' MB';
  const tick = () => new Promise((r) => setTimeout(r, 30));
  const FC = FileCheck;

  /* wrong-format popup: one dialog per Browse/drop, lists every rejected file with the reason */
  function showErr(title, items) {
    const ov = document.createElement('div'); ov.className = 'modal-ov';
    const box = document.createElement('div'); box.className = 'modal';
    const h = document.createElement('h3'); h.textContent = title; h.style.color = '#b3261e';
    const list = document.createElement('div'); list.className = 'modal-list';
    items.forEach((it) => { const b = document.createElement('b'); b.textContent = it.name; const d = document.createElement('div'); d.textContent = it.msg; d.title = it.msg; d.style.whiteSpace = 'normal'; list.append(b, d); });
    const row = document.createElement('div'); row.className = 'modal-btns';
    const ok = document.createElement('button'); ok.textContent = 'OK'; ok.className = 'primary sm'; ok.onclick = () => ov.remove(); row.appendChild(ok);
    box.append(h, list, row); ov.appendChild(box); document.body.appendChild(ov); ok.focus();
  }
  /* check one picked file (name -> bytes -> structure); returns {buf} or {err} */
  async function vet(f, allowXls, validate) {
    try { FC.checkName(f.name, allowXls); const buf = await f.arrayBuffer(); FC.checkBytes(buf, f.name, allowXls); if (validate) validate(buf); return { buf }; }
    catch (e) { return { err: e.message }; }
  }
  /* single-file pickers (Live / Comet / Removal): first file is checked, on failure the old selection stays */
  async function pickOne(list, allowXls, validate, apply, what) {
    const f = Array.from(list)[0]; if (!f) return;
    const r = await vet(f, allowXls, validate);
    if (r.err) { showErr(what + ' - galat file', [{ name: f.name, msg: r.err }]); return; }
    apply(f);
  }

  /* report time = always 'now' (IST); default since = 12:00 AM of the previous day */
  const istNow = new Date(Date.now() + 5.5 * 3600e3);
  const prev = new Date(Date.UTC(istNow.getUTCFullYear(), istNow.getUTCMonth(), istNow.getUTCDate() - 1));
  $('since').value = prev.toISOString().slice(0, 16);

  /* ExcelJS drops the Notes-tab text box - copy it back from the previous file into the output */
  async function keepNotes(out, origBuf) {
    const r = await NotesCore.restoreNotes(JSZip, origBuf, out);
    if (r.restored) { note('Notes tab text box previous report se as-is copy hua'); return r.buf; }
    if (r.why && !/^no (Notes sheet|drawing)/.test(r.why)) note('Notes text box copy nahi ho paya: ' + r.why, 'warn');
    return out;
  }

  const hasName = () => $('reportName').value.replace(/\.xlsx$/i, '').replace(/[\\\/:*?"<>|\u0000-\u001f]/g, '').trim().length > 0;
  const updateGo = () => {
    const ok = hasName() && !(files.length === 0 && !remFile) && !files.some((f) => f._state === 'loading') && !(files.length > 0 && sel.size === 0);
    $('go').disabled = !prevFile || !ok; $('goComet').disabled = !cometFile || !ok; $('goFind').disabled = !files.length || !ok;
  };

  /* ---- title selection (all titles of the uploaded files, multi-select with search) ---- */
  const titles = new Map(), sel = new Set(), known = new Set();
  const titleOf = (r) => String(r['Asset Name'] == null ? '' : r['Asset Name']).trim() || '(No title)';
  function rebuildTitles() {
    titles.clear();
    files.forEach((f) => { if (f._parsed) f._parsed.rows.forEach((r) => { const t = titleOf(r); titles.set(t, (titles.get(t) || 0) + 1); }); });
    for (const t of titles.keys()) if (!known.has(t)) { known.add(t); }          /* new titles start UNselected - user picks them */
    Array.from(known).forEach((t) => { if (!titles.has(t)) { known.delete(t); sel.delete(t); } });
    renderTitles();
  }
  function renderTitles() {
    const shown = shownTitles().sort((a, b) => a.localeCompare(b));
    const box = $('titleList'); box.innerHTML = '';
    ['titleSearch', 'selAll', 'selNone'].forEach((id) => { $(id).disabled = titles.size === 0; });
    if (!shown.length) { const d = document.createElement('div'); d.className = 'tempty'; d.textContent = titles.size ? 'No title matches your search.' : 'Files load hone ke baad saare titles yaha dikhenge.'; box.appendChild(d); }
    const frag = document.createDocumentFragment();
    shown.forEach((t) => {
      const l = document.createElement('label'); l.className = 'titem' + (sel.has(t) ? ' on' : '');
      const cb = document.createElement('input'); cb.type = 'checkbox'; cb.checked = sel.has(t);
      cb.onchange = () => { if (cb.checked) sel.add(t); else sel.delete(t); l.classList.toggle('on', cb.checked); updateTitleCount(); updateGo(); };
      const n = document.createElement('span'); n.className = 'tn'; n.textContent = t; n.title = t;
      const c = document.createElement('span'); c.className = 'tc'; c.textContent = titles.get(t) + ' rows';
      l.append(cb, n, c); frag.appendChild(l);
    });
    box.appendChild(frag); updateTitleCount(); updateGo();
  }
  function updateTitleCount() {
    let rows = 0; sel.forEach((t) => { rows += titles.get(t) || 0; });
    $('titleCount').textContent = titles.size ? sel.size + ' of ' + titles.size + ' titles selected \u00B7 ' + rows + ' rows' : '';
    const sm = $('selSummary'); $('openTitles').disabled = titles.size === 0;
    sm.className = 'selsum' + (titles.size ? ' ok' : '');
    sm.textContent = titles.size ? sel.size + ' of ' + titles.size + ' titles selected' : (files.some((f) => f._state === 'loading') ? 'Loading titles...' : 'Upload files to load titles');
  }
  /* every word typed in the search box must be present in the title (any order, any position) */
  const shownTitles = () => { const w = $('titleSearch').value.trim().toLowerCase().split(/\s+/).filter(Boolean); return Array.from(titles.keys()).filter((t) => { const l = t.toLowerCase(); return w.every((x) => l.includes(x)); }); };
  const drawer = $('drawer'), backdrop = $('backdrop');
  const openDrawer = () => { drawer.classList.add('open'); drawer.setAttribute('aria-hidden', 'false'); backdrop.hidden = false; setTimeout(() => $('titleSearch').focus(), 120); };
  const closeDrawer = () => { drawer.classList.remove('open'); drawer.setAttribute('aria-hidden', 'true'); backdrop.hidden = true; };
  /* resizable drawer: drag the left edge (double-click resets to 1/4 screen); width is remembered */
  const dragger = $('dragger'); const clampW = (w) => Math.max(340, Math.min(w, window.innerWidth - 40));
  try { const w = parseInt(localStorage.getItem('rg_drawer_w'), 10); if (w) drawer.style.width = clampW(w) + 'px'; } catch (e) { /* storage unavailable */ }
  dragger.addEventListener('pointerdown', (e) => { e.preventDefault(); dragger.setPointerCapture(e.pointerId); drawer.classList.add('dragging'); });
  dragger.addEventListener('pointermove', (e) => { if (drawer.classList.contains('dragging')) drawer.style.width = clampW(window.innerWidth - e.clientX) + 'px'; });
  const endDrag = () => { if (!drawer.classList.contains('dragging')) return; drawer.classList.remove('dragging'); try { localStorage.setItem('rg_drawer_w', parseInt(drawer.style.width, 10)); } catch (e) { /* ignore */ } };
  dragger.addEventListener('pointerup', endDrag); dragger.addEventListener('pointercancel', endDrag);
  dragger.addEventListener('dblclick', () => { drawer.style.width = ''; try { localStorage.removeItem('rg_drawer_w'); } catch (e) { /* ignore */ } });
  window.addEventListener('resize', () => { if (drawer.style.width) drawer.style.width = clampW(parseInt(drawer.style.width, 10)) + 'px'; });
  $('reportName').addEventListener('input', () => updateGo());
  $('openTitles').addEventListener('click', openDrawer);
  ['closeTitles', 'doneTitles'].forEach((id) => $(id).addEventListener('click', closeDrawer));
  backdrop.addEventListener('click', closeDrawer);
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeDrawer(); });
  $('titleSearch').addEventListener('input', renderTitles);
  $('selAll').addEventListener('click', () => { shownTitles().forEach((t) => sel.add(t)); renderTitles(); });
  $('selNone').addEventListener('click', () => { shownTitles().forEach((t) => sel.delete(t)); renderTitles(); });
  /* platform buttons: every button is also the Browse button for that platform (file name is never used) */
  const PLATFORMS = ['Facebook', 'Host', 'Instagram', 'Linking', 'Market Place', 'Social Media & UGC', 'Telegram', 'X', 'YouTube'];
  const extraPlats = [];
  const pbox = $('platBtns');
  function renderList() {
    pbox.innerHTML = '';
    PLATFORMS.concat(extraPlats).forEach((name) => {
      const list = files.filter((f) => f._platform === name);
      const loading = list.some((f) => f._state === 'loading'); const err = list.find((f) => f._state === 'error');
      const b = document.createElement('div'); b.className = 'pbtn' + (list.length ? (err ? ' err' : loading ? ' load' : ' ok') : '');
      const lab = document.createElement('label'); lab.className = 'pb-main'; lab.dataset.plat = name;
      const inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.xlsx,.xls'; inp.multiple = true; inp.hidden = true;
      inp.addEventListener('change', (e) => { addFiles(e.target.files, name); e.target.value = ''; });
      const t = document.createElement('b'); t.textContent = name;
      const st = document.createElement('span');
      const rows = list.reduce((n, f) => n + (f._parsed ? f._parsed.rows.length : 0), 0);
      st.textContent = list.length ? (err ? '\u2717 Failed' : loading ? '\u23F3 Loading' : '\u2713 ' + rows + ' rows') : '\u2B06 Browse';
      lab.title = list.length ? list.map((f) => f.name + (f._parsed ? ' (' + f._parsed.rows.length + ' rows)' : '')).join('\n') + (err ? '\n' + (err._err || 'Failed') : '') + '\n\nClick = add more files' : 'Browse ' + name + ' file(s)';
      lab.append(inp, t, st); b.appendChild(lab);
      if (list.length) { const x = document.createElement('button'); x.type = 'button'; x.textContent = '\u2715'; x.title = 'Remove ' + name; x.onclick = () => { for (let i = files.length - 1; i >= 0; i--) if (files[i]._platform === name) files.splice(i, 1); renderList(); rebuildTitles(); }; b.appendChild(x); }
      ['dragenter', 'dragover'].forEach((ev) => b.addEventListener(ev, (e) => { e.preventDefault(); b.classList.add('over'); }));
      ['dragleave', 'drop'].forEach((ev) => b.addEventListener(ev, (e) => { e.preventDefault(); b.classList.remove('over'); }));
      b.addEventListener('drop', (e) => addFiles(e.dataTransfer.files, name));
      pbox.appendChild(b);
    });
    const other = document.createElement('button'); other.type = 'button'; other.className = 'pbtn add'; other.textContent = '+ Other...';
    other.onclick = () => {
      const v = (window.prompt('Platform name:') || '').trim().replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase()); if (!v) return;
      if (!PLATFORMS.concat(extraPlats).includes(v)) extraPlats.push(v);
      renderList(); const lb = Array.from(pbox.querySelectorAll('label.pb-main')).find((l) => l.dataset.plat === v); if (lb) lb.click();
    };
    pbox.appendChild(other);
    updateTitleCount(); updateGo();
  }
  function renderPrev() {
    const ul = $('prevList'); ul.innerHTML = ''; $('dropPrev').classList.toggle('loaded', !!prevFile);
    if (prevFile) {
      const li = document.createElement('li');
      const left = document.createElement('span'); left.className = 'name'; left.textContent = prevFile.name;
      const st = document.createElement('span'); st.className = 'st ok'; st.textContent = '\u2713 File loaded'; left.appendChild(st);
      const right = document.createElement('span'); right.className = 'meta'; right.textContent = fmtMB(prevFile.size) + ' ';
      const x = document.createElement('button'); x.textContent = '\u2715'; x.title = 'Remove'; x.onclick = () => { prevFile = null; renderPrev(); };
      right.appendChild(x); li.append(left, right); ul.appendChild(li);
    }
    updateGo();
  }
  function setPrev(list) { return pickOne(list, false, (b) => FC.validateLive(XLSX, b), (f) => { prevFile = f; renderPrev(); }, 'Previous Live Report'); }
  $('prevFile').addEventListener('change', (e) => { setPrev(e.target.files); e.target.value = ''; });
  const dropPrev = $('dropPrev');
  ['dragenter', 'dragover'].forEach((ev) => dropPrev.addEventListener(ev, (e) => { e.preventDefault(); dropPrev.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => dropPrev.addEventListener(ev, (e) => { e.preventDefault(); dropPrev.classList.remove('over'); }));
  dropPrev.addEventListener('drop', (e) => setPrev(e.dataTransfer.files));
  async function addFiles(list, plat) {
    const fresh = [], errs = [];
    for (const f of list) {
      try { FC.checkName(f.name, true); } catch (e) { errs.push({ name: f.name, msg: e.message }); continue; }
      if (!files.some((g) => g.name === f.name && g.size === f.size && g._platform === plat)) { f._state = 'loading'; f._platform = plat; files.push(f); fresh.push(f); }
    }
    renderList();
    for (const f of fresh) {
      await tick();
      try {
        const buf = await f.arrayBuffer(); FC.checkBytes(buf, f.name, true);
        let parsed; try { parsed = ReportCore.parseWorkbook(XLSX, buf, ''); } catch (e) { throw new Error('Excel file khul nahi paayi / format galat: ' + e.message); }
        FC.validateNewData(parsed, plat);
        f._parsed = parsed; f._role = parsed.role; f._state = 'ready';
      } catch (e) { console.error(e); errs.push({ name: f.name, msg: e.message }); const i = files.indexOf(f); if (i >= 0) files.splice(i, 1); }
      renderList(); rebuildTitles();
    }
    if (errs.length) showErr(plat + ' - ' + errs.length + ' file ka format sahi nahi (add nahi hui)', errs);
  }
  renderList();

  /* ---- pending removal update: file picker + template ---- */
  function renderComet() {
    const ul = $('cometList'); ul.innerHTML = ''; $('dropComet').classList.toggle('loaded', !!cometFile);
    if (cometFile) {
      const li = document.createElement('li');
      const left = document.createElement('span'); left.className = 'name'; left.textContent = cometFile.name;
      const st = document.createElement('span'); st.className = 'st ok'; st.textContent = '\u2713 File loaded'; left.appendChild(st);
      const right = document.createElement('span'); right.className = 'meta'; right.textContent = fmtMB(cometFile.size) + ' ';
      const x = document.createElement('button'); x.textContent = '\u2715'; x.title = 'Remove'; x.onclick = () => { cometFile = null; renderComet(); };
      right.appendChild(x); li.append(left, right); ul.appendChild(li);
    }
    updateGo();
  }
  function setComet(list) { return pickOne(list, false, (b) => FC.validateComet(XLSX, b), (f) => { cometFile = f; renderComet(); }, 'Previous Comet Report'); }
  $('cometFile').addEventListener('change', (e) => { setComet(e.target.files); e.target.value = ''; });
  const dropComet = $('dropComet');
  ['dragenter', 'dragover'].forEach((ev) => dropComet.addEventListener(ev, (e) => { e.preventDefault(); dropComet.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => dropComet.addEventListener(ev, (e) => { e.preventDefault(); dropComet.classList.remove('over'); }));
  dropComet.addEventListener('drop', (e) => setComet(e.dataTransfer.files));
  function renderRem() {
    const ul = $('remList'); ul.innerHTML = ''; $('dropRem').classList.toggle('loaded', !!remFile);
    if (remFile) {
      const li = document.createElement('li');
      const left = document.createElement('span'); left.className = 'name'; left.textContent = remFile.name;
      const t = document.createElement('span'); t.className = 'tag removal'; t.textContent = 'Removal update'; left.appendChild(t);
      const st = document.createElement('span'); st.className = 'st ok'; st.textContent = '\u2713 File loaded'; left.appendChild(st);
      const right = document.createElement('span'); right.className = 'meta'; right.textContent = fmtMB(remFile.size) + ' ';
      const x = document.createElement('button'); x.textContent = '\u2715'; x.title = 'Remove'; x.onclick = () => { remFile = null; renderRem(); };
      right.appendChild(x); li.append(left, right); ul.appendChild(li);
    }
    updateGo();
  }
  function setRem(list) { return pickOne(list, true, (b) => FC.validateRemoval(XLSX, b), (f) => { remFile = f; renderRem(); }, 'Pending removal update'); }
  $('remFile').addEventListener('change', (e) => { setRem(e.target.files); e.target.value = ''; });
  const dropRem = $('dropRem');
  ['dragenter', 'dragover'].forEach((ev) => dropRem.addEventListener(ev, (e) => { e.preventDefault(); dropRem.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => dropRem.addEventListener(ev, (e) => { e.preventDefault(); dropRem.classList.remove('over'); }));
  dropRem.addEventListener('drop', (e) => setRem(e.dataTransfer.files));
  $('tplBtn').addEventListener('click', async () => {
    const tpl = ReportCore.buildRemovalTemplate(ExcelJS); ReportCore.sanitizeWorkbook(tpl);
    const buf = await tpl.xlsx.writeBuffer();
    const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = document.createElement('a'); a.href = url; a.download = 'Pending_Removal_Update_Template.xlsx'; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  });


  /* popup: Facebook / Instagram / YouTube / X rows that are still Active (Removal Status not Dead) -> skip them or add anyway */
  const ACTIVE_CHECK = ['facebook', 'instagram', 'youtube', 'twitter'];
  function askActive(list) {
    return new Promise((resolve) => {
      const ov = document.createElement('div'); ov.className = 'modal-ov';
      const box = document.createElement('div'); box.className = 'modal';
      const by = {}; list.forEach((x) => { (by[x.label] = by[x.label] || []).push(x.url); });
      const h = document.createElement('h3'); h.textContent = 'Active URLs mile (' + list.length + ')';
      const p = document.createElement('p'); p.textContent = 'In URLs ka Removal Status Dead nahi hai ya Removal Time blank hai. Skip karoge to ye report me add nahi honge; warna Findings + Takedown Pending me jayenge.';
      const ul = document.createElement('div'); ul.className = 'modal-list';
      Object.keys(by).forEach((k) => { const t = document.createElement('b'); t.textContent = k + ' - ' + by[k].length; ul.appendChild(t); by[k].forEach((u) => { const d = document.createElement('div'); d.textContent = u; d.title = u; ul.appendChild(d); }); });
      const row = document.createElement('div'); row.className = 'modal-btns';
      const mk = (txt, cls, val) => { const b = document.createElement('button'); b.textContent = txt; b.className = cls; b.onclick = () => { ov.remove(); resolve(val); }; return b; };
      row.append(mk('Add all (keep active)', 'secondary', 'keep'), mk('Skip active URLs', 'primary sm', 'skip'));
      box.append(h, p, ul, row); ov.appendChild(box); document.body.appendChild(ov);
    });
  }

  /* popup: rows with missing Enforcement time / blank Dolby status -> report is generated ONLY if the user clicks "Still want to generate" */
  function askMissing(list) {
    return new Promise((resolve) => {
      const ov = document.createElement('div'); ov.className = 'modal-ov';
      const box = document.createElement('div'); box.className = 'modal';
      const h = document.createElement('h3'); h.textContent = 'Missing data mila (' + list.length + ' rows)';
      const p = document.createElement('p'); p.textContent = 'In rows me Enforcement time missing hai ya Dolby status blank hai. Missing Enforcement time par Identification time use hota hai. Report tabhi banegi jab "Still want to generate" dabaoge.';
      const ul = document.createElement('div'); ul.className = 'modal-list';
      list.forEach((x) => { const d = document.createElement('div'); d.textContent = '[' + x.label + '] ' + x.miss.join(' + ') + ' - ' + x.url; d.title = d.textContent; ul.appendChild(d); });
      const row = document.createElement('div'); row.className = 'modal-btns';
      const mk = (txt, cls, val) => { const b = document.createElement('button'); b.textContent = txt; b.className = cls; b.onclick = () => { ov.remove(); resolve(val); }; return b; };
      row.append(mk('Cancel', 'secondary', false), mk('Still want to generate', 'primary sm', true));
      box.append(h, p, ul, row); ov.appendChild(box); document.body.appendChild(ov);
    });
  }

  /* ---- bottom process line ---- */
  const stepEl = (id) => document.querySelector('#stepList [data-step="' + id + '"]');
  const setStep = (id, state, detail) => { const li = stepEl(id); li.className = 's-' + state; if (detail != null) li.querySelector('small').textContent = detail; };
  const resetSteps = () => { document.querySelectorAll('#stepList li').forEach((li) => { li.className = ''; li.querySelector('small').textContent = 'Waiting'; }); const n = $('notes'); n.innerHTML = ''; n.hidden = true; };
  const note = (m, kind) => { const n = $('notes'); n.hidden = false; const c = document.createElement('span'); c.className = 'note' + (kind ? ' ' + kind : ''); c.textContent = m; c.title = m; n.appendChild(c); };
  const num = (v) => Number(v || 0);
  /* manual report name: used as the download name when filled (invalid file-name characters removed, .xlsx added), else the automatic name */
  const outName = (auto) => { const v = $('reportName').value.replace(/\.xlsx$/i, '').replace(/[\\\/:*?"<>|\u0000-\u001f]/g, '').replace(/\s+/g, ' ').trim(); return v ? v + '.xlsx' : auto; };

  /* shared by Live + Comet: selected titles -> optional active popup -> build records (every new URL is Pending) */
  async function prepareNew(begin, kw, refTs, sinceTs, chosen, extra) {
    await begin('read', 'Filtering...');
    const data = { infringing: [], source: [], social: [], platform: [] }; let nrows = 0;
    for (const f of files) {
      const p = f._parsed; let rows = p.rows.filter((r) => sel.has(titleOf(r)));
      if (f._platform) rows = rows.map((r) => Object.assign({}, r, { __platform: f._platform }));
      data[p.role] = data[p.role].concat(rows); nrows += rows.length;
      if (p.spec && f._platform && p.spec !== f._platform) note('CHECK: ' + f.name + ' ek ' + p.spec + ' file lag rahi hai, lekin platform "' + f._platform + '" select hai', 'warn');
      note(f.name + ' → ' + p.role + (f._platform ? ' [Platform: ' + f._platform + ']' : '') + ': ' + rows.length + ' of ' + p.total + ' rows');
    }
    const blankV = (v) => v == null || (typeof v === 'string' && (v.trim() === '' || v.trim() === '-'));
    const active = data.platform.filter((r) => {
      if (!ACTIVE_CHECK.includes(r.__spec)) return false;
      const sp = ReportCore.PLATFORM_SPECS[r.__spec];
      const notDead = !/^dead$/i.test(String(r[sp.status] == null ? '' : r[sp.status]).trim());
      const noRemTime = !(sp.removal || []).some((k) => !blankV(r[k]));         /* Removal Time blank -> also treated as not removed yet */
      return notDead || noRemTime;
    });
    if (active.length) {
      const choice = await askActive(active.map((r) => ({ label: ReportCore.PLATFORM_SPECS[r.__spec].label, url: String(r[ReportCore.PLATFORM_SPECS[r.__spec].url]) })));
      if (choice === 'skip') { const bad = new Set(active); data.platform = data.platform.filter((r) => !bad.has(r)); nrows -= active.length; note('Active URLs skip kiye: ' + active.length, 'warn'); }
    }
    if (!nrows) throw new Error('Selected titles me koi row nahi mili (ya sab active URLs skip ho gaye).');
    const missing = ReportCore.checkMissing(data);
    if (missing.length) {
      if (!(await askMissing(missing))) throw new Error('Cancelled - missing Enforcement time / Dolby status wale rows ki wajah se report nahi bani.');
      note('Still generated with missing Enforcement time / Dolby status rows: ' + missing.length, 'warn');
    }
    setStep('read', 'done', nrows + ' rows · ' + chosen.length + ' titles');
    await begin('build', 'Building...');
    const res = ReportCore.build(data, Object.assign({ keyword: kw, refTs, sinceTs }, extra || {}));
    setStep('build', 'done', res.findings.length + ' records');
    return res;
  }

  $('go').addEventListener('click', async () => {
    const btn = $('go'); btn.disabled = true; $('goComet').disabled = true; $('goFind').disabled = true; resetSteps(); $('stats').hidden = true; $('result').hidden = true; $('status').className = 'status'; $('status').textContent = '';
    let cur = 'prev';
    const begin = async (id, d) => { cur = id; setStep(id, 'run', d || 'Working...'); await tick(); };
    try {
      if (!hasName()) throw new Error('Pehle Report name (file name) daalo.');
      if (!prevFile) throw new Error('Pehle previous report select karo.');
      if (!files.length && !remFile) throw new Error('New data files ya removal update file me se kuch to upload karo.');
      const chosen = Array.from(sel);
      if (files.length && !chosen.length) throw new Error('Kam se kam ek title select karo.');
      if (files.some((f) => f._state !== 'ready')) throw new Error('Kuch files load nahi hui hain - unhe remove karke dobara add karo.');
      const kw = chosen.length === 1 ? chosen[0].slice(0, 30) : '';
      const refTs = Date.now() + 5.5 * 3600e3;
      const sinceTs = $('since').value ? ReportCore.toTs($('since').value) : null;

      await begin('prev', 'Loading...');
      const prevBuf = await prevFile.arrayBuffer();
      const info = await ReportCore.loadPrevious(ExcelJS, prevBuf);
      const pc = info.prevCounts;
      setStep('prev', 'done', 'F ' + num(pc.findings) + ' · C ' + num(pc.completed) + ' · P ' + num(pc.pending) + ' rows');
      if (info.missing.length) note('WARNING: previous report me nahi mila: ' + info.missing.join(', '), 'warn');

      /* A) new data files FIRST (optional when a removal update is given) */
      let A = { findings: 0, completed: 0, pending: 0, delist: 0, social: 0 }, T, ap = null, res = null;
      if (files.length) {
        res = await prepareNew(begin, kw, refTs, sinceTs, chosen);
        if (!files.some((f) => f._role === 'infringing')) note('WARNING: no Infringing (Linking) file - Linking & Delisting sheets will be empty.', 'warn');
        await begin('append', 'Appending...');
        ap = ReportCore.appendToPrevious(info, res, { refTs, remMoved: null });
        A = ap.added; T = ap.totals;
        setStep('append', 'done', '+' + A.findings + ' F · +' + A.completed + ' C · +' + A.pending + ' P');
        note('HSP/IP/Reason/Server from previous report - exact domain ' + ap.dom.exact + ', base domain ' + ap.dom.base + ', current data ' + ap.dom.fallback);
        if (res.stats.randomDelist) note('Delisting: ' + res.stats.randomDelist + ' rows had no "Delisting Done At" -> random 10-20 min used');
        if (res.stats.dedupDropped) note('Duplicate URLs removed: ' + res.stats.dedupDropped);
        res.warn.slice(0, 20).forEach((w) => note('WARN: ' + w, 'warn'));
      } else {
        ['read', 'build', 'append'].forEach((id) => setStep(id, 'skip', 'No data files'));
      }
      /* B) pending removal update LAST (moves Pending -> Completed, Dolby travels along; also completes URLs appended just now) */
      let rem = null;
      if (remFile) {
        await begin('rem', 'Applying...');
        const entries = ReportCore.parseRemovalFile(XLSX, await remFile.arrayBuffer());
        rem = ReportCore.applyRemovalUpdate(info, entries);
        setStep('rem', 'done', rem.moved + ' of ' + rem.total + ' moved');
        if (rem.notFound.length) note(rem.notFound.length + ' URL Takedown Pending me nahi mile (skip): ' + rem.notFound.slice(0, 5).join(' | ') + (rem.notFound.length > 5 ? ' ...' : ''), 'warn');
        if (rem.alreadyDone.length) note('Already in Takedown Completed (skip): ' + rem.alreadyDone.length);
        if (rem.badTime.length) note('Removal Time samajh nahi aaya (skip) - Excel line ' + rem.badTime.slice(0, 10).map((b) => b.line).join(', '), 'warn');
        if (rem.dupes) note('Duplicate URL rows in removal file ignored: ' + rem.dupes);
        if (rem.clamped) note('Removal time < initiated time tha -> initiated time use hua: ' + rem.clamped);
      } else setStep('rem', 'skip', 'Not used');

      if (ap) { if (rem) T = ReportCore.finishSummary(info, ap, rem.movedCompleted); else T = ap.totals; }
      else T = ReportCore.refreshTotalsOnly(info, { since: sinceTs, remMoved: rem ? rem.movedCompleted : 0 });
      note('Open issues remaining: ' + T.open + (T.stalePending ? ' (' + T.stalePending + ' purane Pending rows ab Completed hain, minus kiye)' : ''));
      const stats = [['New rows - Findings', A.findings], ['New rows - Completed', A.completed + (rem ? rem.movedCompleted : 0)], ['New rows - Pending', A.pending], ['New rows - Delisting', A.delist], ['New rows - Social', A.social + (rem ? rem.movedSocial : 0)]];
      if (rem) stats.push(['Moved Pending → Completed', rem.moved], ['Removal URLs not found', rem.notFound.length]);
      stats.push(['Total infringements', T.total], ['Takedown completed', T.comp], ['Takedown pending', T.open]);
      $('stats').innerHTML = stats.map(([l, v]) => '<div class="stat"><b>' + v + '</b><span>' + l + '</span></div>').join(''); $('stats').hidden = false;

      await begin('final', 'Updating...');
      const fin = ReportCore.finalizeReport(info, refTs);
      setStep('final', 'done', fin.durations + ' durations · ' + fin.delistFmt + ' delist fmt');
      await begin('write', 'Writing...');
      ReportCore.sanitizeWorkbook(info.wb);
      const out = await keepNotes(await info.wb.xlsx.writeBuffer(), prevBuf);
      const name = outName(ReportCore.fileName(refTs, kw || 'Report'));
      const url = URL.createObjectURL(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a'); a.href = url; a.download = name; a.textContent = name;
      $('result').innerHTML = 'Report ready: '; $('result').appendChild(a); $('result').hidden = false; a.click();
      setStep('write', 'done', 'Downloaded');
      $('status').textContent = 'Done.';
    } catch (err) {
      console.error(err); setStep(cur, 'err', err.message.slice(0, 60)); $('status').className = 'status err'; $('status').textContent = 'Error: ' + err.message; note('ERROR: ' + (err.stack || err.message).split('\n')[0], 'err');
    } finally { updateGo(); }
  });

  /* ================= Generate Comet Report ================= */
  $('goComet').addEventListener('click', async () => {
    $('go').disabled = true; $('goComet').disabled = true; $('goFind').disabled = true; resetSteps(); $('stats').hidden = true; $('result').hidden = true; $('status').className = 'status'; $('status').textContent = '';
    let cur = 'prev';
    const begin = async (id, d) => { cur = id; setStep(id, 'run', d || 'Working...'); await tick(); };
    try {
      if (!hasName()) throw new Error('Pehle Report name (file name) daalo.');
      if (!cometFile) throw new Error('Pehle previous Comet report select karo.');
      if (!files.length && !remFile) throw new Error('New data files ya removal update file me se kuch to upload karo.');
      const chosen = Array.from(sel);
      if (files.length && !chosen.length) throw new Error('Kam se kam ek title select karo.');
      if (files.some((f) => f._state !== 'ready')) throw new Error('Kuch files load nahi hui hain - unhe remove karke dobara add karo.');
      const kw = chosen.length === 1 ? chosen[0].slice(0, 30) : '';
      const refTs = Date.now() + 5.5 * 3600e3;
      const sinceTs = $('since').value ? ReportCore.toTs($('since').value) : null;

      await begin('prev', 'Loading Comet...');
      const cometBuf = await cometFile.arrayBuffer();
      const cinfo = await ReportCore.loadComet(ExcelJS, cometBuf);
      setStep('prev', 'done', cinfo.counts.total + ' rows · ' + cinfo.counts.pending + ' pending');
      note('Comet: IST \u2192 EST = minus 10h 30m');

      let ap = { added: 0, skipped: 0 };
      if (files.length) {
        const res = await prepareNew(begin, kw, refTs, sinceTs, chosen);
        await begin('append', 'Appending...');
        ap = ReportCore.appendComet(cinfo, res);
        setStep('append', 'done', '+' + ap.added + ' rows');
        note('HSP/Reason from existing Comet pending rows - exact domain ' + ap.dom.exact + ', base domain ' + ap.dom.base + ', current data ' + ap.dom.fallback);
        if (ap.skipped) note('URLs already in Comet report / duplicates (skip): ' + ap.skipped);
        res.warn.slice(0, 20).forEach((w) => note('WARN: ' + w, 'warn'));
      } else ['read', 'build', 'append'].forEach((id) => setStep(id, 'skip', 'No data files'));
      let rem = null;
      if (remFile) {
        await begin('rem', 'Applying...');
        rem = ReportCore.applyCometRemoval(cinfo, ReportCore.parseRemovalFile(XLSX, await remFile.arrayBuffer()));
        setStep('rem', 'done', rem.moved + ' of ' + rem.total + ' completed');
        if (rem.notFound.length) note(rem.notFound.length + ' URL Comet report me nahi mile (skip): ' + rem.notFound.slice(0, 5).join(' | ') + (rem.notFound.length > 5 ? ' ...' : ''), 'warn');
        if (rem.alreadyDone.length) note('Already completed in Comet (skip): ' + rem.alreadyDone.length);
        if (rem.badTime.length) note('Removal Time samajh nahi aaya (skip) - Excel line ' + rem.badTime.slice(0, 10).map((b) => b.line).join(', '), 'warn');
        if (rem.dupes) note('Duplicate URL rows in removal file ignored: ' + rem.dupes);
        if (rem.clamped) note('Removal time < initiated time tha -> initiated time use hua: ' + rem.clamped);
      } else setStep('rem', 'skip', 'Not used');

      const dashed = ReportCore.dashCompletedReasons(cinfo);
      setStep('final', 'done', 'Pending reason "-": ' + dashed);
      note('Takedown Completed rows ka Reason Pending "-" kiya: ' + dashed);

      const stats = [['New rows added', ap.added], ['Completed (removal)', rem ? rem.moved : 0]];
      if (rem) stats.push(['Removal URLs not found', rem.notFound.length]);
      stats.push(['Total rows', cinfo.counts.total], ['Pending (Y)', cinfo.counts.pending], ['Completed (Y)', cinfo.counts.total - cinfo.counts.pending]);
      $('stats').innerHTML = stats.map(([l, v]) => '<div class="stat"><b>' + v + '</b><span>' + l + '</span></div>').join(''); $('stats').hidden = false;

      await begin('write', 'Writing...');
      ReportCore.sanitizeWorkbook(cinfo.wb);
      const out = await keepNotes(await cinfo.wb.xlsx.writeBuffer(), cometBuf);
      const name = outName(ReportCore.cometFileName(refTs));
      const url = URL.createObjectURL(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a'); a.href = url; a.download = name; a.textContent = name;
      $('result').innerHTML = 'Comet report ready: '; $('result').appendChild(a); $('result').hidden = false; a.click();
      setStep('write', 'done', 'Downloaded');
      $('status').textContent = 'Done.';
    } catch (err) {
      console.error(err); setStep(cur, 'err', err.message.slice(0, 60)); $('status').className = 'status err'; $('status').textContent = 'Error: ' + err.message; note('ERROR: ' + (err.stack || err.message).split('\n')[0], 'err');
    } finally { updateGo(); }
  });

  /* ================= Generate Findings Report (no previous file: built only from the uploaded data) ================= */
  $('goFind').addEventListener('click', async () => {
    $('go').disabled = true; $('goComet').disabled = true; $('goFind').disabled = true; resetSteps(); $('stats').hidden = true; $('result').hidden = true; $('status').className = 'status'; $('status').textContent = '';
    let cur = 'prev';
    const begin = async (id, d) => { cur = id; setStep(id, 'run', d || 'Working...'); await tick(); };
    try {
      if (!hasName()) throw new Error('Pehle Report name (file name) daalo.');
      if (!files.length) throw new Error('Findings report ke liye New data files upload karo.');
      const chosen = Array.from(sel);
      if (!chosen.length) throw new Error('Kam se kam ek title select karo.');
      if (files.some((f) => f._state !== 'ready')) throw new Error('Kuch files load nahi hui hain - unhe remove karke dobara add karo.');
      const kw = chosen.length === 1 ? chosen[0].slice(0, 30) : '';
      const refTs = Date.now() + 5.5 * 3600e3;
      const sinceTs = $('since').value ? ReportCore.toTs($('since').value) : null;
      setStep('prev', 'skip', 'Not needed');
      note('Findings: sirf upload kiye data se banega (previous report nahi chahiye). IST \u2192 EST = minus 10h 30m');

      const res = await prepareNew(begin, kw, refTs, sinceTs, chosen, { allPending: false });      /* completed / pending as given in the data files */
      let entries = null;
      if (remFile) { await begin('rem', 'Reading...'); entries = ReportCore.parseRemovalFile(XLSX, await remFile.arrayBuffer()); } else setStep('rem', 'skip', 'Not used');
      await begin('append', 'Building sheet...');
      const out = ReportCore.buildFindings(ExcelJS, res, { removal: entries });
      setStep('append', 'done', out.stats.total + ' rows');
      if (out.rem) {
        setStep('rem', 'done', out.rem.moved + ' of ' + out.rem.total + ' completed');
        if (out.rem.notFound.length) note(out.rem.notFound.length + ' removal URL data files me nahi mile (skip): ' + out.rem.notFound.slice(0, 5).join(' | ') + (out.rem.notFound.length > 5 ? ' ...' : ''), 'warn');
        if (out.rem.alreadyDone.length) note('Already completed in data (skip): ' + out.rem.alreadyDone.length);
        if (out.rem.badTime.length) note('Removal Time samajh nahi aaya (skip) - Excel line ' + out.rem.badTime.slice(0, 10).map((b) => b.line).join(', '), 'warn');
        if (out.rem.dupes) note('Duplicate URL rows in removal file ignored: ' + out.rem.dupes);
      }
      if (res.stats.dedupDropped) note('Duplicate URLs removed: ' + res.stats.dedupDropped);
      res.warn.slice(0, 20).forEach((w) => note('WARN: ' + w, 'warn'));
      setStep('final', 'skip', 'Not needed');
      const stats = [['Total rows', out.stats.total], ['Takedown Completed (Y)', out.stats.completed], ['Takedown Pending (N)', out.stats.pending]];
      Object.keys(out.stats.byPlat).forEach((p) => stats.push([p, out.stats.byPlat[p]]));
      $('stats').innerHTML = stats.map(([l, v]) => '<div class="stat"><b>' + v + '</b><span>' + l + '</span></div>').join(''); $('stats').hidden = false;

      await begin('write', 'Writing...');
      ReportCore.sanitizeWorkbook(out.wb);
      const buf = await out.wb.xlsx.writeBuffer();
      const name = outName(ReportCore.findingsFileName(refTs, sinceTs));
      const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a'); a.href = url; a.download = name; a.textContent = name;
      $('result').innerHTML = 'Findings report ready: '; $('result').appendChild(a); $('result').hidden = false; a.click();
      setStep('write', 'done', 'Downloaded');
      $('status').textContent = 'Done.';
    } catch (err) {
      console.error(err); setStep(cur, 'err', err.message.slice(0, 60)); $('status').className = 'status err'; $('status').textContent = 'Error: ' + err.message; note('ERROR: ' + (err.stack || err.message).split('\n')[0], 'err');
    } finally { updateGo(); }
  });
})();
