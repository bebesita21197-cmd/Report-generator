/* Keeps the Notes-tab text box (DrawingML shape) of the previous report.
   ExcelJS drops drawings/shapes when it loads + saves a workbook, so after writing the new file we
   copy the Notes sheet drawing (text box + its media) from the ORIGINAL file back into the OUTPUT zip. */
(function (root) {
  'use strict';
  const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const DRAW_TYPE = REL_NS + '/drawing';
  const DRAW_CT = 'application/vnd.openxmlformats-officedocument.drawing+xml';
  const nh = (s) => String(s == null ? '' : s).replace(/ /g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  const attr = (tag, name) => { const m = new RegExp('\\b' + name + '="([^"]*)"').exec(tag); return m ? m[1] : null; };
  const dirOf = (p) => p.slice(0, p.lastIndexOf('/') + 1);
  const relsPathOf = (p) => dirOf(p) + '_rels/' + p.slice(p.lastIndexOf('/') + 1) + '.rels';
  function resolve(base, target) {
    if (target[0] === '/') return target.slice(1);
    const parts = dirOf(base).split('/').filter(Boolean);
    target.split('/').forEach((s) => { if (s === '..') parts.pop(); else if (s && s !== '.') parts.push(s); });
    return parts.join('/');
  }
  async function readText(zip, p) { const f = zip.file(p); return f ? f.async('string') : null; }

  async function sheetPaths(zip) {
    const wbx = await readText(zip, 'xl/workbook.xml'); const rel = await readText(zip, 'xl/_rels/workbook.xml.rels');
    const out = {}; if (!wbx || !rel) return out;
    const rid = {}; (rel.match(/<Relationship\b[^>]*>/g) || []).forEach((t) => { rid[attr(t, 'Id')] = attr(t, 'Target'); });
    (wbx.match(/<sheet\b[^>]*>/g) || []).forEach((t) => {
      const name = attr(t, 'name'); const id = attr(t, 'r:id'); if (name && rid[id]) out[nh(name.replace(/&amp;/g, '&'))] = resolve('xl/workbook.xml', rid[id]);
    });
    return out;
  }

  async function restoreNotes(JSZip, origBuf, outBuf) {
    try {
      const o = await JSZip.loadAsync(origBuf); const n = await JSZip.loadAsync(outBuf);
      const op = (await sheetPaths(o)).notes; const np = (await sheetPaths(n)).notes;
      if (!op || !np) return { restored: false, why: 'no Notes sheet' };

      const oSheet = await readText(o, op); const oRels = await readText(o, relsPathOf(op));
      if (!oSheet || !oRels) return { restored: false, why: 'no drawing in original' };
      const dTag = /<drawing\b[^>]*\/?>/.exec(oSheet); if (!dTag) return { restored: false, why: 'no drawing in original' };
      const dRid = attr(dTag[0], 'r:id');
      const dRel = (oRels.match(/<Relationship\b[^>]*>/g) || []).find((t) => attr(t, 'Id') === dRid);
      if (!dRel) return { restored: false, why: 'drawing rel missing' };
      const oDraw = resolve(op, attr(dRel, 'Target'));
      const drawXml = await readText(o, oDraw); if (!drawXml) return { restored: false, why: 'drawing part missing' };

      let nSheet = await readText(n, np); if (/<drawing\b/.test(nSheet)) return { restored: false, why: 'output already has a drawing' };

      const stamp = 'notes' + Date.now().toString(36);
      const newDraw = 'xl/drawings/' + stamp + 'Drawing.xml';
      const dRelsXml = await readText(o, relsPathOf(oDraw));
      if (dRelsXml) {
        let fixed = dRelsXml; let k = 0;
        for (const t of (dRelsXml.match(/<Relationship\b[^>]*>/g) || [])) {
          if (/TargetMode="External"/i.test(t)) continue;
          const src = resolve(oDraw, attr(t, 'Target')); const f = o.file(src); if (!f) continue;
          const ext = src.slice(src.lastIndexOf('.')); const dst = dirOf(src) + stamp + '_' + (k++) + ext;
          n.file(dst, await f.async('uint8array'));
          const base = dirOf(newDraw).split('/').filter(Boolean); const to = dst.split('/');
          let i = 0; while (i < base.length && i < to.length - 1 && base[i] === to[i]) i++;
          const rel = base.slice(i).map(() => '..').concat(to.slice(i)).join('/');
          fixed = fixed.replace(t, t.replace(/Target="[^"]*"/, 'Target="' + rel + '"'));
          const e = ext.slice(1).toLowerCase(); const ctype = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', emf: 'image/x-emf', wmf: 'image/x-wmf', svg: 'image/svg+xml' }[e];
          if (ctype) { let c = await readText(n, '[Content_Types].xml'); if (!new RegExp('Extension="' + e + '"', 'i').test(c)) { c = c.replace('</Types>', '<Default Extension="' + e + '" ContentType="' + ctype + '"/></Types>'); n.file('[Content_Types].xml', c); } }
        }
        n.file(relsPathOf(newDraw), fixed);
      }
      n.file(newDraw, drawXml);

      let ct = await readText(n, '[Content_Types].xml');
      ct = ct.replace('</Types>', '<Override PartName="/' + newDraw + '" ContentType="' + DRAW_CT + '"/></Types>'); n.file('[Content_Types].xml', ct);

      const nRelsPath = relsPathOf(np); let nRels = await readText(n, nRelsPath);
      const rId = 'rIdNotesDrw1';
      const relTag = '<Relationship Id="' + rId + '" Type="' + DRAW_TYPE + '" Target="/' + newDraw + '"/>';
      if (nRels) nRels = nRels.replace('</Relationships>', relTag + '</Relationships>');
      else nRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + relTag + '</Relationships>';
      n.file(nRelsPath, nRels);

      if (!/xmlns:r=/.test(nSheet.slice(0, 800))) nSheet = nSheet.replace(/<worksheet\b/, '<worksheet xmlns:r="' + REL_NS + '"');
      const tag = '<drawing r:id="' + rId + '"/>';
      const m = /<(legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/.exec(nSheet);
      if (m) nSheet = nSheet.slice(0, m.index) + tag + nSheet.slice(m.index);
      else nSheet = nSheet.replace('</worksheet>', tag + '</worksheet>');
      n.file(np, nSheet);

      const buf = await n.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
      return { restored: true, buf };
    } catch (e) { return { restored: false, why: String((e && e.message) || e) }; }
  }
  root.NotesCore = { restoreNotes };
})(typeof self !== 'undefined' ? self : this);
