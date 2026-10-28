/*
 * WordOps — semua operasi ke dokumen Word (Office.js, WordApi 1.3+).
 * Bergantung pada TplEngine (engine.js).
 */
(function (root) {
  "use strict";
  const E = root.TplEngine;
  const BOOKMARK = "_TplPasteLast";
  const LINK_COLOR = "#0563C1";

  const has = (v) => typeof Office !== "undefined" && Office.context && Office.context.requirements && Office.context.requirements.isSetSupported("WordApi", v);

  /* ---------------- gaya ---------------- */
  const BUILTIN = {
    "normal": "Normal", "no spacing": "NoSpacing", "nospacing": "NoSpacing", "caption": "Caption",
    "bibliography": "Bibliography", "quote": "Quote", "list paragraph": "ListParagraph", "title": "Title", "subtitle": "Subtitle",
  };
  function setStyle(p, name) {
    const k = String(name || "Normal").trim().toLowerCase().replace(/\s+/g, " ");
    const m = /^heading ?([1-9])$/.exec(k);
    if (m) p.styleBuiltIn = "Heading" + m[1];
    else if (BUILTIN[k]) p.styleBuiltIn = BUILTIN[k];
    else p.style = name; // nama gaya kustom — harus persis sama dengan di dokumen
  }

  function customStyleNames(profile) {
    const names = [profile.sectionHeading, profile.h2, profile.h3, profile.body, profile.caption, profile.note, profile.references]
      .map((c) => (c && c.style) || "Normal");
    return [...new Set(names.filter((n) => { const k = n.trim().toLowerCase(); return !BUILTIN[k] && !/^heading ?[1-9]$/.test(k); }))];
  }

  /** Ganti gaya kustom yang tidak ada di dokumen dengan Normal (bila API mendukung pengecekan). */
  async function validateStyles(context, profile) {
    const custom = customStyleNames(profile);
    if (!custom.length || !has("1.5")) return { profile, missing: [] };
    const styles = context.document.getStyles();
    styles.load("items/nameLocal");
    await context.sync();
    const exist = new Set(styles.items.map((s) => s.nameLocal.toLowerCase()));
    const missing = custom.filter((n) => !exist.has(n.toLowerCase()));
    if (!missing.length) return { profile, missing };
    const p = E.clone(profile);
    for (const k of ["sectionHeading", "h2", "h3", "body", "caption", "note", "references"]) {
      if (p[k] && missing.includes(p[k].style)) p[k].style = "Normal";
    }
    return { profile: p, missing };
  }

  /* ---------------- format paragraf baru ---------------- */
  function formatNewParagraph(p, it, profile) {
    const f = E.paraFormat(it, profile);
    setStyle(p, f.style);
    p.alignment = f.align;
    p.spaceBefore = f.before;
    p.spaceAfter = f.after;
    p.lineSpacing = 12 * f.line; // Office.js: nilai poin; UI Word membaginya dengan 12
    p.leftIndent = f.leftIndent;
    p.firstLineIndent = f.firstIndent;
    p.rightIndent = 0;
    // format tanda paragraf (mempengaruhi tinggi baris terakhir & teks yang diketik setelahnya)
    p.font.set({ name: profile.font, size: f.size, bold: f.bold, italic: f.italic, color: profile.color, underline: "None", superscript: false, subscript: false });
    p.font.highlightColor = null;
    if (it.type === "blank") return;
    for (const r of E.finalRuns(it, profile)) {
      if (!r.text) continue;
      const rg = p.insertText(r.text, "End");
      rg.font.set({
        name: profile.font, size: f.size, bold: !!r.b, italic: !!r.i,
        superscript: !!r.sup, subscript: !!r.sub,
        color: r.link ? LINK_COLOR : profile.color, underline: r.link ? "Single" : "None",
      });
      if (r.link) rg.hyperlink = r.link;
    }
  }

  function rectValues(rows) {
    const w = Math.max(1, ...rows.map((r) => r.length));
    return rows.map((r) => { const a = r.slice(0, w); while (a.length < w) a.push(""); return a.map((c) => String(c)); });
  }

  /** Format tabel APA (dipanggil setelah sync agar paragraf sel bisa dimuat). */
  function formatTable(t, rowsCount, profile) {
    const T = profile.table;
    t.paras.items.forEach((p) => {
      p.styleBuiltIn = "Normal";
      p.spaceBefore = 0; p.spaceAfter = 0; p.lineSpacing = 12; p.firstLineIndent = 0; p.leftIndent = 0;
    });
    const tbl = t.table;
    tbl.font.set({ name: profile.font, size: T.size, bold: false, italic: false, color: profile.color });
    tbl.headerRowCount = 1;
    if (T.apaBorders) {
      ["Left", "Right", "InsideHorizontal", "InsideVertical"].forEach((loc) => { tbl.getBorder(loc).type = "None"; });
      ["Top", "Bottom"].forEach((loc) => { const b = tbl.getBorder(loc); b.type = "Single"; b.width = 0.5; b.color = "#000000"; });
    }
    const head = tbl.rows.getFirst();
    head.font.bold = !!T.headerBold;
    if (T.apaBorders) { const hb = head.getBorder("Bottom"); hb.type = "Single"; hb.width = 0.5; hb.color = "#000000"; }
    tbl.horizontalAlignment = "Centered";
    if (T.firstColLeft) for (let r = 0; r < rowsCount; r++) tbl.getCell(r, 0).horizontalAlignment = "Left";
  }

  /**
   * Sisipkan items berurutan mulai dari `start` (Paragraph) di lokasi `loc` ('After'|'Before').
   * return {first, last, count, tables}
   */
  async function insertItems(context, start, loc, items, profile) {
    let cursor = start, cursorIsTable = false;
    let first = null, last = null, count = 0, sinceSync = 0;
    const tables = [];
    for (const it of items) {
      if (it.type === "skip") continue;
      if (it.type === "table") {
        const vals = rectValues(it.rows || []);
        if (!vals.length) continue;
        let host = cursor, hostLoc = loc;
        if (cursorIsTable) { // dua tabel berurutan: beri paragraf pemisah
          host = cursor.insertParagraph("", "After");
          formatNewParagraph(host, { type: "blank" }, profile);
          hostLoc = "After";
          first = first || host; last = host;
        }
        const tbl = host.insertTable(vals.length, vals[0].length, hostLoc, vals);
        tables.push({ table: tbl, rows: vals.length });
        first = first || tbl; last = tbl; cursor = tbl; cursorIsTable = true; loc = "After"; count++;
      } else {
        const p = cursor.insertParagraph("", loc);
        formatNewParagraph(p, it, profile);
        first = first || p; last = p; cursor = p; cursorIsTable = false; loc = "After"; count++;
      }
      if (++sinceSync >= 40) { await context.sync(); sinceSync = 0; }
    }
    await context.sync();
    if (tables.length) {
      for (const t of tables) { t.paras = t.table.getRange("Whole").paragraphs; t.paras.load("items"); }
      await context.sync();
      for (const t of tables) formatTable(t, t.rows, profile);
      await context.sync();
    }
    return { first, last, count, endsWithTable: cursorIsTable };
  }

  /* ---------------- membaca struktur dokumen ---------------- */
  async function readStructure(context, profile, withBold) {
    const paras = context.document.body.paragraphs;
    paras.load(withBold ? "items/text,items/tableNestingLevel,items/font/bold" : "items/text,items/tableNestingLevel");
    await context.sync();
    const infos = paras.items.map((p) => ({ text: p.text, inTable: p.tableNestingLevel > 0, bold: withBold ? p.font.bold : undefined }));
    const heads = E.findSectionHeadings(infos, profile);
    return { paras: paras.items, infos, heads };
  }

  async function markPictures(context, paras, infos, range) {
    const probe = [];
    for (let j = range.end - 1; j > range.start && probe.length < 12; j--) {
      if (infos[j].inTable) break;
      if ((infos[j].text || "").trim()) break;
      const c = paras[j].inlinePictures; c.load("items"); probe.push([j, c]);
    }
    if (!probe.length) return;
    await context.sync();
    probe.forEach(([j, c]) => { infos[j].hasPicture = c.items.length > 0; });
  }

  function bookmarkAndSelect(context, first, last) {
    if (!first) return;
    const rng = first.getRange("Start").expandTo(last.getRange("End"));
    if (has("1.4")) {
      try { context.document.deleteBookmark(BOOKMARK); } catch (e) { /* abaikan */ }
      rng.insertBookmark(BOOKMARK);
    }
    first.getRange("Start").select("Start");
  }

  /* ---------------- API publik ---------------- */

  /** Status bagian di dokumen: {key: true/false} + headings. */
  async function scan(profile) {
    return Word.run(async (context) => {
      const { infos, heads } = await readStructure(context, profile);
      const found = {};
      heads.forEach((h) => { if (h.key) found[h.key] = true; });
      return { found, count: E.countDocument(infos, heads, profile) };
    });
  }

  /**
   * Sisipkan items ke bagian `key`.
   * mode 'section': di akhir bagian (buat judul bagian bila belum ada)
   * mode 'cursor' : di posisi kursor
   */
  async function insertSection(key, items, profile0, mode) {
    return Word.run(async (context) => {
      const { profile, missing } = await validateStyles(context, profile0);
      const sec = profile.sections.find((s) => s.key === key);
      let start, loc, created = false, needTrailingBlank = false;
      let preFirst = null;

      if (mode === "cursor") {
        const sel = context.document.getSelection();
        const p = sel.paragraphs.getLast();
        p.load("text,tableNestingLevel");
        await context.sync();
        if (p.tableNestingLevel > 0) throw new Error("Kursor berada di dalam tabel. Letakkan kursor di luar tabel.");
        start = p; loc = p.text.trim() ? "After" : "Before";
      } else {
        const { paras, infos, heads } = await readStructure(context, profile);
        const range = E.sectionRange(infos, heads, key);
        if (range) {
          await markPictures(context, paras, infos, range);
          const a = E.pickAnchor(infos, range);
          start = paras[a.index]; loc = a.loc;
          needTrailingBlank = profile.layout.blankBeforeSection && range.nextIndex >= 0 && a.trailingBlanks === 0 && a.loc === "After";
          if (a.afterTable && a.loc === "Before") needTrailingBlank = profile.layout.blankBeforeSection;
        } else {
          // judul bagian belum ada -> buat
          const plan = E.planNewSection(heads, key, profile, paras.length);
          let heading;
          if (plan.loc === "Before") {
            const nextHead = paras[plan.index];
            heading = nextHead.insertParagraph("", "Before");
            const prevInfo = infos[plan.index - 1];
            if (prevInfo && (prevInfo.text || "").trim() && !prevInfo.inTable) {
              const blank = heading.insertParagraph("", "Before");
              formatNewParagraph(blank, { type: "blank" }, profile);
              preFirst = blank;
            }
            needTrailingBlank = profile.layout.blankBeforeSection;
          } else {
            const lastP = paras[paras.length - 1];
            const lastEmpty = !(infos[infos.length - 1].text || "").trim();
            if (lastEmpty) {
              heading = lastP.insertParagraph("", "Before");
              const prev = infos[infos.length - 2];
              if (prev && (prev.text || "").trim()) {
                const blank = heading.insertParagraph("", "Before");
                formatNewParagraph(blank, { type: "blank" }, profile);
                preFirst = blank;
              }
            } else {
              const blank = lastP.insertParagraph("", "After");
              formatNewParagraph(blank, { type: "blank" }, profile);
              preFirst = blank;
              heading = blank.insertParagraph("", "After");
            }
          }
          formatNewParagraph(heading, { type: "section", runs: [{ text: sec.title }] }, profile);
          preFirst = preFirst || heading;
          start = heading; loc = "After"; created = true;
        }
      }

      const res = await insertItems(context, start, loc, items, profile);
      let last = res.last;
      if (res.count && needTrailingBlank) {
        const blank = last.insertParagraph("", "After");
        formatNewParagraph(blank, { type: "blank" }, profile);
        last = blank;
      }
      const first = preFirst || res.first;
      bookmarkAndSelect(context, first, last || first);
      await context.sync();

      // hitung ulang kata
      const st = await readStructure(context, profile);
      return { inserted: res.count, created, missingStyles: missing, count: E.countDocument(st.infos, st.heads, profile) };
    });
  }

  /** Terapkan format template ke paragraf yang SUDAH ada di bagian `key`. */
  async function reformatSection(key, profile) {
    return Word.run(async (context) => {
      const sec = profile.sections.find((s) => s.key === key);
      const { paras, infos, heads } = await readStructure(context, profile, true);
      const range = E.sectionRange(infos, heads, key);
      if (!range) throw new Error(`Judul "${sec.title}" tidak ditemukan di dokumen.`);
      const idx = [];
      for (let j = range.start + 1; j < range.end; j++) if (!infos[j].inTable) idx.push(j);
      const items = E.classifyExisting(idx.map((j) => ({ text: infos[j].text, bold: infos[j].bold })), sec, profile);

      // judul bagian
      const hp = paras[range.start];
      applyExisting(hp, { type: "section" }, profile, infos[range.start].bold);
      if (profile.sectionHeading.caps) {
        const t = infos[range.start].text.trim();
        if (t !== t.toUpperCase()) hp.insertText(t.toUpperCase(), "Replace");
      }
      let n = 0;
      idx.forEach((j, k) => {
        const it = items[k];
        if (!it || it.type === "skip") return;
        applyExisting(paras[j], it, profile, infos[j].bold);
        n++;
      });
      await context.sync();
      return { formatted: n, headings: items.filter((i) => i && i.type === "heading").length };
    });
  }

  function applyExisting(p, it, profile, curBold) {
    const f = E.paraFormat(it, profile);
    p.alignment = f.align;
    p.spaceBefore = f.before;
    p.spaceAfter = f.after;
    p.lineSpacing = 12 * f.line;
    p.leftIndent = f.leftIndent;
    p.firstLineIndent = f.firstIndent;
    p.font.name = profile.font;
    p.font.size = f.size;
    if (it.type === "heading" || it.type === "section") {
      p.font.bold = f.bold;
      if (f.italic) p.font.italic = true;
    } else if (curBold === true && it.type !== "caption") {
      p.font.bold = false; // paragraf seluruhnya bold karena gaya template -> normalkan
    }
  }

  async function countWords(profile) {
    return Word.run(async (context) => {
      const { infos, heads } = await readStructure(context, profile);
      return E.countDocument(infos, heads, profile);
    });
  }

  async function undoLast() {
    if (!has("1.4")) throw new Error("Fitur ini butuh Word versi lebih baru (WordApi 1.4). Gunakan Ctrl+Z.");
    return Word.run(async (context) => {
      const rng = context.document.getBookmarkRangeOrNullObject(BOOKMARK);
      rng.load("isNullObject,text");
      await context.sync();
      if (rng.isNullObject) return false;
      rng.delete();
      context.document.deleteBookmark(BOOKMARK);
      await context.sync();
      return true;
    });
  }

  root.WordOps = { scan, insertSection, reformatSection, countWords, undoLast, _internal: { insertItems, formatNewParagraph, setStyle } };
})(typeof self !== "undefined" ? self : this);
