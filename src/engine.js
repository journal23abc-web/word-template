/*
 * TplEngine — mesin deteksi & format untuk add-in "Template Paste".
 * Murni JavaScript (tanpa Office.js) supaya bisa diuji di browser biasa.
 *
 * Alur:  HTML/teks tempelan  ->  blocks  ->  classify()  ->  items
 *        items = [{type, level, runs:[{text,b,i,sup,sub,link}], rows}]
 *        type: 'body' | 'heading' | 'caption' | 'note' | 'ref' | 'table' | 'skip'
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.TplEngine = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  /* ------------------------------------------------------------------ */
  /*  Profil default — diturunkan dari naskah copyedit AIEDS             */
  /* ------------------------------------------------------------------ */
  const DEFAULT_PROFILE = {
    name: "AIEDS (Artificial Intelligence in Educational Decision Sciences)",
    font: "Cambria",
    color: "#000000",
    limits: {
      bodyWords: 8000,        // batas kata isi artikel (Introduction s/d Conclusion)
      abstractWords: 250,
      countReferences: false, // referensi ikut dihitung?
      countTables: false,     // teks di dalam tabel ikut dihitung?
    },
    layout: {
      insertMode: "section",  // 'section' = di akhir bagian; 'cursor' = di posisi kursor
      blankBeforeSection: true, // sisakan 1 baris kosong sebelum judul bagian berikutnya
    },
    sectionHeading: { style: "Normal", size: 12, bold: true, italic: false, caps: true, align: "Centered", before: 0, after: 6, line: 1 },
    h2:      { style: "Normal", size: 11, bold: true,  italic: false, align: "Justified", before: 3, after: 3, line: 1 },
    h3:      { style: "Normal", size: 11, bold: true,  italic: false, align: "Justified", before: 3, after: 3, line: 1 },
    body:    { style: "Normal", size: 11, align: "Justified", firstIndentCm: 0, before: 3, after: 3, line: 1 },
    caption: { style: "Normal", size: 11, align: "Centered", labelBold: true, before: 3, after: 3, line: 1 },
    note:    { style: "Normal", size: 11, align: "Justified", labelItalic: false, before: 3, after: 4, line: 1 },
    table:   { size: 10, headerBold: true, apaBorders: true, firstColLeft: true },
    references: {
      style: "Normal", size: 11, align: "Justified", hangingCm: 1.27, before: 0, after: 6, line: 1,
      sort: false, autoItalic: true, linkDoi: true, stripNumbering: true,
    },
    detect: {
      maxHeadingWords: 15,
      stripHeadingNumbers: true,
      titleCaseHeadings: false,
      keepItalic: true,
      keepBold: true,
      keepSupSub: true,
      mergeBrokenLines: true,
      requireCapsSectionTitle: true,
    },
    sections: [
      { key: "intro",   title: "INTRODUCTION", aliases: ["Introduction", "Pendahuluan", "Background"] },
      { key: "method",  title: "METHOD", aliases: ["Methods", "Methodology", "Research Method", "Research Methods", "Materials and Methods", "Metode", "Metode Penelitian", "Metodologi"] },
      { key: "rd",      title: "RESULTS AND DISCUSSION", aliases: ["Results and Discussions", "Result and Discussion", "Results & Discussion", "Hasil dan Pembahasan"],
        majorSubheadings: ["Results", "Result", "Findings", "Discussion", "Discussions", "Hasil", "Hasil Penelitian", "Pembahasan"] },
      { key: "concl",   title: "CONCLUSION", aliases: ["Conclusions", "Kesimpulan", "Simpulan", "Conclusion and Recommendations"] },
      { key: "contrib", title: "AUTHOR CONTRIBUTION STATEMENT", aliases: ["Author Contributions", "Author Contribution", "Kontribusi Penulis"] },
      { key: "ai",      title: "AI DISCLOSURE STATEMENT", aliases: ["AI Disclosure", "Declaration of AI Use"] },
      { key: "refs",    title: "REFERENCES", aliases: ["Reference", "Bibliography", "Daftar Pustaka", "Referensi", "Works Cited"], kind: "references" },
    ],
  };

  const clone = (o) => JSON.parse(JSON.stringify(o));

  /** Gabungkan profil tersimpan dengan default (agar kunci baru tetap ada). */
  function mergeProfile(p) {
    const out = clone(DEFAULT_PROFILE);
    if (!p) return out;
    for (const k of Object.keys(p)) {
      if (p[k] && typeof p[k] === "object" && !Array.isArray(p[k]) && out[k] && typeof out[k] === "object" && !Array.isArray(out[k])) {
        out[k] = Object.assign(out[k], p[k]);
      } else out[k] = p[k];
    }
    return out;
  }

  /* ------------------------------------------------------------------ */
  /*  Utilitas teks                                                      */
  /* ------------------------------------------------------------------ */
  const WS = /[\s  -​  　]+/g;
  const countWords = (t) => (t || "").replace(WS, " ").trim().split(" ").filter((w) => /[\p{L}\p{N}]/u.test(w)).length;
  const plain = (runs) => runs.map((r) => r.text).join("");
  const sameFmt = (a, b) => !!a.b === !!b.b && !!a.i === !!b.i && !!a.sup === !!b.sup && !!a.sub === !!b.sub && (a.link || "") === (b.link || "");

  /** Rapikan runs: spasi ganda -> satu, trim ujung, gabung run berformat sama. */
  function normalizeRuns(runs) {
    const out = [];
    for (const r0 of runs) {
      if (!r0 || r0.text == null) continue;
      const r = Object.assign({}, r0, { text: String(r0.text).replace(/[\r\n\t\v\f]+/g, " ").replace(WS, " ") });
      if (!r.text) continue;
      const prev = out[out.length - 1];
      if (prev && sameFmt(prev, r)) prev.text += r.text;
      else out.push(r);
    }
    // collapse spasi lintas run
    for (let k = 1; k < out.length; k++) {
      if (out[k - 1].text.endsWith(" ") && out[k].text.startsWith(" ")) out[k].text = out[k].text.slice(1);
    }
    if (out.length) out[0].text = out[0].text.replace(/^ +/, "");
    if (out.length) out[out.length - 1].text = out[out.length - 1].text.replace(/ +$/, "");
    return out.filter((r) => r.text.length);
  }

  /** Buang n karakter pertama dari runs. */
  function dropPrefix(runs, n) {
    const out = [];
    for (const r of runs) {
      if (n <= 0) { out.push(Object.assign({}, r)); continue; }
      if (r.text.length <= n) { n -= r.text.length; continue; }
      out.push(Object.assign({}, r, { text: r.text.slice(n) }));
      n = 0;
    }
    return normalizeRuns(out);
  }

  /** Terapkan fungsi fmt(r) pada rentang karakter [s,e) dari runs (split bila perlu). */
  function patchRange(runs, s, e, patch) {
    const out = [];
    let pos = 0;
    for (const r of runs) {
      const a = pos, b = pos + r.text.length;
      pos = b;
      if (b <= s || a >= e) { out.push(r); continue; }
      const cut1 = Math.max(s, a) - a, cut2 = Math.min(e, b) - a;
      if (cut1 > 0) out.push(Object.assign({}, r, { text: r.text.slice(0, cut1) }));
      out.push(Object.assign({}, r, patch, { text: r.text.slice(cut1, cut2) }));
      if (cut2 < r.text.length) out.push(Object.assign({}, r, { text: r.text.slice(cut2) }));
    }
    return out;
  }

  function mapRunsText(runs, fn) {
    return runs.map((r) => Object.assign({}, r, { text: fn(r.text) }));
  }

  const MINOR = new Set("a an and as at but by for from in into nor of on or per the to vs via with yang dan di ke dari atau untuk pada dalam dengan".split(" "));
  function titleCase(s) {
    let first = true;
    return s.replace(/[\p{L}][\p{L}'’\-]*/gu, (w, off) => {
      const lower = w.toLowerCase();
      const prevChar = s[off - 2] || "";
      const forceCap = first || /[:—–-]/.test(prevChar);
      first = false;
      if (/[A-Z].*[A-Z]/.test(w)) return w; // akronim (FCM, AI, SEM-PLS)
      if (!forceCap && MINOR.has(lower)) return lower;
      return lower.charAt(0).toUpperCase() + lower.slice(1);
    });
  }

  /* ------------------------------------------------------------------ */
  /*  Parser HTML tempelan (Word, Google Docs, ChatGPT, web)             */
  /* ------------------------------------------------------------------ */
  const BLOCK = new Set(["P", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "LI", "BLOCKQUOTE", "PRE", "SECTION", "ARTICLE", "HEADER", "FOOTER", "DD", "DT", "FIGCAPTION", "CAPTION", "ADDRESS", "ASIDE", "MAIN", "NAV", "UL", "OL", "DL", "HR", "CENTER"]);
  const SKIP = new Set(["STYLE", "SCRIPT", "HEAD", "TITLE", "META", "LINK", "NOSCRIPT", "TEMPLATE", "SVG", "MATH", "BUTTON", "OBJECT"]);

  function styleProp(el, prop) {
    const st = el.getAttribute && el.getAttribute("style");
    if (!st) return null;
    const m = new RegExp("(?:^|;)\\s*" + prop + "\\s*:\\s*([^;]+)", "i").exec(st);
    return m ? m[1].trim().toLowerCase() : null;
  }

  function fmtFrom(el, parent) {
    const f = Object.assign({}, parent);
    const t = el.tagName;
    if (t === "B" || t === "STRONG") f.b = true;
    if (t === "I" || t === "EM" || t === "CITE" || t === "VAR" || t === "DFN") f.i = true;
    if (t === "SUP") f.sup = true;
    if (t === "SUB") f.sub = true;
    if (t === "A" && el.getAttribute("href") && /^https?:/i.test(el.getAttribute("href"))) f.href = el.getAttribute("href");
    const fw = styleProp(el, "font-weight");
    if (fw) f.b = /bold|bolder|[6-9]00/.test(fw) ? true : /normal|lighter|[1-4]00/.test(fw) ? false : f.b;
    const fs = styleProp(el, "font-style");
    if (fs) f.i = /italic|oblique/.test(fs) ? true : fs === "normal" ? false : f.i;
    const va = styleProp(el, "vertical-align");
    if (va) { if (/super/.test(va)) f.sup = true; else if (/sub/.test(va)) f.sub = true; else if (va === "baseline") { f.sup = false; f.sub = false; } }
    return f;
  }

  /**
   * @param {string} html
   * @param {DOMParser} [ParserImpl]
   * @returns {{blocks:Array, images:number, tables:number}}
   */
  function parseHtml(html, ParserImpl) {
    const P = ParserImpl || (typeof DOMParser !== "undefined" ? DOMParser : null);
    // Buang komentar kondisional Word & fragment marker
    html = String(html).replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, "").replace(/<!--[\s\S]*?-->/g, "");
    const doc = new P().parseFromString(html, "text/html");
    const blocks = [];
    let cur = null;
    let images = 0, tables = 0;
    const olStack = [];

    const flush = () => {
      if (cur && cur.runs.length) {
        const runs = normalizeRuns(cur.runs);
        if (runs.length) blocks.push({ kind: "p", tag: cur.tag, cls: cur.cls, runs });
      }
      cur = null;
    };
    const ensure = (tag, cls) => { if (!cur) cur = { tag: tag || "p", cls: cls || "", runs: [] }; };

    function walk(node, fmt, blockTag, blockCls) {
      if (node.nodeType === 3) {
        const txt = node.nodeValue;
        if (!txt) return;
        if (!cur && !/\S/.test(txt)) return;
        ensure(blockTag, blockCls);
        cur.runs.push({ text: txt, b: !!fmt.b, i: !!fmt.i, sup: !!fmt.sup, sub: !!fmt.sub, link: fmt.href || "" });
        return;
      }
      if (node.nodeType !== 1) return;
      const el = node;
      const tag = el.tagName.toUpperCase();
      if (SKIP.has(tag)) return;
      if (/mso-list\s*:\s*ignore/i.test(el.getAttribute("style") || "")) return; // nomor/bullet palsu dari Word
      if (tag === "IMG" || tag === "PICTURE" || tag === "CANVAS" || tag === "V:SHAPE" || tag === "V:IMAGEDATA") { images++; return; }
      if (tag === "BR") { flush(); return; }
      if (tag === "TABLE") {
        flush();
        const rows = [];
        el.querySelectorAll("tr").forEach((tr) => {
          if (tr.closest("table") !== el) return; // abaikan tabel bersarang
          const cells = [];
          tr.querySelectorAll("td,th").forEach((td) => {
            if (td.closest("tr") !== tr) return;
            const t = (td.textContent || "").replace(WS, " ").trim();
            cells.push(t);
            const cs = parseInt(td.getAttribute("colspan") || "1", 10);
            for (let k = 1; k < cs && k < 30; k++) cells.push("");
          });
          if (cells.length) rows.push(cells);
        });
        if (rows.length) { tables++; blocks.push({ kind: "table", rows }); }
        return;
      }
      const f = fmtFrom(el, fmt);
      const isBlock = BLOCK.has(tag) || /^(block|list-item|table|flex|grid)/.test(styleProp(el, "display") || "");
      if (tag === "OL") olStack.push(0);
      if (isBlock) {
        flush();
        let bt = /^H[1-6]$/.test(tag) ? tag.toLowerCase() : tag === "LI" ? "li" : "p";
        const cls = el.getAttribute("class") || "";
        const m = /MsoHeading(\d)|Heading(\d)/i.exec(cls);
        if (m) bt = "h" + (m[1] || m[2]);
        if (/MsoTitle/i.test(cls)) bt = "h1";
        const role = el.getAttribute("role"), aria = el.getAttribute("aria-level");
        if (role === "heading" && aria) bt = "h" + aria;
        let prefix = "";
        if (tag === "LI" && olStack.length && el.parentElement && el.parentElement.tagName === "OL") {
          olStack[olStack.length - 1]++;
          prefix = olStack[olStack.length - 1] + ". ";
        }
        if (prefix) { ensure(bt, cls); cur.runs.push({ text: prefix, b: false, i: false }); }
        for (const c of el.childNodes) walk(c, f, bt, cls);
        flush();
      } else {
        for (const c of el.childNodes) walk(c, f, blockTag, blockCls);
      }
      if (tag === "OL") olStack.pop();
    }

    const bodyEl = doc.body || doc.documentElement;
    for (const c of bodyEl.childNodes) walk(c, {}, "p", "");
    flush();
    return { blocks, images, tables, source: "html" };
  }

  /* ------------------------------------------------------------------ */
  /*  Parser teks biasa (Notepad, PDF, Markdown/ChatGPT)                 */
  /* ------------------------------------------------------------------ */
  function mdInline(text) {
    // **bold**, *italic*, _italic_ — hanya bila diapit karakter non-spasi
    const runs = [];
    const re = /(\*\*|__)(?=\S)([\s\S]+?)(?<=\S)\1|(?<![\w*])(\*|_)(?=\S)([^*_\n]+?)(?<=\S)\3(?![\w*])/g;
    let last = 0, m;
    while ((m = re.exec(text))) {
      if (m.index > last) runs.push({ text: text.slice(last, m.index) });
      if (m[1]) runs.push({ text: m[2], b: true });
      else runs.push({ text: m[4], i: true });
      last = re.lastIndex;
    }
    if (last < text.length) runs.push({ text: text.slice(last) });
    return runs;
  }

  const endsSentence = (s) => /[.!?:;)"”’\]]\s*$/.test(s);
  const startsLower = (s) => /^[\p{Ll}(\[,;–-]/u.test(s.trim());

  function parseText(text, opts) {
    opts = opts || {};
    const merge = opts.mergeBrokenLines !== false;
    text = String(text).replace(/\r\n?/g, "\n").replace(/\u2028|\u2029/g, "\n");
    const rawLines = text.split("\n");
    const hasBlankSep = /\n[ \t ]*\n/.test(text.trim());
    let paras = [];
    if (hasBlankSep) {
      // paragraf dipisah baris kosong; baris di dalamnya disambung (khas PDF)
      let acc = [];
      for (const ln of rawLines) {
        if (/^\s*$/.test(ln)) { if (acc.length) paras.push(acc); acc = []; }
        else acc.push(ln.trim());
      }
      if (acc.length) paras.push(acc);
      paras = paras.map((ls) => {
        // jangan sambung bila baris pertama heading markdown
        if (ls.length > 1 && /^#{1,6}\s/.test(ls[0])) return [ls[0], ls.slice(1).join(" ")];
        return [joinLines(ls)];
      }).flat();
    } else {
      const ls = rawLines.map((l) => l.trim()).filter((l) => l.length);
      if (merge && ls.length > 2) paras = mergeLines(ls);
      else paras = ls;
    }
    const blocks = [];
    for (let p of paras) {
      let tag = "p";
      const h = /^(#{1,6})\s+(.*)$/.exec(p);
      if (h) { tag = "h" + h[1].length; p = h[2]; }
      p = p.replace(/^\s*(?:[-*•▪◦·]\s+)/, ""); // bullet
      const runs = normalizeRuns(mdInline(p).map((r) => ({ text: r.text, b: !!r.b, i: !!r.i, sup: false, sub: false, link: "" })));
      if (runs.length) blocks.push({ kind: "p", tag, cls: "", runs, md: !!h });
    }
    return { blocks, images: 0, tables: 0, source: "text" };
  }

  function joinLines(ls) {
    return ls.reduce((acc, l) => {
      if (!acc) return l;
      if (/-$/.test(acc) && /^\p{Ll}/u.test(l)) return acc.slice(0, -1) + l; // pemenggalan kata
      return acc + " " + l;
    }, "");
  }

  function mergeLines(ls) {
    const longs = ls.filter((l) => l.length > 40).map((l) => l.length).sort((a, b) => a - b);
    const median = longs.length ? longs[Math.floor(longs.length / 2)] : 0;
    // Satu baris = satu paragraf (salinan dari Word/web) bila banyak baris sangat panjang.
    // Baris hasil salin PDF umumnya pendek & seragam (< 130 karakter).
    const veryLong = ls.filter((l) => l.length > 160).length;
    if (veryLong >= Math.max(1, ls.length * 0.25)) {
      // tetap sambung baris yang jelas lanjutan (diawali huruf kecil)
      const out0 = [];
      for (const l of ls) {
        const prev = out0[out0.length - 1];
        if (prev != null && !endsSentence(prev) && /^\p{Ll}/u.test(l)) out0[out0.length - 1] = joinLines([prev, l]);
        else out0.push(l);
      }
      return out0;
    }
    const out = [];
    for (const l of ls) {
      const prev = out[out.length - 1];
      if (prev != null && !/^#{1,6}\s/.test(prev) && !endsSentence(prev) &&
          (startsLower(l) || (median > 50 && prev.length >= median * 0.8 && prev.length > 50))) {
        out[out.length - 1] = joinLines([prev, l]);
      } else out.push(l);
    }
    return out;
  }

  /**
   * Titik masuk tunggal: data clipboard -> blocks.
   * @param {{html?:string, text?:string}} clip
   */
  function parseInput(clip, sec, profile, ParserImpl) {
    profile = profile || DEFAULT_PROFILE;
    const isRef = sec && sec.kind === "references";
    const html = clip.html || "";
    // HTML dipakai bila memuat struktur paragraf; jika tidak, pakai teks biasa
    if (html && /<(p|div|h[1-6]|li|br|table|tr)\b/i.test(html)) {
      const r = parseHtml(html, ParserImpl);
      if (r.blocks.length) return r;
    }
    return parseText(clip.text || "", { mergeBrokenLines: profile.detect.mergeBrokenLines && !isRef });
  }

  /* ------------------------------------------------------------------ */
  /*  Klasifikasi                                                        */
  /* ------------------------------------------------------------------ */
  const RE_CAPTION = /^(Table|Tables|Figure|Fig\.|Tabel|Gambar|Grafik|Chart|Diagram)\s+(\d+[A-Za-z]?|[IVX]+)\s*[.:]/i;
  const RE_NOTE = /^(Notes?|Catatan|Sumber|Source|Keterangan)\s*[.:]/i;
  const RE_CITATION = /\([^()]*\b(1[89]|20)\d{2}[a-z]?\b[^()]*\)|\bet al\./i;
  const RE_NUMBERED = /^((?:\d+\.)*\d+)\.?\s+(?=\S)|^([A-H]|[IVX]{1,4})[.)]\s+(?=\p{Lu})/u;

  const normHead = (s) => String(s || "").replace(WS, " ").trim()
    .replace(/^(?:(?:\d+\.)*\d+|[IVX]{1,5}|[A-H])[.)]?\s+/, "")
    .replace(/[.:;]+$/, "").trim().toUpperCase();

  function sectionNames(sec) {
    return [sec.title].concat(sec.aliases || []).map(normHead);
  }

  function looksLikeHeading(text, runs, block, det) {
    const t = text.trim();
    const words = countWords(t);
    if (!words || words > (det.maxHeadingWords || 15) || t.length > 160) return false;
    if (/^\d+(\.\d+)*\.?$/.test(t)) return false;
    if (RE_CAPTION.test(t) || RE_NOTE.test(t)) return false;
    const letters = (t.match(/\p{L}/gu) || []).length;
    const digits = (t.match(/\d/g) || []).length;
    if (letters < 3 || digits > letters * 0.5) return false;
    // sinyal format kuat dari HTML
    if (block.tag && /^h[1-6]$/.test(block.tag)) return true;
    const allBold = runs.length && runs.every((r) => r.b || !/\S/.test(r.text));
    const numbered = RE_NUMBERED.test(t);
    if (/[.,;]$/.test(t) && !(allBold && /\.$/.test(t) && words <= 8)) return false;
    if (/\?$/.test(t) && !allBold) return false;
    if (/:$/.test(t) && !allBold) return false;
    if (RE_CITATION.test(t)) return false;
    if (!/^[\p{Lu}\d"“'‘(\[]/u.test(t)) return false;
    if (allBold || numbered) return true;
    // teks biasa: pendek, tanpa tanda titik akhir, huruf besar di awal
    return words <= (det.maxHeadingWords || 15);
  }

  function numberingDepth(t) {
    const m = /^((?:\d+\.)*\d+)\.?\s+/.exec(t);
    if (m) return m[1].split(".").length;
    if (/^[A-H][.)]\s+/.test(t)) return 2;
    if (/^[IVX]{1,4}[.)]\s+/.test(t)) return 1;
    return 0;
  }

  /**
   * @param {Array} blocks hasil parseHtml/parseText
   * @param {object} sec   konfigurasi bagian (dari profile.sections)
   * @param {object} profile
   */
  function classify(blocks, sec, profile) {
    profile = profile || DEFAULT_PROFILE;
    const det = profile.detect;
    sec = sec || {};
    if (sec.kind === "references") return classifyReferences(blocks, sec, profile);

    const myNames = new Set(sectionNames(sec));
    const majors = new Set((sec.majorSubheadings || []).map(normHead));
    const items = [];

    for (let idx = 0; idx < blocks.length; idx++) {
      const bl = blocks[idx];
      if (bl.kind === "table") { items.push({ type: "table", rows: bl.rows, runs: [] }); continue; }
      let runs = bl.runs;
      const text = plain(runs);
      if (!text.trim()) continue;
      if (myNames.has(normHead(text)) && countWords(text) <= 8) { items.push({ type: "skip", runs, why: "judul bagian" }); continue; }
      if (RE_CAPTION.test(text)) { items.push({ type: "caption", runs }); continue; }
      if (RE_NOTE.test(text)) { items.push({ type: "note", runs }); continue; }
      // heading tidak boleh paragraf terakhir yang berdiri sendiri? -> tetap boleh, tapi butuh sinyal kuat
      const isLast = idx === blocks.length - 1;
      if (looksLikeHeading(text, runs, bl, det)) {
        const strong = (bl.tag && /^h[1-6]$/.test(bl.tag)) || runs.every((r) => r.b || !/\S/.test(r.text)) || RE_NUMBERED.test(text.trim());
        if (!isLast || strong) {
          const next = blocks[idx + 1];
          const nextText = next && next.kind === "p" ? plain(next.runs) : "";
          // dua baris pendek berturut-turut tanpa sinyal kuat -> kemungkinan daftar, bukan heading
          const nextAlsoShort = nextText && countWords(nextText) <= 4 && !/^h[1-6]$/.test(next.tag || "");
          if (strong || !nextAlsoShort || countWords(nextText) === 0 || (next && next.kind === "table")) {
            items.push({ type: "heading", level: 0, runs, tag: bl.tag, depth: numberingDepth(text.trim()), major: majors.has(normHead(text)) });
            continue;
          }
        }
      }
      items.push({ type: "body", runs });
    }

    // Tentukan level heading
    const heads = items.filter((it) => it.type === "heading");
    const tagLevels = [...new Set(heads.map((h) => (h.tag && /^h[1-6]$/.test(h.tag)) ? +h.tag[1] : 99))].filter((x) => x < 99).sort((a, b) => a - b);
    const depths = [...new Set(heads.map((h) => h.depth).filter((d) => d > 0))].sort((a, b) => a - b);
    let seenMajor = false;
    for (const h of heads) {
      let lvl = 2;
      if (h.major) { lvl = 2; seenMajor = true; }
      else if (seenMajor) lvl = 3;
      else if (h.depth > 0 && depths.length > 1) lvl = h.depth === depths[0] ? 2 : 3;
      else if (h.tag && /^h[1-6]$/.test(h.tag) && tagLevels.length > 1) lvl = (+h.tag[1] === tagLevels[0]) ? 2 : 3;
      h.level = lvl;
      // bersihkan teks heading
      let r = h.runs;
      const t = plain(r);
      if (det.stripHeadingNumbers) {
        const m = RE_NUMBERED.exec(t);
        if (m) r = dropPrefix(r, m[0].length);
      }
      r = mapRunsText(r, (s) => s.replace(/^\*+|\*+$/g, ""));
      r = normalizeRuns(r).map((x) => Object.assign({}, x, { text: x.text }));
      if (det.titleCaseHeadings) r = mapRunsText(r, titleCase);
      // buang titik/titik dua di akhir heading
      if (r.length) r[r.length - 1].text = r[r.length - 1].text.replace(/[.:]\s*$/, "");
      h.runs = normalizeRuns(r);
      delete h.tag; delete h.depth; delete h.major;
    }
    return items;
  }

  /* ------------------------------------------------------------------ */
  /*  Referensi                                                          */
  /* ------------------------------------------------------------------ */
  const RE_REF_START = /^(?:\[\d+\]\s*|\d+\.\s+)?(?:[\p{Lu}][\p{L}'’\-]+(?:[ -][\p{L}'’\-]+){0,3},\s*(?:[\p{Lu}]\.|[\p{Lu}][\p{Ll}]+)|[\p{Lu}][\p{L}&.,'’\- ]{1,120}?\s\((?:\d{4}[a-z]?|n\.d\.|in press)[^)]*\)|[\p{Lu}][\p{L}'’\-]+(?: [\p{Lu}][\p{L}'’\-]+)*\.?\s\((?:\d{4}|n\.d\.))/u;

  function splitRefBlocks(blocks, det) {
    // pecah blok menjadi baris; untuk teks biasa sambung baris-baris lanjutan
    let lines = [];
    for (const b of blocks) {
      if (b.kind === "table") { b.rows.forEach((row) => lines.push({ runs: normalizeRuns([{ text: row.join(" ") }]) })); continue; }
      lines.push({ runs: b.runs });
    }
    lines = lines.filter((l) => plain(l.runs).trim());
    if (!det.mergeBrokenLines || lines.length < 2) return lines;
    const starts = lines.map((l) => RE_REF_START.test(plain(l.runs).trim()));
    const nonStart = starts.filter((s) => !s).length;
    if (nonStart === 0) return lines;
    // mode sambung: baris yang bukan awal referensi digabung ke referensi sebelumnya
    const out = [];
    lines.forEach((l, k) => {
      const t = plain(l.runs).trim();
      const prev = out[out.length - 1];
      const prevT = prev ? plain(prev.runs).trim() : "";
      const prevLine = k > 0 ? plain(lines[k - 1].runs).trim() : "";
      const prevEndsRef = /https?:\/\/\S+$|doi\.org\/\S+$/i.test(prevLine) ||
        (/[.)]$/.test(prevLine) && !/(?:^|[\s,(])\p{Lu}\.$/u.test(prevLine)); // "…, D." = inisial penulis, belum selesai
      const continuation = prev && (!starts[k] || startsLower(t) || !prevEndsRef || /[,&]$|\band$/.test(prevLine));
      if (continuation) {
        const glue = /-$/.test(prevT) && /^\p{Ll}/u.test(t) ? "" : " ";
        if (!glue) prev.runs[prev.runs.length - 1].text = prev.runs[prev.runs.length - 1].text.replace(/-$/, "");
        prev.runs = normalizeRuns(prev.runs.concat([{ text: glue }], l.runs));
      } else out.push({ runs: l.runs.slice() });
    });
    return out;
  }

  /** Miringkan "Nama Jurnal, Volume" (APA 7) bila tempelan tidak membawa italic. */
  function apaAutoItalic(runs) {
    if (runs.some((r) => r.i)) return runs;
    const t = plain(runs);
    // ... (2020). Judul artikel. Nama Jurnal, 12(3), 45–67.
    const yr = /\((?:\d{4}[a-z]?|n\.d\.|in press)[^)]*\)\.\s+/i.exec(t);
    if (!yr) return runs;
    const after = yr.index + yr[0].length;
    const rest = t.slice(after);
    const j = /^(.+?[.?!])\s+([^.?!]*?\S),\s*(\d+)(?=\s*(?:\(|,|\.|$))/.exec(rest);
    if (j && j[2].length > 2 && !/^(In|Dalam)\s/.test(j[2]) && !/https?:/i.test(j[2])) {
      const jStart = rest.indexOf(j[2], j[1].length);
      const volEnd = rest.indexOf(j[3], jStart + j[2].length) + j[3].length;
      if (jStart > 0 && volEnd > jStart) return patchRange(runs, after + jStart, after + volEnd, { i: true });
    }
    // Buku: (2020). Judul buku (edisi). Penerbit.  -> miringkan judul bila tidak ada pola jurnal
    const bk = /^([^.]+(?:\.[^.\s][^.]*)*?)(?:\s\([^)]*ed\.[^)]*\))?\.\s+[\p{Lu}][^.]+\.?\s*(?:https?:\/\/\S+)?\s*$/u.exec(rest);
    if (bk && !/\d+\(\d+\)/.test(rest) && !/,\s*\d+\s*[,(]/.test(rest) && !/^In\s/i.test(rest)) {
      return patchRange(runs, after, after + bk[1].length, { i: true });
    }
    return runs;
  }

  function linkUrls(runs) {
    const t = plain(runs);
    const re = /\bhttps?:\/\/[^\s<>"]+[^\s<>".,;)\]]|\bdoi:\s*(10\.\d{4,9}\/\S+[^\s.,;)])/gi;
    let m, out = runs;
    const spans = [];
    while ((m = re.exec(t))) spans.push([m.index, m.index + m[0].length, m[1] ? "https://doi.org/" + m[1] : m[0]]);
    for (const [s, e, url] of spans) out = patchRange(out, s, e, { link: url });
    return out;
  }

  function classifyReferences(blocks, sec, profile) {
    const det = profile.detect, rc = profile.references;
    const myNames = new Set(sectionNames(sec));
    let refs = splitRefBlocks(blocks, det);
    const items = [];
    for (const r of refs) {
      let runs = normalizeRuns(r.runs);
      const t = plain(runs).trim();
      if (!t) continue;
      if (myNames.has(normHead(t)) && countWords(t) <= 4) { items.push({ type: "skip", runs, why: "judul bagian" }); continue; }
      if (rc.stripNumbering) {
        const m = /^(?:\[\d+\]|\d+\.(?!\d))\s*/.exec(t);
        if (m) runs = dropPrefix(runs, m[0].length);
      }
      // normalisasi doi: -> https://doi.org/
      runs = mapRunsText(runs, (s) => s.replace(/\bdoi:\s*(10\.)/gi, "https://doi.org/$1"));
      if (rc.autoItalic) runs = apaAutoItalic(runs);
      if (rc.linkDoi) runs = linkUrls(runs);
      items.push({ type: "ref", runs: normalizeRuns(runs) });
    }
    if (rc.sort) {
      const key = (it) => plain(it.runs).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
      const skips = items.filter((i) => i.type !== "ref");
      const rs = items.filter((i) => i.type === "ref").sort((a, b) => key(a).localeCompare(key(b)));
      return skips.concat(rs);
    }
    return items;
  }

  /* ------------------------------------------------------------------ */
  /*  Format akhir per item (dipakai saat menyisipkan ke Word)           */
  /* ------------------------------------------------------------------ */
  const CM = 28.3464567;

  /** Format paragraf (satuan pt) untuk satu item. */
  function paraFormat(it, profile) {
    const P = profile;
    let c, first = 0, left = 0;
    switch (it.type) {
      case "section": c = P.sectionHeading; break;
      case "heading": c = it.level === 3 ? P.h3 : P.h2; break;
      case "caption": c = P.caption; break;
      case "note": c = P.note; break;
      case "ref": c = P.references; left = (c.hangingCm || 0) * CM; first = -left; break;
      case "blank": c = Object.assign({}, P.body, { before: 0, after: 0 }); break;
      default: c = P.body; first = (c.firstIndentCm || 0) * CM;
    }
    return {
      style: c.style || "Normal",
      size: c.size || P.body.size,
      align: c.align || "Justified",
      before: +c.before || 0,
      after: +c.after || 0,
      line: +c.line || 1,
      firstIndent: Math.round(first * 100) / 100,
      leftIndent: Math.round(left * 100) / 100,
      bold: !!c.bold,
      italic: !!c.italic,
    };
  }

  /** Runs final (format karakter) untuk satu item. */
  function finalRuns(it, profile) {
    const det = profile.detect;
    let runs = (it.runs || []).map((r) => Object.assign({}, r));
    if (it.type === "section") {
      const t = plain(runs);
      return [{ text: profile.sectionHeading.caps ? t.toUpperCase() : t, b: !!profile.sectionHeading.bold, i: !!profile.sectionHeading.italic }];
    }
    if (it.type === "heading") {
      const c = it.level === 3 ? profile.h3 : profile.h2;
      return runs.map((r) => ({ text: r.text, b: !!c.bold, i: !!c.italic || (det.keepItalic && !!r.i), sup: det.keepSupSub && r.sup, sub: det.keepSupSub && r.sub }));
    }
    runs = runs.map((r) => ({
      text: r.text,
      b: det.keepBold && !!r.b,
      i: det.keepItalic && !!r.i,
      sup: det.keepSupSub && !!r.sup,
      sub: det.keepSupSub && !!r.sub,
      link: r.link || "",
    }));
    // Paragraf isi yang seluruhnya bold (mis. dari style template) -> bukan bold
    if (it.type === "body" && runs.length && runs.every((r) => r.b) && countWords(plain(runs)) > (det.maxHeadingWords || 15)) runs.forEach((r) => (r.b = false));
    if (it.type === "caption") {
      const t = plain(runs);
      const m = RE_CAPTION.exec(t);
      runs.forEach((r) => (r.b = false));
      if (m && profile.caption.labelBold) runs = patchRange(runs, 0, m[0].length, { b: true });
    }
    if (it.type === "note") {
      const m = RE_NOTE.exec(plain(runs));
      if (m && profile.note.labelItalic) runs = patchRange(runs, 0, m[0].length - 1, { i: true });
    }
    if (it.type === "ref") runs.forEach((r) => (r.b = false));
    if ((it.type === "body" || it.type === "note") && profile.references.linkDoi) runs = runs.map((r) => r); // biarkan
    return normalizeRuns(runs);
  }

  /* ------------------------------------------------------------------ */
  /*  Analisis struktur dokumen (berdasarkan teks paragraf)              */
  /* ------------------------------------------------------------------ */
  /**
   * infos: [{text, inTable}]
   * return: [{index, key|null, title}]
   */
  function findSectionHeadings(infos, profile) {
    const det = profile.detect;
    const lookup = new Map();
    for (const s of profile.sections) for (const n of sectionNames(s)) lookup.set(n, s.key);
    const out = [];
    infos.forEach((p, i) => {
      if (p.inTable) return;
      const t = (p.text || "").replace(WS, " ").trim();
      if (!t || countWords(t) > 8) return;
      const letters = t.replace(/[^\p{L}]/gu, "");
      const isCaps = letters.length >= 3 && letters === letters.toUpperCase();
      const key = lookup.get(normHead(t));
      if (key && (!det.requireCapsSectionTitle || isCaps)) out.push({ index: i, key, title: t });
      else if (isCaps && letters.length >= 4 && !/[.,;]$/.test(t) && !/\d/.test(t)) out.push({ index: i, key: null, title: t });
    });
    return out;
  }

  /** Rentang [start, end) paragraf milik bagian `key` (start = indeks judul). */
  function sectionRange(infos, heads, key) {
    const k = heads.findIndex((h) => h.key === key);
    if (k < 0) return null;
    const next = heads[k + 1];
    return { start: heads[k].index, end: next ? next.index : infos.length, nextIndex: next ? next.index : -1 };
  }

  /**
   * Tentukan titik sisip di akhir bagian.
   * infos boleh punya hasPicture. Return {index, loc:'After'|'Before', trailingBlanks}
   */
  function pickAnchor(infos, range) {
    const { start, end } = range;
    for (let j = end - 1; j > start; j--) {
      const p = infos[j];
      if (p.inTable) {
        // akhir tabel: paragraf setelah tabel
        const a = j + 1;
        if (a >= end) return { index: end < infos.length ? end : j, loc: end < infos.length ? "Before" : "After", trailingBlanks: 0, afterTable: true };
        return { index: a, loc: "After", trailingBlanks: Math.max(0, end - a - 1), afterTable: true };
      }
      if ((p.text || "").replace(WS, "").length || p.hasPicture) {
        let blanks = 0;
        for (let q = j + 1; q < end; q++) if (!(infos[q].text || "").replace(WS, "").length) blanks++;
        return { index: j, loc: "After", trailingBlanks: blanks };
      }
    }
    return { index: start, loc: "After", trailingBlanks: Math.max(0, end - start - 1) };
  }

  /** Di mana judul bagian yang belum ada harus dibuat. */
  function planNewSection(heads, key, profile, infosLen) {
    const order = profile.sections.map((s) => s.key);
    const myPos = order.indexOf(key);
    // judul bagian berikutnya (menurut urutan profil) yang sudah ada di dokumen
    for (const h of heads) {
      if (!h.key) continue;
      const pos = order.indexOf(h.key);
      if (pos > myPos) return { index: h.index, loc: "Before" };
    }
    return { index: infosLen - 1, loc: "End" };
  }

  /** Hitung kata per bagian. */
  function countDocument(infos, heads, profile) {
    const L = profile.limits;
    const res = { sections: [], total: 0, abstract: 0, limit: L.bodyWords, abstractLimit: L.abstractWords };
    for (let k = 0; k < heads.length; k++) {
      const h = heads[k];
      const end = heads[k + 1] ? heads[k + 1].index : infos.length;
      let w = 0;
      for (let j = h.index + 1; j < end; j++) {
        if (infos[j].inTable && !L.countTables) continue;
        w += countWords(infos[j].text);
      }
      const sec = profile.sections.find((s) => s.key === h.key);
      const isRef = sec && sec.kind === "references";
      const counted = !isRef || L.countReferences;
      res.sections.push({ key: h.key, title: h.title, words: w, counted });
      if (counted) res.total += w;
    }
    // abstrak: paragraf setelah "ABSTRACT" sampai ARTICLE HISTORY / KEYWORDS / judul bagian
    const ai = infos.findIndex((p) => /^(abstract|abstrak)$/i.test((p.text || "").trim()));
    if (ai >= 0) {
      for (let j = ai + 1; j < infos.length; j++) {
        const t = (infos[j].text || "").trim();
        if (/^(article history|keywords?|kata kunci|riwayat artikel)\b/i.test(t)) break;
        if (heads.some((h) => h.index === j)) break;
        if (!infos[j].inTable && infos[ai].inTable) break;
        res.abstract += countWords(t);
      }
    }
    return res;
  }

  /** Klasifikasi ulang paragraf yang sudah ada di dokumen (untuk "Rapikan"). */
  function classifyExisting(paras, sec, profile) {
    const blocks = paras.map((p) => ({ kind: "p", tag: "p", runs: [{ text: p.text || "", b: p.bold === true }] }));
    const det = Object.assign({}, profile.detect, { mergeBrokenLines: false });
    const prof = Object.assign({}, profile, { detect: det, references: Object.assign({}, profile.references, { sort: false }) });
    if (sec.kind === "references") return blocks.map((b) => ({ type: plain(b.runs).trim() ? "ref" : "skip" }));
    const items = [];
    blocks.forEach((b, i) => {
      if (!plain(b.runs).trim()) { items[i] = { type: "skip" }; return; }
      const one = classify([b].concat(blocks[i + 1] ? [blocks[i + 1]] : []), {}, prof)[0] || { type: "body" };
      items[i] = one;
    });
    // level heading untuk R&D
    const majors = new Set((sec.majorSubheadings || []).map(normHead));
    let seen = false;
    items.forEach((it, i) => {
      if (it.type !== "heading") return;
      const t = normHead(paras[i].text);
      if (majors.has(t)) { it.level = 2; seen = true; } else it.level = seen ? 3 : 2;
    });
    return items;
  }

  return {
    DEFAULT_PROFILE, mergeProfile, clone,
    parseInput, parseHtml, parseText, classify, classifyReferences, classifyExisting,
    paraFormat, finalRuns, countWords, plain, normalizeRuns, titleCase, apaAutoItalic, linkUrls,
    findSectionHeadings, sectionRange, pickAnchor, planNewSection, countDocument, normHead,
    CM,
  };
});
