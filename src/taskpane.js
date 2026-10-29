/* UI task pane "Template Paste" */
(function () {
  "use strict";
  const E = window.TplEngine;
  const STORE = "tplpaste.v1";
  const $ = (s, el) => (el || document).querySelector(s);
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const fmtN = (n) => Number(n || 0).toLocaleString("id-ID");

  const state = {
    inWord: false,
    store: { active: null, profiles: {} },
    profile: null,
    mode: "section",
    pending: {}, // key -> {blocks:[], overrides:{}, items:[], images:0}
    backup: {},  // key -> OOXML isi yang dihapus (hanya selama panel terbuka)
    found: {},
    open: {},
  };

  /* ---------------- penyimpanan profil ---------------- */
  function loadStore() {
    try {
      const raw = localStorage.getItem(STORE);
      if (raw) state.store = JSON.parse(raw);
    } catch (e) { /* localStorage tidak tersedia */ }
    const st = state.store;
    if (!st.profiles) st.profiles = {};
    // migrasi: nama lama JEPSD -> JEPSI
    Object.keys(st.profiles).forEach((n) => {
      const p = st.profiles[n];
      if (p.base === "jepsd") p.base = "jepsi";
      if (n.startsWith("JEPSD")) {
        const nn = n.replace(/^JEPSD/, "JEPSI");
        p.name = nn; st.profiles[nn] = p; delete st.profiles[n];
        if (st.active === n) st.active = nn;
      }
    });
    if (st.presetsAdded) st.presetsAdded = st.presetsAdded.map((b) => (b === "jepsd" ? "jepsi" : b));
    // tandai profil lama yang berasal dari template bawaan
    Object.values(st.profiles).forEach((p) => { if (!p.base) { const pr = E.PRESETS.find((x) => x.name === p.name); p.base = pr ? pr.base : "aieds"; } });
    // tambahkan template bawaan yang belum pernah ditambahkan (yang sudah dihapus pengguna tidak dimunculkan lagi)
    st.presetsAdded = st.presetsAdded || Object.values(st.profiles).map((p) => p.base);
    E.PRESETS.forEach((pr) => {
      if (st.presetsAdded.includes(pr.base) && Object.keys(st.profiles).length) return;
      let n = pr.name, k = 2;
      while (st.profiles[n]) n = pr.name + " " + k++;
      st.profiles[n] = Object.assign(E.clone(pr), { name: n });
      if (!st.presetsAdded.includes(pr.base)) st.presetsAdded.push(pr.base);
    });
    // profil lama: tambahkan kartu ABSTRACT bila belum ada
    Object.values(st.profiles).forEach((p) => {
      if (Array.isArray(p.sections) && !p.sections.some((x) => x.kind === "abstract")) {
        p.sections.unshift(E.clone(E.presetFor(p.base).sections.find((x) => x.kind === "abstract")));
      }
    });
    if (!st.active || !st.profiles[st.active]) st.active = Object.keys(st.profiles)[0];
    state.profile = E.mergeProfile(st.profiles[st.active]);
    state.mode = state.profile.layout.insertMode || "section";
  }
  let saveTimer = null;
  function saveStore() {
    state.store.profiles[state.store.active] = state.profile;
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => { try { localStorage.setItem(STORE, JSON.stringify(state.store)); } catch (e) { /* abaikan */ } }, 250);
    $("#profileLabel").textContent = state.store.active;
  }

  /* ---------------- toast ---------------- */
  let toastTimer;
  function toast(msg, kind, ms) {
    const t = $("#toast");
    t.textContent = msg;
    t.className = "toast show " + (kind || "");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (t.className = "toast"), ms || (kind === "err" ? 7000 : 3500));
  }
  function errMsg(e) {
    let m = (e && e.message) || String(e);
    if (e && e.debugInfo && e.debugInfo.errorLocation) m += " (" + e.debugInfo.errorLocation + ")";
    return m;
  }
  function needWord() {
    if (!state.inWord) { toast("Buka add-in ini dari Microsoft Word untuk menyisipkan ke dokumen.", "err"); return false; }
    return true;
  }

  /* ---------------- tab ---------------- */
  function initTabs() {
    document.querySelectorAll(".tab").forEach((b) => b.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("active", x === b));
      ["paste", "words", "template"].forEach((k) => $("#tab-" + k).classList.toggle("hidden", k !== b.dataset.tab));
      if (b.dataset.tab === "words") recount();
    }));
  }

  /* ---------------- mode sisip ---------------- */
  function initMode() {
    const hints = {
      section: "Isi ditambahkan di bawah konten yang sudah ada pada bagian tersebut. Judul bagian dibuat otomatis bila belum ada.",
      cursor: "Isi disisipkan di posisi kursor. Bila kursor di baris kosong, isi masuk tepat di baris itu.",
    };
    const sync = () => {
      document.querySelectorAll("#modeSeg button").forEach((b) => b.classList.toggle("on", b.dataset.mode === state.mode));
      $("#modeHint").textContent = hints[state.mode];
    };
    document.querySelectorAll("#modeSeg button").forEach((b) => b.addEventListener("click", () => {
      state.mode = b.dataset.mode; state.profile.layout.insertMode = state.mode; saveStore(); sync();
    }));
    sync();
  }

  /* ---------------- kartu bagian ---------------- */
  const TYPE_OPTS = [
    ["body", "Isi"], ["heading2", "Heading 2"], ["heading3", "Heading 3"],
    ["quote", "Kutipan"], ["caption", "Caption"], ["note", "Catatan"], ["ref", "Referensi"], ["table", "Tabel"], ["skip", "Lewati"],
  ];
  const typeKey = (it) => (it.type === "heading" ? "heading" + (it.level || 2) : it.type);

  function pend(key) {
    return state.pending[key] || (state.pending[key] = { blocks: [], overrides: {}, items: [], images: 0 });
  }
  function sec(key) { return state.profile.sections.find((s) => s.key === key); }

const isAbs = (key) => { const x = sec(key); return !!x && x.kind === "abstract"; };
  function reclassify(key) {
    const p = pend(key);
    p.items = isAbs(key) ? E.classifyAbstract(p.blocks, sec(key), state.profile) : E.classify(p.blocks, sec(key), state.profile);
    Object.entries(p.overrides).forEach(([i, v]) => {
      const it = p.items[i];
      if (!it) return;
      if (v.startsWith("heading")) { it.type = "heading"; it.level = +v.slice(7); }
      else it.type = v;
    });
  }

  function runsHtml(runs, max) {
    let out = "", n = 0;
    for (const r of runs || []) {
      if (n >= max) break;
      let t = r.text.slice(0, max - n); n += t.length;
      t = esc(t);
      if (r.i) t = "<i>" + t + "</i>";
      if (r.b) t = "<b>" + t + "</b>";
      out += t;
    }
    return out + (n >= max ? "…" : "");
  }

  function itemWords(items) {
    return items.filter((i) => i.type !== "skip").reduce((a, i) => a + (i.type === "table" ? 0 : E.countWords((i.label ? i.label + " " : "") + E.plain(i.runs))), 0);
  }

  function renderSections() {
    const host = $("#sections");
    host.innerHTML = "";
    for (const s of state.profile.sections) {
      const p = pend(s.key);
      const live = p.items.filter((i) => i.type !== "skip");
      const card = document.createElement("section");
      card.className = "sec" + (state.open[s.key] ? " open" : "");
      card.dataset.key = s.key;
      const abs = s.kind === "abstract";
      const docPill = state.inWord ? (state.found[s.key] ? '<span class="pill indoc" title="Judul bagian ditemukan di dokumen">ada</span>'
        : abs ? '<span class="pill new" title="Judul ABSTRACT tidak ditemukan; gunakan mode Posisi kursor">tidak ada</span>' : '<span class="pill new" title="Judul bagian akan dibuat otomatis">baru</span>') : "";
      const words = itemWords(p.items), absLim = +state.profile.limits.abstractWords || 0;
      const readyPill = !live.length ? "" : abs
        ? `<span class="pill ready${absLim && words > absLim ? " over" : ""}">${fmtN(words)}${absLim ? " / " + fmtN(absLim) : ""} kata</span>`
        : `<span class="pill ready">${live.length} siap · ${fmtN(words)} kata</span>`;
      card.innerHTML = `
        <div class="sec-head">
          <svg class="chev" viewBox="0 0 16 16"><path d="M6 3l5 5-5 5" fill="none" stroke="currentColor" stroke-width="1.8"/></svg>
          <div class="sec-title">${esc(s.title)}</div>
          ${readyPill}
          ${docPill}
        </div>
        <div class="sec-body">
          <div class="drop" contenteditable="true" spellcheck="false" data-key="${esc(s.key)}">Klik di sini lalu tekan Ctrl+V (⌘V) untuk menempel ${abs ? "abstrak (menggantikan abstrak yang ada di dokumen)" : "isi " + esc(titleCase(s.title))}</div>
          ${p.images ? `<p class="warnline">${p.images} gambar di tempelan diabaikan — sisipkan gambar langsung di Word.</p>` : ""}
          <div class="preview ${p.items.length ? "" : "hidden"}"></div>
          <div class="sec-meta ${p.items.length ? "" : "hidden"}">${summary(p.items)}</div>
          <div class="sec-actions">
            <button class="btn primary small act-insert" ${live.length ? "" : "disabled"}>${abs ? "Ganti abstrak di dokumen" : "Sisipkan ke dokumen"}</button>
            ${abs ? "" : '<button class="btn small act-reformat" title="Terapkan format template ke isi yang sudah ada di bagian ini">Rapikan isi yang ada</button>'}
            <button class="btn small ghost act-clear" ${p.items.length ? "" : "disabled"} title="Kosongkan tempelan di panel ini (dokumen tidak berubah)">Kosongkan tempelan</button>
          </div>
          <div class="sec-actions danger-zone">
            <button class="btn small danger act-wipe" ${state.inWord && !state.found[s.key] ? "disabled" : ""} title="Hapus semua isi bagian ini di dokumen; judul bagian tetap ada">${abs ? "Kosongkan abstrak di dokumen" : "Hapus isi bagian ini di dokumen"}</button>
            ${state.backup[s.key] ? `<button class="btn small act-restore" title="Kembalikan isi yang terakhir dihapus">↶ Kembalikan isi yang dihapus</button>` : ""}
          </div>
        </div>`;
      host.appendChild(card);
      renderPreview(card, s.key);

      card.querySelector(".sec-head").addEventListener("click", () => { state.open[s.key] = !state.open[s.key]; card.classList.toggle("open"); });
      const drop = card.querySelector(".drop");
      drop.addEventListener("paste", (ev) => onPaste(ev, s.key, drop));
      drop.addEventListener("keydown", (ev) => { if (!(ev.ctrlKey || ev.metaKey) && ev.key.length === 1) ev.preventDefault(); });
      drop.addEventListener("drop", (ev) => ev.preventDefault());
      card.querySelector(".act-insert").addEventListener("click", () => doInsert([s.key]));
      if (!abs) card.querySelector(".act-reformat").addEventListener("click", () => doReformat(s.key));
      card.querySelector(".act-clear").addEventListener("click", () => { state.pending[s.key] = null; renderSections(); });
      card.querySelector(".act-wipe").addEventListener("click", () => doWipe([s.key]));
      const rb = card.querySelector(".act-restore");
      if (rb) rb.addEventListener("click", () => doRestore(s.key));
    }
  }

  function summary(items) {
    if (items.some((i) => i.type === "abstract")) {
      const lab = items.filter((i) => i.type === "abstract" && i.label).length;
      const sk = items.filter((i) => i.type === "skip" && i.why === "keywords").length;
      return (lab ? `Abstrak terstruktur: ${lab} label dibuat bold.` : "Abstrak satu bagian.") + (sk ? " Baris Keywords tidak dimasukkan." : "") + " Abstrak lama di dokumen akan diganti (bisa dikembalikan).";
    }
    const c = (t, l) => items.filter((i) => i.type === t && (l == null || i.level === l)).length;
    const parts = [];
    if (c("body")) parts.push(c("body") + " paragraf");
    if (c("heading")) parts.push(c("heading") + " sub-heading");
    if (c("ref")) parts.push(c("ref") + " referensi");
    if (c("caption")) parts.push(c("caption") + " caption");
    if (c("table")) parts.push(c("table") + " tabel");
    if (c("note")) parts.push(c("note") + " catatan");
    if (c("quote")) parts.push(c("quote") + " kutipan");
    return parts.join(" · ") + (parts.length ? ". Ubah jenis lewat menu di kiri bila deteksi keliru." : "");
  }

  function renderPreview(card, key) {
    const p = pend(key);
    const box = card.querySelector(".preview");
    box.innerHTML = "";
    p.items.forEach((it, i) => {
      const row = document.createElement("div");
      row.className = "prow t-" + it.type + (it.level === 3 ? " l3" : "");
      const sel = document.createElement("select");
      const opts = isAbs(key) ? [["abstract", "Abstrak"], ["skip", "Lewati"]] : TYPE_OPTS;
      opts.forEach(([v, l]) => {
        if (v === "table" && it.type !== "table") return;
        if (it.type === "table" && !["table", "skip"].includes(v)) return;
        const o = document.createElement("option"); o.value = v; o.textContent = l; sel.appendChild(o);
      });
      sel.value = typeKey(it);
      sel.addEventListener("change", () => {
        p.overrides[i] = sel.value; reclassify(key);
        const open = state.open[key]; renderSections(); state.open[key] = open;
      });
      const txt = document.createElement("div");
      txt.className = "txt";
      txt.innerHTML = it.type === "abstract" && it.label ? "<b>" + esc(it.label) + " –</b> " + runsHtml(it.runs, 160) : it.type === "table" ? esc(`Tabel ${it.rows.length} baris × ${Math.max(...it.rows.map((r) => r.length))} kolom: ${it.rows[0].join(" | ")}`).slice(0, 200) : runsHtml(it.runs, 180);
      row.append(sel, txt);
      box.appendChild(row);
    });
  }

  function titleCase(s) { return E.titleCase(String(s).toLowerCase()); }

  function onPaste(ev, key, drop) {
    ev.preventDefault();
    const cd = ev.clipboardData || window.clipboardData;
    const clip = { html: cd.getData("text/html") || "", text: cd.getData("text/plain") || cd.getData("Text") || "" };
    const parsed = E.parseInput(clip, sec(key), state.profile);
    if (!parsed.blocks.length) { toast("Tidak ada teks yang terbaca dari clipboard.", "err"); return; }
    const p = pend(key);
    p.blocks = p.blocks.concat(parsed.blocks);
    p.images += parsed.images || 0;
    reclassify(key);
    state.open[key] = true;
    renderSections();
    const d = document.querySelector(`.drop[data-key="${CSS.escape(key)}"]`);
    if (d) { d.classList.add("flash"); d.textContent = `✓ ${parsed.blocks.length} blok ditambahkan — tempel lagi untuk menambah`; setTimeout(() => d.classList.remove("flash"), 1500); }
  }

  /* ---------------- aksi Word ---------------- */
  async function doInsert(keys) {
    if (!needWord()) return;
    const todo = keys.filter((k) => pend(k).items.some((i) => i.type !== "skip"));
    if (!todo.length) { toast("Belum ada isi yang ditempel."); return; }
    if (state.mode === "cursor" && todo.length > 1) { toast("Mode kursor hanya untuk satu bagian sekaligus. Pilih 'Akhir bagian'.", "err"); return; }
    setBusy(true);
    try {
      let last = null, total = 0;
      for (const k of todo) {
        const items = pend(k).items;
        if (isAbs(k)) {
          const r = await WordOps.insertAbstract(items, state.profile, state.mode);
          if (r.backup) state.backup[k] = r.backup;
          total += r.inserted;
          const lim = +state.profile.limits.abstractWords || 0;
          if (lim && r.words > lim) toast(`Abstrak ${fmtN(r.words)} kata — melebihi batas ${fmtN(lim)}.`, "err");
          state.pending[k] = null;
          if (todo.length === 1) { await refreshScan(); renderSections(); toast(`✓ Abstrak diganti (${fmtN(r.words)} kata${lim ? " / " + fmtN(lim) : ""}). Abstrak lama bisa dikembalikan.`, lim && r.words > lim ? "err" : "ok", 6000); return; }
          continue;
        }
        last = await WordOps.insertSection(k, items, state.profile, state.mode);
        total += last.inserted;
        if (last.missingStyles && last.missingStyles.length) toast("Gaya tidak ditemukan, dipakai Normal: " + last.missingStyles.join(", "), "err");
        state.pending[k] = null;
      }
      await refreshScan();
      renderSections();
      const c = last.count;
      const over = c.limit && c.total > c.limit;
      toast(`✓ ${total} blok disisipkan. Total isi ${fmtN(c.total)}${c.limit ? " / " + fmtN(c.limit) : ""} kata${over ? " — MELEBIHI BATAS" : ""}.`, over ? "err" : "ok", 6000);
      if (!$("#tab-words").classList.contains("hidden")) renderCounts(c);
    } catch (e) {
      console.error(e);
      toast("Gagal menyisipkan: " + errMsg(e), "err");
    } finally { setBusy(false); }
  }

  /** Hapus isi bagian di dokumen (judul tetap). */
  async function doWipe(keys) {
    if (!needWord()) return;
    await refreshScan();
    const targets = keys.filter((k) => state.found[k]);
    if (!targets.length) { toast("Judul bagian tidak ditemukan di dokumen."); return; }
    const words = targets.reduce((a, k) => {
      const r = state.lastCount && state.lastCount.sections.find((x) => x.key === k);
      return a + (r ? r.words : 0);
    }, 0);
    const names = targets.map((k) => sec(k).title).join(", ");
    const msg = targets.length === 1 && isAbs(targets[0])
      ? "Kosongkan teks abstrak di dokumen? Judul ABSTRACT dan isi kotak lain (keywords, article history) tidak berubah."
      : targets.length === 1
      ? `Hapus SEMUA isi ${names} di dokumen (±${fmtN(words)} kata, termasuk sub-heading, tabel & gambar)? Judul bagian tetap ada.`
      : `Hapus SEMUA isi ${targets.length} bagian (${names}) — ±${fmtN(words)} kata, termasuk tabel & gambar? Judul-judul bagian tetap ada.`;
    if (!(await ask(msg, null, { ok: "Hapus", danger: true }))) return;
    setBusy(true);
    try {
      let total = 0;
      // dari bagian paling bawah ke atas
      const order = state.profile.sections.map((x) => x.key).filter((k) => targets.includes(k)).reverse();
      for (const k of order) {
        const r = isAbs(k) ? await WordOps.clearAbstract(state.profile) : await WordOps.clearSection(k, state.profile);
        if (r.ooxml) state.backup[k] = r.ooxml;
        total += r.deleted;
      }
      await refreshScan();
      renderSections();
      toast(total ? `✓ Isi dihapus (${total} paragraf). Bisa dipulihkan dengan "Kembalikan" atau Ctrl+Z.` : "Bagian sudah kosong.", total ? "ok" : "", 6000);
    } catch (e) { toast("Gagal menghapus: " + errMsg(e), "err"); } finally { setBusy(false); }
  }

  async function doRestore(key) {
    if (!needWord() || !state.backup[key]) return;
    setBusy(true);
    try {
      if (isAbs(key)) await WordOps.restoreAbstract(state.backup[key], state.profile);
      else await WordOps.restoreSection(key, state.backup[key], state.profile);
      delete state.backup[key];
      await refreshScan(); renderSections();
      toast("✓ Isi " + sec(key).title + " dikembalikan.", "ok");
    } catch (e) { toast("Gagal mengembalikan: " + errMsg(e) + ". Coba Ctrl+Z.", "err"); } finally { setBusy(false); }
  }

  async function doReformat(key) {
    if (!needWord()) return;
    setBusy(true);
    try {
      const r = await WordOps.reformatSection(key, state.profile);
      toast(`✓ ${r.formatted} paragraf dirapikan (${r.headings} sub-heading).`, "ok");
    } catch (e) { toast(errMsg(e), "err"); } finally { setBusy(false); }
  }

  async function doUndo() {
    if (!needWord()) return;
    try {
      const ok = await WordOps.undoLast();
      toast(ok ? "✓ Sisipan terakhir dihapus." : "Tidak ada sisipan yang bisa dihapus.", ok ? "ok" : "");
      await refreshScan(); renderSections();
    } catch (e) { toast(errMsg(e), "err"); }
  }

  function setBusy(b) {
    document.querySelectorAll("button").forEach((x) => {
      if (b) { x.dataset.wasDisabled = x.disabled ? "1" : ""; x.disabled = true; }
      else if (x.dataset.wasDisabled !== undefined) { x.disabled = x.dataset.wasDisabled === "1"; delete x.dataset.wasDisabled; }
    });
    document.body.style.cursor = b ? "progress" : "";
  }

  async function refreshScan() {
    if (!state.inWord) return;
    try { const r = await WordOps.scan(state.profile); state.found = r.found; state.lastCount = r.count; } catch (e) { console.warn(e); }
  }

  /* ---------------- hitung kata ---------------- */
  async function recount() {
    renderLimits();
    if (!state.inWord) { renderCounts(null); return; }
    try { state.lastCount = await WordOps.countWords(state.profile); renderCounts(state.lastCount); }
    catch (e) { toast(errMsg(e), "err"); }
  }

  /** Aturan hitung kata — tersimpan per profil jurnal. */
  function renderLimits() {
    const L = state.profile.limits;
    L.excludeKeys = L.excludeKeys || [];
    const host = $("#limitSettings");
    const unset = !L.bodyWords && !L.abstractWords && !L.minBodyWords;
    const num = (k, label, step) => `<label for="lim_${k}">${label}</label><input type="number" min="0" step="${step}" id="lim_${k}" data-k="${k}" value="${+L[k] || 0}">`;
    const chk = (id, label, on, attrs) => `<label for="${id}">${label}</label><input type="checkbox" id="${id}" ${attrs} ${on ? "checked" : ""}>`;
    host.innerHTML = `<fieldset class="limits"><details ${unset || state.limitsOpen ? "open" : ""}>
      <summary>Aturan hitung kata · ${esc(state.store.active.split(" (")[0])}</summary>
      <div class="fgrid">
        ${num("bodyWords", "Maksimal kata isi", 100)}
        ${num("minBodyWords", "Minimal kata isi", 100)}
        ${num("abstractWords", "Maksimal kata abstrak", 10)}
        <div class="fdesc">0 = tidak dibatasi. Isi sesuai ketentuan jurnal; tersimpan untuk profil ini saja.</div>
        ${chk("lim_tables", "Teks di dalam tabel ikut dihitung", L.countTables, 'data-b="countTables"')}
        <div class="subhead">Bagian yang ikut dihitung</div>
        ${state.profile.sections.filter((s) => s.kind !== "abstract").map((s) => {
          const on = s.kind === "references" ? L.countReferences : !L.excludeKeys.includes(s.key);
          return chk("lim_s_" + s.key, esc(s.title), on, `data-sec="${esc(s.key)}" data-ref="${s.kind === "references" ? 1 : ""}"`);
        }).join("")}
      </div></details></fieldset>
      ${unset ? '<p class="warnline">Batas kata untuk jurnal ini belum diatur.</p>' : ""}`;
    host.querySelector("details").addEventListener("toggle", (e) => { state.limitsOpen = e.target.open; });
    host.querySelectorAll("input").forEach((inp) => inp.addEventListener("change", () => {
      if (inp.dataset.k) L[inp.dataset.k] = Math.max(0, parseInt(inp.value, 10) || 0);
      else if (inp.dataset.b) L[inp.dataset.b] = inp.checked;
      else if (inp.dataset.sec) {
        if (inp.dataset.ref) L.countReferences = inp.checked;
        else {
          L.excludeKeys = L.excludeKeys.filter((k) => k !== inp.dataset.sec);
          if (!inp.checked) L.excludeKeys.push(inp.dataset.sec);
        }
      }
      state.limitsOpen = true;
      saveStore();
      recount();
    }));
  }

  function renderCounts(c) {
    const T = $("#wordTotals"), W = $("#wordTable");
    if (!c) { T.innerHTML = '<p class="hint">Buka di Word untuk menghitung kata dokumen.</p>'; W.innerHTML = ""; return; }
    const card = (label, n, lim, min) => {
      const over = lim && n > lim, under = min && n < min;
      const pct = lim ? Math.min(100, (n / lim) * 100) : 0;
      let info = lim ? "/ " + fmtN(lim) + " kata" : "kata";
      if (over) info += " · lebih " + fmtN(n - lim);
      else if (lim) info += " · sisa " + fmtN(lim - n);
      if (under) info += " · kurang " + fmtN(min - n) + " dari minimal " + fmtN(min);
      return `<div class="total-card ${over ? "over" : under ? "under" : ""}"><div class="of">${label}</div>
        <div><span class="big">${fmtN(n)}</span> <span class="of">${info}</span></div>
        ${lim ? `<div class="bar ${over ? "over" : ""}"><span style="width:${pct}%"></span></div>` : ""}</div>`;
    };
    T.innerHTML = card("Isi artikel (bagian yang dicentang)", c.total, c.limit, c.minLimit) + card("Abstrak", c.abstract, c.abstractLimit, 0);
    W.innerHTML = `<table class="wtable">${c.sections.map((s) => `<tr class="${s.counted ? "" : "nc"}"><td>${esc(s.title)}</td><td class="n">${fmtN(s.words)}</td></tr>`).join("")}</table>
      ${c.sections.length ? "" : '<p class="hint">Belum ada judul bagian yang dikenali di dokumen.</p>'}`;
  }

  /* ---------------- editor template ---------------- */
  const ALIGN = [["Left", "Kiri"], ["Centered", "Tengah"], ["Right", "Kanan"], ["Justified", "Justify"]];
  const paraFields = (opt) => [
    ["size", "Ukuran font (pt)", "number", { step: 0.5, min: 6 }],
    ...(opt.bold ? [["bold", "Tebal (bold)", "bool"], ["italic", "Miring (italic)", "bool"]] : []),
    ...(opt.caps ? [["caps", "HURUF KAPITAL", "bool"]] : []),
    ["align", "Perataan", "select", { options: ALIGN }],
    ...(opt.indent ? [["firstIndentCm", "Indentasi baris pertama (cm)", "number", { step: 0.01 }]] : []),
    ...(opt.hanging ? [["hangingCm", "Indentasi gantung / hanging (cm)", "number", { step: 0.01 }]] : []),
    ["before", "Spasi sebelum (pt)", "number", { step: 1, min: 0 }],
    ["after", "Spasi sesudah (pt)", "number", { step: 1, min: 0 }],
    ["line", "Spasi baris (1 = tunggal)", "number", { step: 0.05, min: 0.5 }],
    ["style", "Gaya Word (style)", "text", { desc: "Biarkan Normal kecuali jurnal memakai gaya khusus; tulis persis seperti di Word." }],
  ];
  const GROUPS = [
    { title: "Umum", fields: [
      ["font", "Font", "text"], ["color", "Warna teks", "color"],
      ["layout.blankBeforeSection", "Sisakan baris kosong sebelum judul bagian berikutnya", "bool"],
    ] },
    { title: "Judul bagian (INTRODUCTION, METHOD, …)", prefix: "sectionHeading", fields: paraFields({ bold: true, caps: true }) },
    { title: "Sub-heading tingkat 2", prefix: "h2", fields: paraFields({ bold: true }) },
    { title: "Sub-heading tingkat 3", prefix: "h3", fields: paraFields({ bold: true }) },
    { title: "Paragraf isi", prefix: "body", fields: paraFields({ indent: true }) },
    { title: "Caption tabel & gambar", prefix: "caption", fields: [["labelBold", "Label \"Table 1.\" tebal", "bool"], ...paraFields({})] },
    { title: "Catatan tabel (Note.)", prefix: "note", fields: [["labelItalic", "Label \"Note.\" miring", "bool"], ...paraFields({})] },
    { title: "Abstrak", prefix: "abstract", fields: [
      ["size", "Ukuran font (pt)", "number", { step: 0.5, min: 6 }],
      ["italic", "Miring (italic)", "bool"],
      ["align", "Perataan", "select", { options: ALIGN }],
      ["before", "Spasi sebelum (pt)", "number", { step: 1, min: 0 }],
      ["after", "Spasi sesudah (pt)", "number", { step: 1, min: 0 }],
      ["line", "Spasi baris (1 = tunggal)", "number", { step: 0.01, min: 0.5 }],
      ["labelBold", "Label abstrak terstruktur bold (Purpose – …)", "bool"],
      ["labelSep", "Pemisah label (kosong = ikuti tempelan)", "text", { desc: "Contoh: \" – \" (AIEDS) atau \": \"" }],
      ["joinLabeled", "Komponen terstruktur", "select", { options: [["linebreak", "1 paragraf, dipisah line break"], ["paragraphs", "Paragraf terpisah"]] }],
      ["labels", "Label yang dikenali (pisahkan koma)", "list"],
    ] },
    { title: "Kutipan langsung (wawancara)", prefix: "quote", fields: [
      ["italic", "Miring (italic)", "bool"],
      ["leftCm", "Indentasi kiri (cm)", "number", { step: 0.01, min: 0 }],
      ["rightCm", "Indentasi kanan (cm)", "number", { step: 0.01, min: 0 }],
      ...paraFields({}),
    ] },
    { title: "Tabel yang ditempel", prefix: "table", fields: [
      ["cellAlign", "Perataan isi sel", "select", { options: ALIGN }],
      ["size", "Ukuran font tabel (pt)", "number", { step: 0.5 }], ["headerBold", "Baris judul tebal", "bool"],
      ["apaBorders", "Garis gaya APA (atas, bawah, bawah judul)", "bool"], ["firstColLeft", "Kolom pertama rata kiri", "bool"],
    ] },
    { title: "Referensi", prefix: "references", fields: [
      ...paraFields({ hanging: true }),
      ["autoItalic", "Italic otomatis APA (nama jurnal, volume / judul buku) bila tempelan tanpa italic", "bool"],
      ["linkDoi", "Jadikan DOI/URL hyperlink", "bool"],
      ["stripNumbering", "Hapus penomoran [1] / 1.", "bool"],
      ["sort", "Urutkan A–Z (yang ditempel)", "bool"],
    ] },
    { title: "Deteksi otomatis", prefix: "detect", fields: [
      ["maxHeadingWords", "Maks. kata untuk sub-heading", "number", { step: 1, min: 3 }],
      ["stripHeadingNumbers", "Hapus nomor heading (2.1, A.)", "bool"],
      ["titleCaseHeadings", "Ubah sub-heading ke Title Case", "bool"],
      ["keepItalic", "Pertahankan italic dari tempelan", "bool"],
      ["keepBold", "Pertahankan bold di dalam paragraf", "bool"],
      ["keepSupSub", "Pertahankan superscript/subscript", "bool"],
      ["mergeBrokenLines", "Sambung baris terputus (salinan PDF)", "bool"],
      ["detectQuotes", "Kenali kutipan langsung (diawali tanda kutip “…”)", "bool"],
      ["requireCapsSectionTitle", "Judul bagian di dokumen ditulis KAPITAL", "bool", { desc: "Membedakan judul bagian RESULTS dari sub-heading Results." }],
    ] },
  ];

  const getPath = (o, p) => p.split(".").reduce((a, k) => (a == null ? a : a[k]), o);
  const setPath = (o, p, v) => { const ks = p.split("."); let a = o; ks.slice(0, -1).forEach((k) => (a = a[k] = a[k] || {})); a[ks[ks.length - 1]] = v; };

  function renderTemplate() {
    const sel = $("#profileSelect");
    const opts = Object.keys(state.store.profiles).map((n) => `<option ${n === state.store.active ? "selected" : ""}>${esc(n)}</option>`).join("");
    sel.innerHTML = opts;
    $("#journalSelect").innerHTML = opts;
    const form = $("#templateForm");
    form.innerHTML = "";
    GROUPS.forEach((g, gi) => {
      const fs = document.createElement("fieldset");
      const det = document.createElement("details");
      if (gi === 0) det.open = true;
      det.innerHTML = `<summary>${esc(g.title)}</summary>`;
      const grid = document.createElement("div");
      grid.className = "fgrid";
      g.fields.forEach(([k, label, type, opt]) => {
        opt = opt || {};
        const path = g.prefix ? g.prefix + "." + k : k;
        const id = "f_" + path.replace(/\./g, "_");
        const val = getPath(state.profile, path);
        const lab = document.createElement("label"); lab.htmlFor = id; lab.textContent = label;
        let inp;
        if (type === "select") {
          inp = document.createElement("select");
          opt.options.forEach(([v, l]) => { const o = document.createElement("option"); o.value = v; o.textContent = l; inp.appendChild(o); });
          inp.value = val;
        } else {
          inp = document.createElement("input");
          inp.type = type === "bool" ? "checkbox" : type;
          if (type === "list") inp.type = "text";
          if (type === "bool") inp.checked = !!val; else inp.value = val == null ? "" : Array.isArray(val) ? val.join(", ") : val;
          if (opt.step) inp.step = opt.step;
          if (opt.min != null) inp.min = opt.min;
        }
        inp.id = id;
        inp.addEventListener("change", () => {
          let v = type === "bool" ? inp.checked : type === "number" ? parseFloat(inp.value) : type === "list" ? inp.value.split(",").map((x) => x.trim()).filter(Boolean) : inp.value;
          if (type === "number" && isNaN(v)) v = 0;
          setPath(state.profile, path, v);
          saveStore();
          Object.keys(state.pending).forEach((key) => state.pending[key] && reclassify(key));
          renderSections();
        });
        grid.append(lab, inp);
        if (opt.desc) { const d = document.createElement("div"); d.className = "fdesc"; d.textContent = opt.desc; grid.appendChild(d); }
      });
      det.appendChild(grid);
      fs.appendChild(det);
      form.appendChild(fs);
    });
    form.appendChild(renderSectionEditor());
  }

  function renderSectionEditor() {
    const fs = document.createElement("fieldset");
    const det = document.createElement("details");
    det.innerHTML = "<summary>Daftar bagian artikel</summary>";
    const box = document.createElement("div");
    box.className = "secedit";
    const list = state.profile.sections;
    list.forEach((s, i) => {
      const el = document.createElement("div");
      el.className = "secitem";
      el.innerHTML = `
        <span class="lab">Judul di dokumen</span><input type="text" data-f="title" value="${esc(s.title)}">
        <span class="lab">Nama lain (pisahkan koma)</span><input type="text" data-f="aliases" value="${esc((s.aliases || []).join(", "))}">
        <span class="lab">Sub-heading tingkat 2 khusus (mis. Results, Discussion)</span><input type="text" data-f="majors" value="${esc((s.majorSubheadings || []).join(", "))}">
        <div class="tools">${s.kind === "abstract" ? '<span class="lab">Abstrak (mengganti isi kotak abstrak)</span>' : `<label><input type="checkbox" data-f="refs" ${s.kind === "references" ? "checked" : ""}> Bagian referensi</label>`}
          <span><button type="button" class="icon-btn" data-a="up" title="Naik">↑</button> <button type="button" class="icon-btn" data-a="down" title="Turun">↓</button> <button type="button" class="icon-btn" data-a="del" title="Hapus">✕</button></span></div>`;
      el.querySelectorAll("input").forEach((inp) => inp.addEventListener("change", () => {
        const f = inp.dataset.f;
        if (f === "title") s.title = inp.value.trim() || s.title;
        if (f === "aliases") s.aliases = inp.value.split(",").map((x) => x.trim()).filter(Boolean);
        if (f === "majors") s.majorSubheadings = inp.value.split(",").map((x) => x.trim()).filter(Boolean);
        if (f === "refs") { if (inp.checked) s.kind = "references"; else delete s.kind; }
        saveStore(); renderSections();
      }));
      el.querySelectorAll("button").forEach((b) => b.addEventListener("click", () => {
        const a = b.dataset.a;
        if (a === "del") { ask(`Hapus bagian ${s.title}?`).then((ok) => { if (!ok) return; list.splice(i, 1); saveStore(); renderTemplate(); renderSections(); }); return; }
        if (a === "up" && i > 0) [list[i - 1], list[i]] = [list[i], list[i - 1]];
        if (a === "down" && i < list.length - 1) [list[i + 1], list[i]] = [list[i], list[i + 1]];
        saveStore(); renderTemplate(); renderSections();
        $("#templateForm details:last-of-type").open = true;
      }));
      box.appendChild(el);
    });
    const add = document.createElement("button");
    add.type = "button"; add.className = "btn small"; add.textContent = "+ Tambah bagian";
    add.addEventListener("click", () => {
      const key = "s" + Date.now().toString(36);
      list.splice(Math.max(0, list.findIndex((x) => x.kind === "references")), 0, { key, title: "NEW SECTION", aliases: [] });
      saveStore(); renderTemplate(); renderSections();
      $("#templateForm details:last-of-type").open = true;
    });
    box.appendChild(add);
    det.appendChild(box); fs.appendChild(det);
    return fs;
  }

  // Dialog sendiri — window.confirm/prompt diblokir di Office
  function ask(title, defVal, opt) {
    opt = opt || {};
    return new Promise((resolve) => {
      const d = $("#askDialog"), inp = $("#askInput");
      $("#askTitle").textContent = title;
      $("#askOk").textContent = opt.ok || "OK";
      $("#askOk").classList.toggle("danger-solid", !!opt.danger);
      inp.classList.toggle("hidden", defVal == null);
      inp.value = defVal == null ? "" : defVal;
      const done = (v) => { d.close(); $("#askOk").onclick = $("#askCancel").onclick = null; resolve(v); };
      $("#askOk").onclick = () => done(defVal == null ? true : inp.value.trim());
      $("#askCancel").onclick = () => done(defVal == null ? false : null);
      d.showModal();
      if (defVal != null) inp.focus();
    });
  }

  function initProfiles() {
    const switchProfile = (name) => {
      state.store.active = name;
      state.profile = E.mergeProfile(state.store.profiles[name]);
      saveStore(); renderTemplate();
      // isi yang sudah ditempel tetap ada, diklasifikasi ulang dengan aturan jurnal baru
      Object.keys(state.pending).forEach((k) => { if (state.pending[k]) { if (sec(k)) reclassify(k); else state.pending[k] = null; } });
      renderSections(); renderLimits(); refreshScan().then(renderSections);
      toast("Template aktif: " + name.split(" (")[0]);
    };
    $("#profileSelect").addEventListener("change", (e) => switchProfile(e.target.value));
    $("#journalSelect").addEventListener("change", (e) => switchProfile(e.target.value));
    $("#profDup").addEventListener("click", () => {
      let n = state.store.active + " (salinan)", k = 2;
      while (state.store.profiles[n]) n = state.store.active + " (salinan " + k++ + ")";
      const p = E.clone(state.profile); p.name = n;
      state.store.profiles[n] = p; state.store.active = n; state.profile = p;
      saveStore(); renderTemplate(); toast("Profil diduplikat: " + n);
    });
    $("#profRename").addEventListener("click", async () => {
      const cur = state.store.active;
      const n = await ask("Nama profil baru:", cur);
      if (!n || n === cur) return;
      if (state.store.profiles[n]) { toast("Nama sudah dipakai.", "err"); return; }
      state.store.profiles[n] = state.profile; delete state.store.profiles[cur];
      state.profile.name = n; state.store.active = n; saveStore(); renderTemplate();
    });
    $("#profDelete").addEventListener("click", async () => {
      if (Object.keys(state.store.profiles).length < 2) { toast("Minimal harus ada satu profil."); return; }
      if (!(await ask("Hapus profil " + state.store.active + "?"))) return;
      delete state.store.profiles[state.store.active];
      state.store.active = Object.keys(state.store.profiles)[0];
      state.profile = E.mergeProfile(state.store.profiles[state.store.active]);
      saveStore(); renderTemplate(); renderSections();
    });
    $("#profReset").addEventListener("click", async () => {
      const pr = E.presetFor(state.profile.base);
      if (!(await ask("Kembalikan profil ini ke pengaturan bawaan " + pr.name.split(" (")[0] + "?"))) return;
      const keep = state.store.active;
      state.profile = E.clone(pr); state.profile.name = keep;
      saveStore(); renderTemplate(); renderSections(); renderLimits(); toast("Profil dikembalikan ke bawaan.");
    });
    const dlg = $("#jsonDialog");
    $("#profExport").addEventListener("click", () => {
      $("#jsonTitle").textContent = "Ekspor profil — salin & simpan teks ini";
      $("#jsonText").value = JSON.stringify(state.profile, null, 2);
      $("#jsonApply").classList.add("hidden");
      dlg.showModal();
    });
    $("#profImport").addEventListener("click", () => {
      $("#jsonTitle").textContent = "Impor profil — tempel JSON profil di bawah";
      $("#jsonText").value = "";
      $("#jsonApply").classList.remove("hidden");
      dlg.showModal();
    });
    $("#jsonCopy").addEventListener("click", () => { $("#jsonText").select(); try { document.execCommand("copy"); toast("Disalin."); } catch (e) { /* */ } });
    $("#jsonClose").addEventListener("click", () => dlg.close());
    $("#jsonApply").addEventListener("click", () => {
      try {
        const p = E.mergeProfile(JSON.parse($("#jsonText").value));
        let n = p.name || "Profil impor", k = 2;
        while (state.store.profiles[n]) n = (p.name || "Profil impor") + " " + k++;
        p.name = n;
        state.store.profiles[n] = p; state.store.active = n; state.profile = p;
        saveStore(); renderTemplate(); renderSections(); dlg.close(); toast("Profil diimpor: " + n, "ok");
      } catch (e) { toast("JSON tidak valid: " + e.message, "err"); }
    });
  }

  /* ---------------- start ---------------- */
  function boot(inWord) {
    state.inWord = inWord;
    loadStore();
    $("#profileLabel").textContent = state.store.active;
    $("#notInWord").classList.toggle("hidden", inWord);
    state.open[state.profile.sections[0] && state.profile.sections[0].key] = true;
    initTabs(); initMode(); initProfiles();
    renderTemplate(); renderSections(); renderLimits(); renderCounts(null);
    $("#insertAll").addEventListener("click", () => doInsert(state.profile.sections.map((s) => s.key)));
    $("#undoLast").addEventListener("click", doUndo);
    $("#wipeAll").addEventListener("click", () => doWipe(state.profile.sections.filter((x) => x.kind !== "abstract").map((x) => x.key)));
    $("#recount").addEventListener("click", recount);
    if (inWord) refreshScan().then(renderSections);
  }

  if (typeof Office !== "undefined" && Office.onReady) {
    let booted = false;
    Office.onReady((info) => { if (booted) return; booted = true; boot(info && info.host === Office.HostType.Word); });
    // di luar Office, onReady kadang tidak terpanggil
    setTimeout(() => { if (!booted) { booted = true; boot(false); } }, 2500);
  } else {
    document.addEventListener("DOMContentLoaded", () => boot(false));
  }

  // untuk pengujian
  window.__tplState = state;
})();
