// A lab's spreadsheet, docked under the worksheet so a student can work the
// data without Excel and copy the results into the form above. Formulas come
// from sheet-engine.js; this file is the grid, the editing, the chart and the
// change log. Everything stays in this browser, like the worksheet's answers.
(function () {
  "use strict";
  const cfg = window.WORKSHEET && window.WORKSHEET.sheet;
  const toggle = document.getElementById("sheet-toggle");
  if (!cfg || !toggle || !window.SheetEngine) return;
  const E = window.SheetEngine;
  const key = "sheet:" + window.WORKSHEET.slug + ":lab" + window.WORKSHEET.lab;

  // ---- the panel ---------------------------------------------------------
  const dock = document.createElement("section");
  dock.className = "sheet-dock";
  dock.hidden = true;
  dock.setAttribute("aria-label", "Υπολογιστικό φύλλο");
  dock.innerHTML = `
    <div class="sheet-bar">
      <strong class="sheet-title"></strong>
      <span class="sheet-actions">
        <button type="button" data-act="fill" title="Αντιγράφει τον τύπο του επιλεγμένου κελιού προς τα κάτω (Ctrl+D)">Συμπλήρωση προς τα κάτω</button>
        <button type="button" data-act="addcol">+ Στήλη</button>
        <button type="button" data-act="addrow">+ Γραμμή</button>
        <button type="button" data-act="delrow">Διαγραφή γραμμής</button>
        <button type="button" data-act="freeze" aria-pressed="false">Πάγωμα 1ης γραμμής</button>
        <button type="button" data-act="chart">Γράφημα</button>
        <button type="button" data-act="log">Ιστορικό αλλαγών</button>
        <button type="button" data-act="csv">Λήψη CSV</button>
        <a class="sheet-original" download>Αρχικό CSV</a>
        <button type="button" data-act="reset">Επαναφορά</button>
        <button type="button" data-act="help" aria-expanded="false">Βοήθεια</button>
        <button type="button" data-act="size" title="Μεγαλύτερο ή μικρότερο παράθυρο">⇕</button>
        <button type="button" data-act="close" aria-label="Κλείσιμο">✕</button>
      </span>
    </div>
    <div class="sheet-help" hidden>
      <p>Οι τύποι ξεκινούν με <code>=</code>, όπως στο Excel: <code>=AVERAGE(D2:D17)</code>,
      <code>=D2/(E2/100)^2</code>, <code>=COUNTIF(H2:H17,"&gt;30")</code>. Τα ορίσματα χωρίζονται με
      <code>,</code> ή <code>;</code>. Αν το πληκτρολόγιο είναι στα ελληνικά, ο τύπος διαβάζεται κανονικά.
      Όσο γράφετε τύπο, ένα κλικ σε κελί (ή σύρσιμο σε περιοχή) βάζει τη διεύθυνσή του.</p>
      <p>Πλήκτρα: βέλη για μετακίνηση, Enter ή F2 για επεξεργασία, Delete για καθάρισμα,
      Ctrl+C / Ctrl+V για αντιγραφή και επικόλληση (οι τύποι προσαρμόζονται), Ctrl+D για συμπλήρωση προς τα κάτω.
      Το <code>$</code> κρατά σταθερή μια στήλη ή γραμμή (<code>$D$2</code>).</p>
      <p>Συναρτήσεις: <span class="sheet-fns"></span>.</p>
    </div>
    <div class="sheet-formula">
      <span class="sheet-name" aria-label="Κελί"></span>
      <input type="text" class="sheet-input" aria-label="Περιεχόμενο κελιού" autocomplete="off" spellcheck="false">
    </div>
    <div class="sheet-grid" tabindex="0"><table></table></div>
    <div class="sheet-log" hidden><ol></ol></div>
    <div class="sheet-status" role="status" aria-live="polite"></div>`;
  document.body.appendChild(dock);

  const $ = sel => dock.querySelector(sel);
  const table = $("table");
  const grid = $(".sheet-grid");
  const bar = $(".sheet-input");
  const nameBox = $(".sheet-name");
  const statusLine = $(".sheet-status");
  const logBox = $(".sheet-log");
  $(".sheet-title").textContent = "Υπολογιστικό φύλλο · " + cfg.name;
  $(".sheet-fns").textContent = E.FUNCTION_NAMES.filter(n => n !== "STDEVP").join(", ");
  const original = $(".sheet-original");
  original.href = cfg.csv;
  original.setAttribute("download", cfg.name);

  // ---- state -------------------------------------------------------------
  let sheet = null;
  let pristine = null;
  let log = [];
  let frozen = false;
  let values = [];
  let sel = { r: 1, c: 0, r2: 1, c2: 0 };
  let editor = null;  // {r, c, input, before, pick: {start, end} | null}
  let dragging = false;
  let clip = null;    // {r, c, cells}: what Ctrl+C took, so a paste can move formulas

  function save() {
    try {
      localStorage.setItem(key, JSON.stringify({ version: cfg.version, cells: sheet.cells, log, frozen }));
    } catch (e) { /* storage refused: the sheet still works for this visit */ }
  }
  function restore() {
    try {
      const s = JSON.parse(localStorage.getItem(key) || "null");
      if (s && s.version === cfg.version && Array.isArray(s.cells)) {
        sheet = new E.Sheet(s.cells);
        log = Array.isArray(s.log) ? s.log : [];
        frozen = !!s.frozen;
        return true;
      }
    } catch (e) { /* unreadable draft: start from the file */ }
    return false;
  }

  function say(text) { statusLine.textContent = text || ""; }
  const now = () => new Date().toLocaleTimeString("el-GR", { hour: "2-digit", minute: "2-digit" });
  function note(entry) {
    log.push(Object.assign({ at: now() }, entry));
    if (!logBox.hidden) renderLog();
  }

  // ---- drawing -----------------------------------------------------------
  const inSel = (r, c) => r >= Math.min(sel.r, sel.r2) && r <= Math.max(sel.r, sel.r2)
    && c >= Math.min(sel.c, sel.c2) && c <= Math.max(sel.c, sel.c2);
  const box = () => ({ r1: Math.min(sel.r, sel.r2), r2: Math.max(sel.r, sel.r2),
                       c1: Math.min(sel.c, sel.c2), c2: Math.max(sel.c, sel.c2) });

  function render() {
    values = sheet.evaluate();
    const rows = sheet.rows, cols = sheet.cols;
    const head = ["<thead><tr><th class='corner'></th>"];
    for (let c = 0; c < cols; c++) head.push(`<th class="col${c === sel.c ? " on" : ""}">${E.colName(c)}</th>`);
    head.push("</tr></thead><tbody>");
    const body = [];
    for (let r = 0; r < rows; r++) {
      const cls = ["row"];
      if (r === 0 && frozen) cls.push("frozen");
      body.push(`<tr class="${cls.join(" ")}"><th class="rown${r === sel.r ? " on" : ""}">${r + 1}</th>`);
      for (let c = 0; c < cols; c++) {
        const v = values[r][c];
        const k = [];
        if (typeof v === "number") k.push("num");
        if (E.isErr(v)) k.push("err");
        if (sheet.raw(r, c).startsWith("=")) k.push("fx");
        if (inSel(r, c)) k.push("sel");
        if (r === sel.r && c === sel.c) k.push("active");
        body.push(`<td data-r="${r}" data-c="${c}" class="${k.join(" ")}">${escapeHtml(E.display(v))}</td>`);
      }
      body.push("</tr>");
    }
    table.innerHTML = head.join("") + body.join("") + "</tbody>";
    nameBox.textContent = E.cellName(sel.r, sel.c) + (sel.r !== sel.r2 || sel.c !== sel.c2
      ? ":" + E.cellName(sel.r2, sel.c2) : "");
    if (!editor || editor.viaBar !== true) bar.value = sheet.raw(sel.r, sel.c);
    $("[data-act=freeze]").setAttribute("aria-pressed", String(frozen));
    summary();
  }
  function escapeHtml(s) {
    return s.replace(/[&<>"]/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[ch]));
  }

  // The figures a spreadsheet shows at the bottom for a selected range.
  function summary() {
    const b = box();
    if (b.r1 === b.r2 && b.c1 === b.c2) {
      const v = values[sel.r] && values[sel.r][sel.c];
      say(E.isErr(v) ? errorHelp(v.code) : "");
      return;
    }
    const nums = [];
    let filled = 0;
    for (let r = b.r1; r <= b.r2; r++) for (let c = b.c1; c <= b.c2; c++) {
      const v = values[r] && values[r][c];
      if (v !== null && v !== "" && v !== undefined) filled++;
      if (typeof v === "number") nums.push(v);
    }
    const sum = nums.reduce((a, x) => a + x, 0);
    say(`Μη κενά: ${filled} · Αριθμοί: ${nums.length}` +
        (nums.length ? ` · Άθροισμα: ${E.display(sum)} · Μέσος: ${E.display(sum / nums.length)}` : ""));
  }
  function errorHelp(code) {
    return {
      "#DIV/0!": "#DIV/0!: διαίρεση με το μηδέν, ή μέσος όρος χωρίς αριθμούς.",
      "#NAME?": "#NAME?: άγνωστη συνάρτηση ή λάθος στο όνομα. Δείτε τη «Βοήθεια».",
      "#VALUE!": "#VALUE!: πράξη με κείμενο, ή τύπος που δεν ολοκληρώθηκε (παρένθεση, όρισμα).",
      "#REF!": "#REF!: αναφορά σε κελί που δεν υπάρχει πια (π.χ. διαγραμμένη γραμμή).",
      "#CIRC!": "#CIRC!: ο τύπος αναφέρεται στον εαυτό του.",
      "#NUM!": "#NUM!: αριθμητικά αδύνατο (π.χ. ρίζα αρνητικού).",
    }[code] || code;
  }

  // ---- editing -----------------------------------------------------------
  function td(r, c) { return table.querySelector(`td[data-r="${r}"][data-c="${c}"]`); }

  function startEdit(initial) {
    if (editor) return;
    const cell = td(sel.r, sel.c);
    if (!cell) return;
    const input = document.createElement("input");
    input.type = "text";
    input.className = "sheet-edit";
    input.spellcheck = false;
    input.autocomplete = "off";
    const before = sheet.raw(sel.r, sel.c);
    input.value = initial === undefined ? before : initial;
    cell.textContent = "";
    cell.appendChild(input);
    editor = { r: sel.r, c: sel.c, input, before, pick: null };
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
    bar.value = input.value;
    input.addEventListener("input", () => { bar.value = input.value; editor && (editor.pick = null); });
    input.addEventListener("keydown", e => {
      if (e.key === "Enter") { e.preventDefault(); commit(); move(e.shiftKey ? -1 : 1, 0); }
      else if (e.key === "Tab") { e.preventDefault(); commit(); move(0, e.shiftKey ? -1 : 1); }
      else if (e.key === "Escape") { e.preventDefault(); cancel(); }
    });
    input.addEventListener("blur", () => setTimeout(() => {
      if (editor && editor.input === input && document.activeElement !== bar && !dragging) commit();
    }, 0));
  }
  function text() { return editor ? editor.input.value : bar.value; }
  function commit(viaBarText) {
    const r = editor ? editor.r : sel.r, c = editor ? editor.c : sel.c;
    const before = editor ? editor.before : sheet.raw(r, c);
    const after = viaBarText !== undefined ? viaBarText : text();
    editor = null;
    if (after !== before) {
      sheet.set(r, c, after);
      note({ kind: "edit", cell: E.cellName(r, c), from: before, to: after });
      save();
    }
    render();
    grid.focus({ preventScroll: true });
  }
  function cancel() { editor = null; render(); grid.focus({ preventScroll: true }); }

  // While a formula is being typed, a click on a cell writes its address in,
  // and dragging widens it to a range, as a spreadsheet's point mode does.
  function pointing() {
    if (!editor) return false;
    const v = editor.input.value;
    if (!v.startsWith("=")) return false;
    if (editor.pick) return true;
    return /[=(,;+\-*/^&<>:]\s*$/.test(v);
  }
  function pick(r, c, extend) {
    const input = editor.input;
    if (!extend || !editor.pick) {
      editor.pick = { start: input.value.length, r, c };
      input.value += E.cellName(r, c);
    } else {
      const p = editor.pick;
      const r1 = Math.min(p.r, r), r2 = Math.max(p.r, r), c1 = Math.min(p.c, c), c2 = Math.max(p.c, c);
      const ref = r1 === r2 && c1 === c2 ? E.cellName(r1, c1) : E.cellName(r1, c1) + ":" + E.cellName(r2, c2);
      input.value = input.value.slice(0, p.start) + ref;
    }
    bar.value = input.value;
  }

  function move(dr, dc, extend) {
    const r = Math.max(0, Math.min(sheet.rows - 1, (extend ? sel.r2 : sel.r) + dr));
    const c = Math.max(0, Math.min(sheet.cols - 1, (extend ? sel.c2 : sel.c) + dc));
    if (extend) { sel.r2 = r; sel.c2 = c; }
    else sel = { r, c, r2: r, c2: c };
    render();
    const cell = td(extend ? sel.r2 : sel.r, extend ? sel.c2 : sel.c);
    if (cell) cell.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  function clearSelection() {
    const b = box();
    let changed = 0;
    for (let r = b.r1; r <= b.r2; r++) for (let c = b.c1; c <= b.c2; c++) {
      const before = sheet.raw(r, c);
      if (before !== "") { sheet.set(r, c, ""); note({ kind: "edit", cell: E.cellName(r, c), from: before, to: "" }); changed++; }
    }
    if (changed) save();
    render();
  }

  // ---- copy, paste, fill -------------------------------------------------
  function copy(e) {
    const b = box();
    const cells = [];
    for (let r = b.r1; r <= b.r2; r++) {
      const row = [];
      for (let c = b.c1; c <= b.c2; c++) row.push(sheet.raw(r, c));
      cells.push(row);
    }
    clip = { r: b.r1, c: b.c1, cells };
    const tsv = cells.map((row, i) => row.map((_, j) => E.display(values[b.r1 + i][b.c1 + j])).join("\t")).join("\n");
    e.clipboardData.setData("text/plain", tsv);
    e.preventDefault();
    say("Αντιγράφηκε " + nameBox.textContent + ".");
  }
  function paste(e) {
    const textData = e.clipboardData.getData("text/plain");
    e.preventDefault();
    const b = box();
    const mine = clip && clip.cells.map((row, i) => row.map((_, j) => E.display(values[clip.r + i] && values[clip.r + i][clip.c + j] !== undefined ? values[clip.r + i][clip.c + j] : null)).join("\t")).join("\n") === textData.replace(/\r/g, "").replace(/\n$/, "");
    const source = mine ? clip.cells : textData.replace(/\r/g, "").replace(/\n$/, "").split("\n").map(l => l.split("\t"));
    const h = source.length, w = source[0].length;
    // one copied cell into a larger selection fills all of it, as Excel does
    const rows = h === 1 && w === 1 ? b.r2 - b.r1 + 1 : h;
    const cols = h === 1 && w === 1 ? b.c2 - b.c1 + 1 : w;
    for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
      const r = b.r1 + i, c = b.c1 + j;
      let t = source[i % h][j % w];
      if (mine) t = E.shift(t, r - (clip.r + (i % h)), c - (clip.c + (j % w)));
      const before = sheet.raw(r, c);
      if (t !== before) { sheet.set(r, c, t); note({ kind: "edit", cell: E.cellName(r, c), from: before, to: t }); }
    }
    sel.r2 = b.r1 + rows - 1; sel.c2 = b.c1 + cols - 1; sel.r = b.r1; sel.c = b.c1;
    save();
    render();
  }
  function lastFilledRow() {
    for (let r = sheet.rows - 1; r >= 0; r--) if (sheet.cells[r].some(t => t !== "")) return r;
    return 0;
  }
  function fillDown() {
    const b = box();
    let r2 = b.r2;
    if (b.r1 === b.r2) {
      // one cell: down to the last row that holds data, as a double click on
      // Excel's fill handle does
      r2 = lastFilledRow();
      if (r2 <= b.r1) { say("Δεν υπάρχουν γραμμές από κάτω για να συμπληρωθούν."); return; }
    }
    let n = 0;
    for (let c = b.c1; c <= b.c2; c++) {
      const top = sheet.raw(b.r1, c);
      for (let r = b.r1 + 1; r <= r2; r++) {
        const t = E.shift(top, r - b.r1, 0);
        const before = sheet.raw(r, c);
        if (t !== before) { sheet.set(r, c, t); n++; }
      }
    }
    if (n) note({ kind: "fill", cell: E.cellName(b.r1, b.c1) + (b.c2 > b.c1 ? ":" + E.cellName(b.r1, b.c2) : ""), to: E.cellName(r2, b.c2) });
    sel.r2 = r2;
    save();
    render();
    say(n ? `Συμπληρώθηκαν ${n} κελιά ως τη γραμμή ${r2 + 1}.` : "");
  }

  // ---- rows and columns --------------------------------------------------
  function deleteRows() {
    const b = box();
    if (b.r1 === 0) { say("Η 1η γραμμή είναι οι επικεφαλίδες· δεν διαγράφεται."); return; }
    const ids = [];
    for (let r = b.r2; r >= b.r1; r--) {
      ids.unshift((r + 1) + (sheet.raw(r, 0) ? " (" + sheet.raw(r, 0) + ")" : ""));
      sheet.deleteRow(r);
    }
    note({ kind: "delrow", rows: ids.join(", ") });
    const r = Math.min(b.r1, sheet.rows - 1);
    sel = { r, c: sel.c, r2: r, c2: sel.c };
    save();
    render();
    say("Διαγράφηκε: γραμμή " + ids.join(", ") + ". Οι τύποι προσαρμόστηκαν.");
  }

  // ---- the change log ----------------------------------------------------
  function renderLog() {
    const ol = logBox.querySelector("ol");
    ol.innerHTML = "";
    if (!log.length) { ol.innerHTML = "<li class='empty'>Καμία αλλαγή ακόμη.</li>"; return; }
    for (const e of log) {
      const li = document.createElement("li");
      if (e.kind === "edit") li.textContent = `${e.at} · ${e.cell}: «${e.from}» → «${e.to}»`;
      else if (e.kind === "delrow") li.textContent = `${e.at} · διαγραφή γραμμής ${e.rows}`;
      else if (e.kind === "fill") li.textContent = `${e.at} · συμπλήρωση από ${e.cell} ως ${e.to}`;
      else if (e.kind === "addcol") li.textContent = `${e.at} · νέα στήλη ${e.cell}`;
      else if (e.kind === "addrow") li.textContent = `${e.at} · νέα γραμμή ${e.cell}`;
      ol.appendChild(li);
    }
  }

  // ---- the chart ---------------------------------------------------------
  const dialog = document.createElement("dialog");
  dialog.className = "sheet-chart";
  dialog.innerHTML = `
    <div class="material-bar">
      <strong>Γράφημα</strong>
      <button type="button" class="chart-png">Λήψη εικόνας</button>
      <button type="button" class="material-close chart-close" aria-label="Κλείσιμο">✕</button>
    </div>
    <div class="chart-controls">
      <label>Στήλη <select class="chart-col"></select></label>
      <label><input type="radio" name="chart-kind" value="hist" checked> Ιστόγραμμα</label>
      <label><input type="radio" name="chart-kind" value="bar"> Ραβδόγραμμα ανά γραμμή</label>
      <label class="chart-bins-label">Διαστήματα <input type="number" class="chart-bins" min="2" max="20" step="1"></label>
    </div>
    <div class="chart-area"></div>`;
  document.body.appendChild(dialog);
  const q = s => dialog.querySelector(s);

  function numericColumns() {
    const out = [];
    for (let c = 0; c < sheet.cols; c++) {
      let n = 0;
      for (let r = 1; r < sheet.rows; r++) if (typeof values[r][c] === "number") n++;
      if (n >= 2) out.push(c);
    }
    return out;
  }
  function openChart() {
    const cols = numericColumns();
    if (!cols.length) { say("Δεν υπάρχει στήλη με αριθμούς για γράφημα."); return; }
    const select = q(".chart-col");
    select.innerHTML = cols.map(c => `<option value="${c}">${E.colName(c)} · ${escapeHtml(E.display(values[0][c]) || "")}</option>`).join("");
    if (cols.includes(sel.c)) select.value = String(sel.c);
    q(".chart-bins").value = "";
    drawChart();
    dialog.showModal();
  }
  function drawChart() {
    const c = Number(q(".chart-col").value);
    const kind = dialog.querySelector("input[name=chart-kind]:checked").value;
    q(".chart-bins-label").hidden = kind !== "hist";
    const points = [];
    for (let r = 1; r < sheet.rows; r++) {
      const v = values[r][c];
      if (typeof v === "number") points.push({ v, label: E.display(values[r][0]) || String(r + 1) });
    }
    const title = (E.display(values[0][c]) || E.colName(c)) + ` (n = ${points.length})`;
    let bars;
    if (kind === "bar") {
      bars = points.map(p => ({ label: p.label, value: p.v, show: E.display(p.v) }));
    } else {
      const xs = points.map(p => p.v);
      const lo = Math.min(...xs), hi = Math.max(...xs);
      let k = Number(q(".chart-bins").value) || Math.ceil(Math.log2(xs.length) + 1);
      k = Math.max(2, Math.min(20, k));
      q(".chart-bins").value = String(k);
      const width = (hi - lo) / k || 1;
      const counts = new Array(k).fill(0);
      for (const x of xs) counts[Math.min(k - 1, Math.floor((x - lo) / width))]++;
      const edge = i => E.display(Math.round((lo + i * width) * 10) / 10);
      bars = counts.map((n, i) => ({ label: edge(i) + "-" + edge(i + 1), value: n, show: String(n) }));
    }
    q(".chart-area").innerHTML = svgChart(title, bars, kind === "hist" ? "Συχνότητα" : E.display(values[0][c]) || "");
  }
  function svgChart(title, bars, yLabel) {
    const W = 720, H = 360, L = 56, R = 16, T = 36, B = 78;
    const max = Math.max(1, ...bars.map(b => b.value));
    const step = niceStep(max);
    const top = Math.ceil(max / step) * step;
    const iw = W - L - R, ih = H - T - B;
    const bw = iw / bars.length;
    const y = v => T + ih - (v / top) * ih;
    const parts = [`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" role="img" aria-label="${escapeHtml(title)}">`,
      `<rect width="${W}" height="${H}" fill="#fff"/>`,
      `<text x="${W / 2}" y="22" text-anchor="middle" font-size="16" font-weight="600" fill="#1F4E79">${escapeHtml(title)}</text>`];
    for (let v = 0; v <= top + 1e-9; v += step) {
      parts.push(`<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" stroke="#e3e3e3"/>`,
        `<text x="${L - 6}" y="${y(v) + 4}" text-anchor="end" font-size="11" fill="#404040">${E.display(v)}</text>`);
    }
    parts.push(`<text transform="translate(14 ${T + ih / 2}) rotate(-90)" text-anchor="middle" font-size="12" fill="#404040">${escapeHtml(yLabel)}</text>`);
    bars.forEach((b, i) => {
      const x = L + i * bw + bw * 0.12, w = bw * 0.76, yy = y(Math.max(0, b.value));
      parts.push(`<rect x="${x}" y="${yy}" width="${w}" height="${T + ih - yy}" fill="#2E74B5"/>`);
      if (bw > 18) parts.push(`<text x="${x + w / 2}" y="${yy - 4}" text-anchor="middle" font-size="11" fill="#1F4E79">${escapeHtml(b.show)}</text>`);
      const lx = L + i * bw + bw / 2, ly = T + ih + 14;
      parts.push(`<text x="${lx}" y="${ly}" text-anchor="end" font-size="11" fill="#404040" transform="rotate(-40 ${lx} ${ly})">${escapeHtml(b.label)}</text>`);
    });
    parts.push(`<line x1="${L}" x2="${W - R}" y1="${T + ih}" y2="${T + ih}" stroke="#404040"/></svg>`);
    return parts.join("");
  }
  function niceStep(max) {
    const raw = max / 5;
    const mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const n = raw / mag;
    return Math.max(1, (n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10) * mag);
  }
  function chartPng() {
    const svg = q(".chart-area svg");
    if (!svg) return;
    const img = new Image();
    const url = URL.createObjectURL(new Blob([svg.outerHTML], { type: "image/svg+xml" }));
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = 1440; canvas.height = 720;
      canvas.getContext("2d").drawImage(img, 0, 0, 1440, 720);
      URL.revokeObjectURL(url);
      canvas.toBlob(b => download(b, "grafima.png"), "image/png");
    };
    img.src = url;
  }
  q(".chart-col").addEventListener("change", drawChart);
  q(".chart-bins").addEventListener("change", drawChart);
  dialog.querySelectorAll("input[name=chart-kind]").forEach(r => r.addEventListener("change", drawChart));
  q(".chart-close").addEventListener("click", () => dialog.close());
  q(".chart-png").addEventListener("click", chartPng);
  dialog.addEventListener("click", e => { if (e.target === dialog) dialog.close(); });

  function download(blob, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // ---- wiring ------------------------------------------------------------
  table.addEventListener("mousedown", e => {
    const cell = e.target.closest("td");
    if (!cell || e.target.classList.contains("sheet-edit")) return;
    const r = Number(cell.dataset.r), c = Number(cell.dataset.c);
    if (pointing()) {
      e.preventDefault();
      pick(r, c, e.shiftKey);
      dragging = "pick";
      editor.input.focus();
      return;
    }
    if (editor) commit();
    e.preventDefault();
    if (e.shiftKey) { sel.r2 = r; sel.c2 = c; } else sel = { r, c, r2: r, c2: c };
    dragging = "select";
    render();
    grid.focus({ preventScroll: true });
  });
  table.addEventListener("mouseover", e => {
    if (!dragging) return;
    const cell = e.target.closest("td");
    if (!cell) return;
    const r = Number(cell.dataset.r), c = Number(cell.dataset.c);
    if (dragging === "pick" && editor) { pick(r, c, true); return; }
    if (dragging === "select" && (sel.r2 !== r || sel.c2 !== c)) { sel.r2 = r; sel.c2 = c; render(); }
  });
  document.addEventListener("mouseup", () => {
    if (dragging === "pick" && editor) editor.input.focus();
    dragging = false;
  });
  table.addEventListener("dblclick", e => { if (e.target.closest("td") && !editor) startEdit(); });

  grid.addEventListener("keydown", e => {
    if (editor) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[e.key]) { e.preventDefault(); move(...arrows[e.key], e.shiftKey); }
    else if (e.key === "Enter" || e.key === "F2") { e.preventDefault(); startEdit(); }
    else if (e.key === "Tab") { e.preventDefault(); move(0, e.shiftKey ? -1 : 1); }
    else if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); clearSelection(); }
    else if (ctrl && (e.key === "d" || e.key === "D")) { e.preventDefault(); fillDown(); }
    else if (ctrl && (e.key === "a" || e.key === "A")) { e.preventDefault(); sel = { r: 0, c: 0, r2: sheet.rows - 1, c2: sheet.cols - 1 }; render(); }
    else if (!ctrl && !e.altKey && e.key.length === 1) { e.preventDefault(); startEdit(e.key); }
  });
  grid.addEventListener("copy", e => { if (!editor) copy(e); });
  grid.addEventListener("paste", e => { if (!editor) paste(e); });

  bar.addEventListener("focus", () => { if (editor) { editor.viaBar = true; } });
  bar.addEventListener("input", () => { if (editor) editor.input.value = bar.value; });
  bar.addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); const t = bar.value; if (editor) editor.input.value = t; commit(t); move(1, 0); }
    else if (e.key === "Escape") { e.preventDefault(); cancel(); }
  });
  bar.addEventListener("change", () => { if (!editor && bar.value !== sheet.raw(sel.r, sel.c)) commit(bar.value); });

  dock.querySelector(".sheet-actions").addEventListener("click", e => {
    const btn = e.target.closest("button[data-act]");
    if (!btn) return;
    if (editor) commit();
    switch (btn.dataset.act) {
      case "fill": fillDown(); break;
      case "addcol":
        sheet.addColumn();
        note({ kind: "addcol", cell: E.colName(sheet.cols - 1) });
        sel = { r: 0, c: sheet.cols - 1, r2: 0, c2: sheet.cols - 1 };
        save(); render(); say("Νέα στήλη " + E.colName(sheet.cols - 1) + ". Γράψτε στο κελί " + E.cellName(0, sheet.cols - 1) + " το όνομά της.");
        break;
      case "addrow":
        sheet.addRow();
        note({ kind: "addrow", cell: String(sheet.rows) });
        sel = { r: sheet.rows - 1, c: 0, r2: sheet.rows - 1, c2: 0 };
        save(); render();
        break;
      case "delrow": deleteRows(); break;
      case "freeze": frozen = !frozen; save(); render(); break;
      case "chart": openChart(); break;
      case "log":
        logBox.hidden = !logBox.hidden;
        btn.setAttribute("aria-pressed", String(!logBox.hidden));
        if (!logBox.hidden) renderLog();
        break;
      case "csv": {
        const stem = cfg.name.replace(/\.csv$/i, "");
        download(new Blob([E.toCsv(values)], { type: "text/csv;charset=utf-8" }), stem + "_epexergasmeno.csv");
        break;
      }
      case "reset":
        if (!confirm("Επαναφορά στο αρχικό αρχείο; Οι αλλαγές και το ιστορικό τους θα χαθούν.")) break;
        sheet = new E.Sheet(pristine); log = []; frozen = false;
        sel = { r: 1, c: 0, r2: 1, c2: 0 };
        save(); render(); if (!logBox.hidden) renderLog();
        break;
      case "help": {
        const help = $(".sheet-help");
        help.hidden = !help.hidden;
        btn.setAttribute("aria-expanded", String(!help.hidden));
        break;
      }
      case "size": dock.classList.toggle("tall"); syncSpace(); break;
      case "close": show(false); break;
    }
  });

  function syncSpace() {
    document.body.style.paddingBottom = dock.hidden ? "" : dock.offsetHeight + "px";
  }
  function show(open) {
    dock.hidden = !open;
    toggle.setAttribute("aria-pressed", String(open));
    syncSpace();
    if (open) { render(); grid.focus({ preventScroll: true }); }
    try { sessionStorage.setItem(key + ":open", open ? "1" : ""); } catch (e) { /* ignore */ }
  }
  window.addEventListener("resize", syncSpace);

  async function load() {
    toggle.disabled = true;
    try {
      const res = await fetch(cfg.csv);
      if (!res.ok) throw new Error(res.status);
      pristine = E.parseCsv(await res.text());
      if (!restore()) sheet = new E.Sheet(pristine);
      toggle.disabled = false;
      let reopen = false;
      try { reopen = sessionStorage.getItem(key + ":open") === "1"; } catch (e) { /* ignore */ }
      if (reopen) show(true);
    } catch (e) {
      toggle.title = "Το αρχείο δεδομένων δεν φορτώθηκε.";
    }
  }
  toggle.addEventListener("click", () => show(dock.hidden));
  load();
})();
