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
        <button type="button" data-act="undo" title="Αναίρεση της τελευταίας αλλαγής (Ctrl+Z)" aria-label="Αναίρεση" disabled>↶</button>
        <button type="button" data-act="redo" title="Επανάληψη της αλλαγής που αναιρέθηκε (Ctrl+Y)" aria-label="Επανάληψη" disabled>↷</button>
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
      Ctrl+C / Ctrl+V για αντιγραφή και επικόλληση (οι τύποι προσαρμόζονται), Ctrl+D για συμπλήρωση προς τα κάτω,
      Ctrl+Z για αναίρεση και Ctrl+Y για επανάληψη (και με τα κουμπιά ↶ ↷).
      Το <code>$</code> κρατά σταθερή μια στήλη ή γραμμή (<code>$D$2</code>).</p>
      <p>Το μικρό τετράγωνο κάτω δεξιά στην επιλογή συμπληρώνει όταν το σύρετε (κάτω ή δεξιά)· με διπλό κλικ
      συμπληρώνει ως το τέλος των δεδομένων. Δύο αριθμοί στη σειρά (1, 2) συνεχίζουν ως ακολουθία.
      Κλικ στο γράμμα μιας στήλης ή στον αριθμό μιας γραμμής την επιλέγει ολόκληρη. Το δεξί κλικ ανοίγει
      μενού με εισαγωγή και διαγραφή γραμμών και στηλών, αντιγραφή και επικόλληση. Τα «+ Στήλη» και
      «+ Γραμμή» προσθέτουν δίπλα στο επιλεγμένο κελί.</p>
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
  let fillTo = null;  // {r, c}: where a drag of the fill handle has reached
  let downAt = null;  // where the button went down: a redraw under a still pointer is not a drag
  // Undo keeps whole copies of the cells. Every change ends in save(), so save()
  // is where the state before it is taken; the data file is small enough that a
  // copy per change costs nothing. Kept for this visit only, as Excel does.
  const UNDO_LIMIT = 100;
  let undo = [];
  let redo = [];
  let saved = null;   // the cells as of the last save, as JSON: what an undo returns to

  function save() {
    const cells = JSON.stringify(sheet.cells);
    if (saved !== null && cells !== saved) {
      undo.push(saved);
      if (undo.length > UNDO_LIMIT) undo.shift();
      redo = [];
    }
    saved = cells;
    syncUndo();
    store();
  }
  function store() {
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

  function syncUndo() {
    $("[data-act=undo]").disabled = !undo.length;
    $("[data-act=redo]").disabled = !redo.length;
  }
  // Steps one change back (or forward) and selects the cells it touched, so the
  // student sees what came back. A row or column added or deleted changes the
  // shape, and then the selection only stays inside the sheet.
  function step(from, to, kind) {
    if (editor) commit();
    if (!from.length) { say(kind === "undo" ? "Δεν υπάρχει αλλαγή για αναίρεση." : "Δεν υπάρχει αλλαγή για επανάληψη."); return; }
    const before = sheet.cells;
    const target = JSON.parse(from.pop());
    to.push(saved);
    sheet = new E.Sheet(target);
    saved = JSON.stringify(sheet.cells);
    let b = null;
    if (before.length === sheet.rows && before[0].length === sheet.cols) {
      for (let r = 0; r < sheet.rows; r++) for (let c = 0; c < sheet.cols; c++) {
        if (before[r][c] === sheet.cells[r][c]) continue;
        b = b ? { r1: Math.min(b.r1, r), r2: Math.max(b.r2, r), c1: Math.min(b.c1, c), c2: Math.max(b.c2, c) }
              : { r1: r, r2: r, c1: c, c2: c };
      }
    }
    if (b) sel = { r: b.r1, c: b.c1, r2: b.r2, c2: b.c2 };
    else {
      const r = Math.min(sel.r, sheet.rows - 1), c = Math.min(sel.c, sheet.cols - 1);
      sel = { r, c, r2: r, c2: c };
    }
    note({ kind });
    syncUndo();
    store();
    render();
    say(kind === "undo" ? "Αναιρέθηκε η τελευταία αλλαγή." : "Έγινε ξανά η αλλαγή.");
    grid.focus({ preventScroll: true });
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
    const head = ["<thead><tr><th class='corner' title='Επιλογή όλων'></th>"];
    for (let c = 0; c < cols; c++) head.push(`<th class="col${inSel(0, c) || (c >= box().c1 && c <= box().c2) ? " on" : ""}" data-col="${c}" title="Επιλογή στήλης">${E.colName(c)}</th>`);
    head.push("</tr></thead><tbody>");
    const body = [];
    for (let r = 0; r < rows; r++) {
      const cls = ["row"];
      if (r === 0 && frozen) cls.push("frozen");
      body.push(`<tr class="${cls.join(" ")}"><th class="rown${r >= box().r1 && r <= box().r2 ? " on" : ""}" data-row="${r}" title="Επιλογή γραμμής">${r + 1}</th>`);
      for (let c = 0; c < cols; c++) {
        const v = values[r][c];
        const k = [];
        if (typeof v === "number") k.push("num");
        if (E.isErr(v)) k.push("err");
        if (sheet.raw(r, c).startsWith("=")) k.push("fx");
        if (inSel(r, c)) k.push("sel");
        if (r === sel.r && c === sel.c) k.push("active");
        if (inPreview(r, c)) k.push("fillprev");
        const b = box();
        const handle = !editor && r === (fillTo ? Math.max(b.r2, fillTo.r) : b.r2) && c === (fillTo ? Math.max(b.c2, fillTo.c) : b.c2)
          ? '<span class="fill-handle" title="Σύρετε για συμπλήρωση· διπλό κλικ ως το τέλος των δεδομένων"></span>' : "";
        body.push(`<td data-r="${r}" data-c="${c}" class="${k.join(" ")}">${escapeHtml(E.display(v))}${handle}</td>`);
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
  function inPreview(r, c) {
    if (!fillTo) return false;
    const b = box();
    return r >= b.r1 && c >= b.c1 && r <= Math.max(b.r2, fillTo.r) && c <= Math.max(b.c2, fillTo.c) && !inSel(r, c);
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
    // stopped here: the grid's own keydown would otherwise read the same Enter
    // after the commit and open the next cell for editing
    input.addEventListener("keydown", e => {
      if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); commit(); move(e.shiftKey ? -1 : 1, 0); }
      else if (e.key === "Tab") { e.preventDefault(); e.stopPropagation(); commit(); move(0, e.shiftKey ? -1 : 1); }
      else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); cancel(); }
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
    if (e) { e.clipboardData.setData("text/plain", tsv); e.preventDefault(); }
    else if (navigator.clipboard) navigator.clipboard.writeText(tsv).catch(() => {});
    clip.text = tsv;
    say("Αντιγράφηκε " + nameBox.textContent + ".");
  }
  function paste(e) {
    e.preventDefault();
    pasteText(e.clipboardData.getData("text/plain"));
  }
  function pasteText(textData) {
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

  // The fill handle, as in a spreadsheet: drag the selection down or across and
  // formulas are copied with their references moved, text and single numbers
  // repeat, and two or more numbers in a steady step continue as a series.
  function series(cells) {
    if (cells.length < 2) return null;
    const xs = cells.map(t => E.literal(t));
    if (!xs.every(x => typeof x === "number") || cells.some(t => t.startsWith("="))) return null;
    const step = xs[1] - xs[0];
    for (let i = 2; i < xs.length; i++) if (Math.abs(xs[i] - xs[i - 1] - step) > 1e-9) return null;
    return { last: xs[xs.length - 1], step };
  }
  function dragFill(b, toR, toC) {
    let n = 0;
    const put = (r, c, t) => { if (sheet.raw(r, c) !== t) { sheet.set(r, c, t); n++; } };
    const tidy = x => String(Math.round(x * 1e9) / 1e9);
    if (toR > b.r2) {
      const h = b.r2 - b.r1 + 1;
      for (let c = b.c1; c <= b.c2; c++) {
        const src = [];
        for (let r = b.r1; r <= b.r2; r++) src.push(sheet.raw(r, c));
        const run = series(src);
        for (let r = b.r2 + 1; r <= toR; r++) {
          const i = (r - b.r1) % h, from = b.r1 + i;
          put(r, c, run ? tidy(run.last + run.step * (r - b.r2)) : E.shift(src[i], r - from, 0));
        }
      }
    } else if (toC > b.c2) {
      const w = b.c2 - b.c1 + 1;
      for (let r = b.r1; r <= b.r2; r++) {
        const src = [];
        for (let c = b.c1; c <= b.c2; c++) src.push(sheet.raw(r, c));
        const run = series(src);
        for (let c = b.c2 + 1; c <= toC; c++) {
          const i = (c - b.c1) % w, from = b.c1 + i;
          put(r, c, run ? tidy(run.last + run.step * (c - b.c2)) : E.shift(src[i], 0, c - from));
        }
      }
    } else return;
    note({ kind: "fill", cell: E.cellName(b.r1, b.c1) + ":" + E.cellName(b.r2, b.c2), to: E.cellName(Math.max(toR, b.r2), Math.max(toC, b.c2)) });
    sel = { r: b.r1, c: b.c1, r2: Math.max(toR, b.r2), c2: Math.max(toC, b.c2) };
    save();
    say(n ? `Συμπληρώθηκαν ${n} κελιά.` : "");
  }

  // ---- rows and columns --------------------------------------------------
  function insertColumn(at) {
    sheet.insertColumn(at);
    note({ kind: "addcol", cell: E.colName(at) });
    sel = { r: 0, c: at, r2: 0, c2: at };
    save(); render();
    say("Νέα στήλη " + E.colName(at) + ". Γράψτε στο " + E.cellName(0, at) + " το όνομά της· οι τύποι προσαρμόστηκαν.");
  }
  function insertRow(at) {
    if (at === 0) { say("Η 1η γραμμή είναι οι επικεφαλίδες· η νέα γραμμή μπαίνει από κάτω τους."); at = 1; }
    sheet.insertRow(at);
    note({ kind: "addrow", cell: String(at + 1) });
    sel = { r: at, c: sel.c, r2: at, c2: sel.c };
    save(); render();
    say("Νέα γραμμή " + (at + 1) + "· οι τύποι προσαρμόστηκαν.");
  }
  function deleteColumns() {
    const b = box();
    if (b.c2 - b.c1 + 1 >= sheet.cols) { say("Δεν διαγράφονται όλες οι στήλες."); return; }
    const names = [];
    for (let c = b.c2; c >= b.c1; c--) {
      names.unshift(E.colName(c) + (sheet.raw(0, c) ? " (" + sheet.raw(0, c) + ")" : ""));
      sheet.deleteColumn(c);
    }
    note({ kind: "delcol", cols: names.join(", ") });
    const c = Math.min(b.c1, sheet.cols - 1);
    sel = { r: sel.r, c, r2: sel.r, c2: c };
    save(); render();
    say("Διαγράφηκε: στήλη " + names.join(", ") + ". Οι τύποι προσαρμόστηκαν.");
  }
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
      else if (e.kind === "delcol") li.textContent = `${e.at} · διαγραφή στήλης ${e.cols}`;
      else if (e.kind === "addrow") li.textContent = `${e.at} · νέα γραμμή ${e.cell}`;
      else if (e.kind === "undo") li.textContent = `${e.at} · αναίρεση της προηγούμενης αλλαγής`;
      else if (e.kind === "redo") li.textContent = `${e.at} · επανάληψη της αλλαγής που αναιρέθηκε`;
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
      <label><input type="radio" name="chart-kind" value="box"> Θηκόγραμμα</label>
      <label class="chart-group-label">Ομάδες <select class="chart-group"></select></label>
      <label class="chart-bins-label">Διαστήματα <input type="number" class="chart-bins" min="2" max="20" step="1"></label>
    </div>
    <div class="chart-area"></div>
    <div class="chart-notes"></div>`;
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
  // A boxplot as the lab teaches it: the box from Q1 to Q3 with the median,
  // whiskers to the furthest values inside 1.5 x IQR, and anything beyond them
  // drawn as a point named by its ID, so the student sees which row it is.
  function boxStats(points) {
    const xs = points.map(p => p.v);
    const q1 = E.quantile(xs, 0.25), med = E.quantile(xs, 0.5), q3 = E.quantile(xs, 0.75);
    const iqr = q3 - q1, lo = q1 - 1.5 * iqr, hi = q3 + 1.5 * iqr;
    const inside = xs.filter(x => x >= lo && x <= hi);
    return {
      n: xs.length, min: Math.min(...xs), max: Math.max(...xs), q1, med, q3, iqr, lo, hi,
      wlo: Math.min(...inside), whi: Math.max(...inside),
      out: points.filter(p => p.v < lo || p.v > hi),
    };
  }
  function svgBox(title, groups, yLabel) {
    const W = 720, H = 380, L = 64, R = 24, T = 40, B = 46;
    const all = groups.flatMap(g => g.points.map(p => p.v));
    let min = Math.min(...all), max = Math.max(...all);
    if (min === max) { min -= 1; max += 1; }
    const step = niceStep(max - min);
    const bottom = Math.floor(min / step) * step, top = Math.ceil(max / step) * step;
    const ih = H - T - B, iw = W - L - R;
    const y = v => T + ih - ((v - bottom) / (top - bottom)) * ih;
    const slot = iw / groups.length;
    const parts = [`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" role="img" aria-label="${escapeHtml(title)}">`,
      `<rect width="${W}" height="${H}" fill="#fff"/>`,
      `<text x="${W / 2}" y="22" text-anchor="middle" font-size="16" font-weight="600" fill="#1F4E79">${escapeHtml(title)}</text>`];
    for (let v = bottom; v <= top + 1e-9; v += step) {
      parts.push(`<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" stroke="#e3e3e3"/>`,
        `<text x="${L - 6}" y="${y(v) + 4}" text-anchor="end" font-size="11" fill="#404040">${E.display(Math.round(v * 1e6) / 1e6)}</text>`);
    }
    parts.push(`<text transform="translate(16 ${T + ih / 2}) rotate(-90)" text-anchor="middle" font-size="12" fill="#404040">${escapeHtml(yLabel)}</text>`);
    groups.forEach((g, i) => {
      const s = g.stats, cx = L + slot * i + slot / 2, bw = Math.min(110, slot * 0.45);
      const x1 = cx - bw / 2, x2 = cx + bw / 2;
      parts.push(
        `<line x1="${cx}" x2="${cx}" y1="${y(s.whi)}" y2="${y(s.q3)}" stroke="#1F4E79" stroke-width="1.5"/>`,
        `<line x1="${cx}" x2="${cx}" y1="${y(s.q1)}" y2="${y(s.wlo)}" stroke="#1F4E79" stroke-width="1.5"/>`,
        `<line x1="${cx - bw / 4}" x2="${cx + bw / 4}" y1="${y(s.whi)}" y2="${y(s.whi)}" stroke="#1F4E79" stroke-width="1.5"/>`,
        `<line x1="${cx - bw / 4}" x2="${cx + bw / 4}" y1="${y(s.wlo)}" y2="${y(s.wlo)}" stroke="#1F4E79" stroke-width="1.5"/>`,
        `<rect x="${x1}" y="${y(s.q3)}" width="${bw}" height="${Math.max(1, y(s.q1) - y(s.q3))}" fill="#DCE8F4" stroke="#1F4E79" stroke-width="1.5"/>`,
        `<line x1="${x1}" x2="${x2}" y1="${y(s.med)}" y2="${y(s.med)}" stroke="#B00020" stroke-width="2.5"/>`);
      for (const o of s.out) {
        parts.push(`<circle cx="${cx}" cy="${y(o.v)}" r="4.5" fill="#fff" stroke="#B00020" stroke-width="2"/>`,
          `<text x="${cx + 9}" y="${y(o.v) + 4}" font-size="11" fill="#B00020">${escapeHtml(o.label)} (${E.display(o.v)})</text>`);
      }
      parts.push(`<text x="${cx}" y="${H - B + 20}" text-anchor="middle" font-size="12" fill="#404040">${escapeHtml(g.name)} (n = ${s.n})</text>`);
    });
    parts.push(`<line x1="${L}" x2="${W - R}" y1="${T + ih}" y2="${T + ih}" stroke="#404040"/></svg>`);
    return parts.join("");
  }
  function boxTable(groups) {
    const d = v => E.display(Math.round(v * 1e6) / 1e6);
    const rows = [["", ...groups.map(g => escapeHtml(g.name))],
      ["Ελάχιστο", ...groups.map(g => d(g.stats.min))],
      ["Q1 (1ο τεταρτημόριο)", ...groups.map(g => d(g.stats.q1))],
      ["Διάμεσος", ...groups.map(g => d(g.stats.med))],
      ["Q3 (3ο τεταρτημόριο)", ...groups.map(g => d(g.stats.q3))],
      ["Μέγιστο", ...groups.map(g => d(g.stats.max))],
      ["IQR = Q3 - Q1", ...groups.map(g => d(g.stats.iqr))],
      ["Όρια (Q1 - 1,5·IQR, Q3 + 1,5·IQR)", ...groups.map(g => d(g.stats.lo) + " έως " + d(g.stats.hi))],
      ["Ακραίες τιμές", ...groups.map(g => g.stats.out.length ? g.stats.out.map(o => escapeHtml(o.label) + " (" + d(o.v) + ")").join(", ") : "καμία")]];
    return "<table class='box-stats'>" + rows.map((r, i) => "<tr>" + r.map(c => (i === 0 ? `<th>${c}</th>` : `<td>${c}</td>`)).join("") + "</tr>").join("") + "</table>";
  }
  function categoryColumns(exclude) {
    const out = [];
    for (let c = 0; c < sheet.cols; c++) {
      if (c === exclude) continue;
      const seen = new Set();
      let text = 0;
      for (let r = 1; r < sheet.rows; r++) {
        const v = values[r][c];
        if (v === null || v === "") continue;
        if (typeof v === "string") text++;
        seen.add(E.display(v));
      }
      if (text && seen.size >= 2 && seen.size <= 8) out.push(c);
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
    q(".chart-group-label").hidden = kind !== "box";
    q(".chart-notes").innerHTML = "";
    if (kind === "box") {
      const select = q(".chart-group");
      const keep = select.value;
      const cats = categoryColumns(c);
      select.innerHTML = "<option value=''>Χωρίς ομάδες</option>" +
        cats.map(k => `<option value="${k}">${E.colName(k)} · ${escapeHtml(E.display(values[0][k]) || "")}</option>`).join("");
      if (cats.map(String).includes(keep)) select.value = keep;
      const g = select.value === "" ? null : Number(select.value);
      const buckets = new Map();
      for (let r = 1; r < sheet.rows; r++) {
        const v = values[r][c];
        if (typeof v !== "number") continue;
        const name = g === null ? (E.display(values[0][c]) || E.colName(c)) : (E.display(values[r][g]) || "(κενό)");
        if (!buckets.has(name)) buckets.set(name, []);
        buckets.get(name).push({ v, label: E.display(values[r][0]) || String(r + 1) });
      }
      const groups = [...buckets.entries()].filter(([, pts]) => pts.length).map(([name, points]) => ({ name, points, stats: boxStats(points) }));
      if (!groups.length) { q(".chart-area").innerHTML = ""; return; }
      const label = E.display(values[0][c]) || E.colName(c);
      q(".chart-area").innerHTML = svgBox(label + (g === null ? "" : " ανά " + (E.display(values[0][g]) || E.colName(g))), groups, label);
      q(".chart-notes").innerHTML = boxTable(groups) +
        "<p class='box-note'>Τεταρτημόρια όπως η <code>QUARTILE.INC</code> του Excel. Ακραία είναι η τιμή έξω από τα όρια· σημαίνει «έλεγξέ με», όχι «διάγραψέ με».</p>";
      return;
    }
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
  q(".chart-group").addEventListener("change", drawChart);
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
  // The column letters and row numbers select a whole column or row, Shift
  // widens it, and the corner selects everything, as in a spreadsheet.
  function headerOf(target) {
    const th = target.closest("th");
    if (!th) return null;
    if (th.dataset.col !== undefined) return { col: Number(th.dataset.col) };
    if (th.dataset.row !== undefined) return { row: Number(th.dataset.row) };
    if (th.classList.contains("corner")) return { all: true };
    return null;
  }
  function selectHeader(h, extend) {
    const lastR = sheet.rows - 1, lastC = sheet.cols - 1;
    if (h.all) sel = { r: 0, c: 0, r2: lastR, c2: lastC };
    else if (h.col !== undefined) {
      if (extend && sel.r === 0 && sel.r2 === lastR) sel.c2 = h.col;
      else sel = { r: 0, c: h.col, r2: lastR, c2: h.col };
    } else {
      if (extend && sel.c === 0 && sel.c2 === lastC) sel.r2 = h.row;
      else sel = { r: h.row, c: 0, r2: h.row, c2: lastC };
    }
    render();
  }
  table.addEventListener("mousedown", e => {
    const h = headerOf(e.target);
    if (!h || pointing()) return;
    if (editor) commit();
    e.preventDefault();
    const whole = h.col !== undefined ? sel.r === 0 && sel.r2 === sheet.rows - 1 && h.col >= Math.min(sel.c, sel.c2) && h.col <= Math.max(sel.c, sel.c2)
      : h.row !== undefined ? sel.c === 0 && sel.c2 === sheet.cols - 1 && h.row >= Math.min(sel.r, sel.r2) && h.row <= Math.max(sel.r, sel.r2) : false;
    // a right click inside a selected column or row keeps it, for the menu
    if (e.button === 2 && whole) return;
    selectHeader(h, e.shiftKey);
    if (e.button === 0) { dragging = h.col !== undefined ? "cols" : h.row !== undefined ? "rows" : false; downAt = null; }
    grid.focus({ preventScroll: true });
  });
  table.addEventListener("mouseover", e => {
    if (!(e.buttons & 1) || (dragging !== "cols" && dragging !== "rows")) return;
    const h = headerOf(e.target);
    if (dragging === "cols" && h && h.col !== undefined && sel.c2 !== h.col) { sel.c2 = h.col; render(); }
    if (dragging === "rows" && h && h.row !== undefined && sel.r2 !== h.row) { sel.r2 = h.row; render(); }
  });

  table.addEventListener("mousedown", e => {
    if (headerOf(e.target)) return;
    if (e.target.classList.contains("fill-handle")) {
      e.preventDefault();
      e.stopPropagation();
      if (editor) commit();
      dragging = "fill";
      fillTo = { r: box().r2, c: box().c2 };
      return;
    }
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
    if (e.button !== 0) return;  // the right button opens the menu, see below
    e.preventDefault();
    if (e.shiftKey) { sel.r2 = r; sel.c2 = c; } else sel = { r, c, r2: r, c2: c };
    dragging = "select";
    downAt = { x: e.clientX, y: e.clientY };
    render();
    grid.focus({ preventScroll: true });
  });
  table.addEventListener("mouseover", e => {
    if (!dragging || !(e.buttons & 1)) return;
    if (dragging === "select" && downAt && Math.hypot(e.clientX - downAt.x, e.clientY - downAt.y) < 4) return;
    const cell = e.target.closest("td");
    if (!cell) return;
    const r = Number(cell.dataset.r), c = Number(cell.dataset.c);
    if (dragging === "pick" && editor) { pick(r, c, true); return; }
    if (dragging === "fill") {
      const b = box();
      // one direction at a time, whichever the pointer has gone further in
      const down = r - b.r2, across = c - b.c2;
      const next = down >= across ? { r: Math.max(b.r2, r), c: b.c2 } : { r: b.r2, c: Math.max(b.c2, c) };
      if (!fillTo || next.r !== fillTo.r || next.c !== fillTo.c) { fillTo = next; render(); }
      return;
    }
    if (dragging === "select" && (sel.r2 !== r || sel.c2 !== c)) { sel.r2 = r; sel.c2 = c; render(); }
  });
  document.addEventListener("mouseup", () => {
    if (dragging === "fill") {
      const to = fillTo;
      const b = box();
      fillTo = null;
      dragging = false;
      // a click on the handle that went nowhere leaves the grid as it is, so
      // the second click of a double click lands on the same handle
      if (!to || (to.r === b.r2 && to.c === b.c2)) return;
      dragFill(b, to.r, to.c);
      render();
      grid.focus({ preventScroll: true });
      return;
    }
    if (dragging === "pick" && editor) editor.input.focus();
    dragging = false;
  });
  table.addEventListener("dblclick", e => {
    if (e.target.classList.contains("fill-handle")) {
      // down to the last row holding data, as a double click on Excel's handle does
      const b = box();
      const last = lastFilledRow();
      if (last > b.r2) { dragFill(b, last, b.c2); render(); }
      return;
    }
    if (e.target.closest("td") && !editor) startEdit();
  });

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
  // Undo answers anywhere in the panel, a toolbar button just clicked included,
  // except in a text box, where Ctrl+Z belongs to the text being typed. By
  // physical key as well, so a Greek layout (ζ, υ) undoes like a Latin one.
  dock.addEventListener("keydown", e => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.target.tagName === "INPUT") return;
    const z = e.code === "KeyZ" || e.key === "z" || e.key === "Z";
    const y = e.code === "KeyY" || e.key === "y" || e.key === "Y";
    if (z && !e.shiftKey) { e.preventDefault(); step(undo, redo, "undo"); }
    else if (y || (z && e.shiftKey)) { e.preventDefault(); step(redo, undo, "redo"); }
  });
  grid.addEventListener("copy", e => { if (!editor) copy(e); });
  grid.addEventListener("paste", e => { if (!editor) paste(e); });
  grid.addEventListener("cut", e => { if (!editor) { copy(e); clearSelection(); } });

  // ---- the right-click menu ----------------------------------------------
  const menu = document.createElement("div");
  menu.className = "sheet-menu";
  menu.setAttribute("role", "menu");
  menu.hidden = true;
  const ITEMS = [
    ["undo", "Αναίρεση", "Ctrl+Z"], ["redo", "Επανάληψη", "Ctrl+Y"], null,
    ["cut", "Αποκοπή", "Ctrl+X"], ["copy", "Αντιγραφή", "Ctrl+C"], ["paste", "Επικόλληση", "Ctrl+V"], null,
    ["rowAbove", "Εισαγωγή γραμμής από πάνω"], ["rowBelow", "Εισαγωγή γραμμής από κάτω"],
    ["colLeft", "Εισαγωγή στήλης αριστερά"], ["colRight", "Εισαγωγή στήλης δεξιά"], null,
    ["delRow", "Διαγραφή γραμμής"], ["delCol", "Διαγραφή στήλης"], ["clear", "Καθαρισμός περιεχομένων", "Delete"], null,
    ["fill", "Συμπλήρωση προς τα κάτω", "Ctrl+D"], ["chart", "Γράφημα αυτής της στήλης"],
  ];
  menu.innerHTML = ITEMS.map(it => (it
    ? `<button type="button" role="menuitem" data-menu="${it[0]}"><span>${it[1]}</span><kbd>${it[2] || ""}</kbd></button>`
    : "<hr>")).join("");
  dock.appendChild(menu);
  function closeMenu() { menu.hidden = true; }
  table.addEventListener("contextmenu", e => {
    const cell = e.target.closest("td");
    const head = headerOf(e.target);
    if (!cell && !head) return;
    e.preventDefault();
    if (editor) commit();
    if (cell) {
      const r = Number(cell.dataset.r), c = Number(cell.dataset.c);
      if (!inSel(r, c)) { sel = { r, c, r2: r, c2: c }; render(); }
    }
    menu.querySelector("[data-menu=paste]").disabled = !clip && !(navigator.clipboard && navigator.clipboard.readText);
    menu.querySelector("[data-menu=undo]").disabled = !undo.length;
    menu.querySelector("[data-menu=redo]").disabled = !redo.length;
    menu.hidden = false;
    const w = menu.offsetWidth, h = menu.offsetHeight;
    menu.style.left = Math.min(e.clientX, window.innerWidth - w - 4) + "px";
    menu.style.top = Math.max(4, Math.min(e.clientY, window.innerHeight - h - 4)) + "px";
    menu.querySelector("button").focus();
  });
  menu.addEventListener("click", async e => {
    const btn = e.target.closest("button[data-menu]");
    if (!btn) return;
    closeMenu();
    const b = box();
    switch (btn.dataset.menu) {
      case "undo": step(undo, redo, "undo"); break;
      case "redo": step(redo, undo, "redo"); break;
      case "cut": copy(); clearSelection(); break;
      case "copy": copy(); break;
      case "paste": {
        let t = clip && clip.text;
        try { if (navigator.clipboard && navigator.clipboard.readText) t = await navigator.clipboard.readText(); }
        catch (refused) { /* no clipboard permission: use what this sheet copied */ }
        if (t) pasteText(t); else say("Δεν υπάρχει κάτι για επικόλληση.");
        break;
      }
      case "rowAbove": insertRow(b.r1); break;
      case "rowBelow": insertRow(b.r2 + 1); break;
      case "colLeft": insertColumn(b.c1); break;
      case "colRight": insertColumn(b.c2 + 1); break;
      case "delRow": deleteRows(); break;
      case "delCol": deleteColumns(); break;
      case "clear": clearSelection(); break;
      case "fill": fillDown(); break;
      case "chart": openChart(); break;
    }
    grid.focus({ preventScroll: true });
  });
  menu.addEventListener("keydown", e => {
    const items = Array.from(menu.querySelectorAll("button:not(:disabled)"));
    const i = items.indexOf(document.activeElement);
    if (e.key === "Escape") { e.preventDefault(); closeMenu(); grid.focus(); }
    else if (e.key === "ArrowDown") { e.preventDefault(); items[(i + 1) % items.length].focus(); }
    else if (e.key === "ArrowUp") { e.preventDefault(); items[(i - 1 + items.length) % items.length].focus(); }
  });
  document.addEventListener("mousedown", e => { if (!menu.hidden && !menu.contains(e.target)) closeMenu(); });
  grid.addEventListener("scroll", closeMenu);
  window.addEventListener("blur", closeMenu);

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
      case "undo": step(undo, redo, "undo"); break;
      case "redo": step(redo, undo, "redo"); break;
      case "fill": fillDown(); break;
      case "addcol": insertColumn(box().c2 + 1); break;
      case "addrow": insertRow(box().r2 + 1); break;
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
        if (!confirm("Επαναφορά στο αρχικό αρχείο; Το ιστορικό των αλλαγών θα χαθεί (η Αναίρεση φέρνει πίσω τα δεδομένα).")) break;
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
      saved = JSON.stringify(sheet.cells);
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
