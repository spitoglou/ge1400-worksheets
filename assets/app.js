// Fill the real worksheet PDF from the page and hand it back as a reader would
// save it: the same file, its fields carrying /V, appearances drawn in a font
// that has Greek, nothing flattened. The field names on the page are the PDF's
// own (tools/teach/webforms.py checks that at build time).
(function () {
  "use strict";
  const cfg = window.WORKSHEET;
  const form = document.getElementById("worksheet");
  const status = document.getElementById("status");
  const key = "worksheet:" + cfg.slug + ":lab" + cfg.lab;
  const FONT_URL = "assets/NotoSans-Regular.ttf";

  const inputs = () => Array.from(form.querySelectorAll("input[name], textarea[name]"));

  function say(text, kind) {
    status.textContent = text;
    status.className = kind || "";
  }

  // Links in a worksheet point at browsers and tools the lab uses; keep the
  // student's answers on screen while they look.
  form.querySelectorAll("a[href^='http']").forEach(a => {
    a.target = "_blank";
    a.rel = "noopener";
  });

  // Drafts live in this browser only. Storage can be missing or refused
  // (private windows, blocked site data), and the page must work without it.
  function collect() {
    const values = {};
    for (const el of inputs()) {
      if (el.type === "checkbox") { if (el.checked) values[el.name] = true; }
      else if (el.value) values[el.name] = el.value;
    }
    return values;
  }

  function apply(values) {
    for (const el of inputs()) {
      const v = values[el.name];
      if (el.type === "checkbox") el.checked = v === true;
      else el.value = typeof v === "string" ? v : "";
    }
  }

  function store() {
    try {
      localStorage.setItem(key, JSON.stringify({ version: cfg.version, values: collect() }));
      return true;
    } catch (e) {
      return false;
    }
  }

  let timer = null;
  form.addEventListener("input", () => {
    clearTimeout(timer);
    timer = setTimeout(() => say(store() ? "Αποθηκεύτηκε σε αυτόν τον browser." : ""), 400);
  });

  try {
    const saved = JSON.parse(localStorage.getItem(key) || "null");
    if (saved && saved.values) {
      apply(saved.values);
      if (Object.keys(saved.values).length) say("Οι απαντήσεις σας επανήλθαν από αυτόν τον browser.");
    }
  } catch (e) { /* no storage: start empty */ }

  async function fetchBytes(url) {
    const r = await fetch(url, { cache: "no-cache" });
    if (!r.ok) throw new Error(url + ": " + r.status);
    return r.arrayBuffer();
  }

  function download(bytes, name) {
    const url = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  const exportBtn = document.getElementById("export");
  exportBtn.addEventListener("click", async () => {
    const values = collect();
    // A blank upload is the failure this page exists to prevent.
    if (!Object.keys(values).length) {
      say("Δεν έχετε συμπληρώσει κανένα πεδίο. Το PDF δεν δημιουργήθηκε.", "error");
      return;
    }
    exportBtn.disabled = true;
    say("Δημιουργία PDF…");
    try {
      const { PDFDocument } = PDFLib;
      const [pdfBytes, fontBytes] = await Promise.all([fetchBytes(cfg.pdf), fetchBytes(FONT_URL)]);
      const doc = await PDFDocument.load(pdfBytes);
      doc.registerFontkit(fontkit);
      const font = await doc.embedFont(fontBytes, { subset: true });
      const pdfForm = doc.getForm();
      for (const [name, v] of Object.entries(values)) {
        if (v === true) pdfForm.getCheckBox(name).check();
        else pdfForm.getTextField(name).setText(v);
      }
      // Drawn here in a Greek font. Left to save(), pdf-lib would draw them in
      // Helvetica, which cannot encode a single Greek letter.
      pdfForm.updateFieldAppearances(font);
      const out = await doc.save({ updateFieldAppearances: false });
      download(out, cfg.filename);
      say("Το PDF κατέβηκε: " + cfg.filename + ". Ανεβάστε το στο e-class.", "ok");
    } catch (e) {
      console.error(e);
      say("Το PDF δεν δημιουργήθηκε (" + e.message + "). Δοκιμάστε ξανά ή ενημερώστε τον διδάσκοντα.", "error");
    } finally {
      exportBtn.disabled = false;
    }
  });

  document.getElementById("import").addEventListener("change", async (ev) => {
    const file = ev.target.files[0];
    ev.target.value = "";
    if (!file) return;
    try {
      const { PDFDocument, PDFTextField, PDFCheckBox } = PDFLib;
      const doc = await PDFDocument.load(await file.arrayBuffer());
      const byName = new Map(doc.getForm().getFields().map(f => [f.getName(), f]));
      const ours = inputs().filter(el => byName.has(el.name));
      if (!ours.length || ours.length !== inputs().length) {
        say("Αυτό το PDF δεν είναι το φύλλο αυτού του εργαστηρίου (ή είναι ισοπεδωμένο). Δεν φορτώθηκε τίποτα.", "error");
        return;
      }
      const values = {};
      for (const el of ours) {
        const f = byName.get(el.name);
        if (f instanceof PDFCheckBox) { if (f.isChecked()) values[el.name] = true; }
        else if (f instanceof PDFTextField) { const t = f.getText(); if (t) values[el.name] = t; }
      }
      apply(values);
      store();
      say("Φορτώθηκαν " + Object.keys(values).length + " απαντήσεις από το " + file.name + ".", "ok");
    } catch (e) {
      console.error(e);
      say("Το αρχείο δεν διαβάστηκε ως PDF (" + e.message + ").", "error");
    }
  });

  // Two presses, so a stray click cannot wipe an evening's work.
  const clearBtn = document.getElementById("clear");
  let armed = null;
  clearBtn.addEventListener("click", () => {
    if (!armed) {
      clearBtn.textContent = "Σίγουρα; Πατήστε ξανά";
      armed = setTimeout(() => { armed = null; clearBtn.textContent = "Καθαρισμός"; }, 4000);
      return;
    }
    clearTimeout(armed);
    armed = null;
    clearBtn.textContent = "Καθαρισμός";
    apply({});
    try { localStorage.removeItem(key); } catch (e) { /* nothing stored */ }
    say("Όλα τα πεδία καθαρίστηκαν.");
  });
})();
