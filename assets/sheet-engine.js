// The formula engine behind a lab's spreadsheet panel (sheet.js): cells hold
// what the student typed, and a cell that starts with "=" is evaluated with
// the spreadsheet syntax the worksheets teach (=AVERAGE(D2:D17), =D2/(E2/100)^2).
//
// Written here rather than taken from a library because the lab needs a small,
// exact subset that works offline on any lab terminal, and because two things
// the libraries do not do matter in this room: a formula typed with the Greek
// keyboard layout still on (=ΑΩΕΡΑΓΕ(Δ2:Δ17)) is read as the Latin it was meant
// to be, and ";" is accepted as the argument separator a Greek Excel uses.
(function (root) {
  "use strict";

  const BOM = String.fromCharCode(0xFEFF);

  // Greek keyboard layout to the Latin key in the same place, for formula text
  // outside quotes. Covers both the look-alike capitals and the ones that only
  // share a key (Δ is D, Ω is V, Ψ is C).
  const LAYOUT = {
    "Α": "A", "Β": "B", "Ψ": "C", "Δ": "D", "Ε": "E", "Φ": "F", "Γ": "G", "Η": "H",
    "Ι": "I", "Ξ": "J", "Κ": "K", "Λ": "L", "Μ": "M", "Ν": "N", "Ο": "O", "Π": "P",
    "Ρ": "R", "Σ": "S", "Τ": "T", "Θ": "U", "Ω": "V", "Χ": "X", "Υ": "Y", "Ζ": "Z",
    "α": "a", "β": "b", "ψ": "c", "δ": "d", "ε": "e", "φ": "f", "γ": "g", "η": "h",
    "ι": "i", "ξ": "j", "κ": "k", "λ": "l", "μ": "m", "ν": "n", "ο": "o", "π": "p",
    "ρ": "r", "σ": "s", "ς": "w", "τ": "t", "θ": "u", "ω": "v", "χ": "x", "υ": "y",
    "ζ": "z", "ά": "a", "έ": "e", "ή": "h", "ί": "i", "ό": "o", "ύ": "y", "ώ": "v",
  };

  // Text compared the way a student reads it: case folded, and the Greek
  // capitals that look like Latin ones taken as those letters, so "A" finds
  // the sex column's Greek «Α».
  const LOOKALIKE = { "α": "a", "β": "b", "ε": "e", "ζ": "z", "η": "h", "ι": "i", "κ": "k",
                      "μ": "m", "ν": "n", "ο": "o", "ρ": "p", "τ": "t", "υ": "y", "χ": "x" };
  const fold = t => Array.from(String(t).toLowerCase(), ch => LOOKALIKE[ch] || ch).join("");

  class SheetError {
    constructor(code) { this.code = code; }
    toString() { return this.code; }
  }
  const err = code => new SheetError(code);
  const isErr = v => v instanceof SheetError;

  function colName(c) {
    let s = "";
    c += 1;
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  function colIndex(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  const cellName = (r, c) => colName(c) + (r + 1);

  // Formula text outside quoted strings, with the Greek layout mapped to Latin.
  function normalise(text) {
    let out = "";
    let quoted = false;
    for (const ch of text) {
      if (ch === '"') quoted = !quoted;
      out += quoted || ch === '"' ? ch : (LAYOUT[ch] || ch);
    }
    return out;
  }

  // A literal a student typed: a number when it reads as one (with "." or ","
  // as the decimal mark), otherwise the text as typed.
  function literal(raw) {
    const t = String(raw).trim();
    if (t === "") return null;
    if (/^[-+]?(\d+([.,]\d*)?|[.,]\d+)([eE][-+]?\d+)?$/.test(t)) return Number(t.replace(",", "."));
    const up = t.toUpperCase();
    if (up === "TRUE") return true;
    if (up === "FALSE") return false;
    return String(raw);
  }

  // ---- tokens and parser -------------------------------------------------
  function tokenise(src) {
    const toks = [];
    let i = 0;
    while (i < src.length) {
      const ch = src[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"') {
        let j = i + 1, s = "";
        while (j < src.length) {
          if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
          s += src[j++];
        }
        if (j >= src.length) throw err("#VALUE!");
        toks.push({ t: "str", v: s });
        i = j + 1;
        continue;
      }
      if (src.startsWith("#REF!", i)) throw err("#REF!");
      const num = /^(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?/.exec(src.slice(i));
      if (num) { toks.push({ t: "num", v: Number(num[0]) }); i += num[0].length; continue; }
      const word = /^\$?[A-Za-z][A-Za-z0-9.]*\$?\d*/.exec(src.slice(i));
      if (word) {
        const w = word[0];
        const ref = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(w);
        if (ref && src[i + w.length] !== "(") {
          toks.push({ t: "ref", c: colIndex(ref[2]), r: Number(ref[4]) - 1 });
        } else {
          toks.push({ t: "name", v: w.replace(/\$/g, "").toUpperCase() });
        }
        i += w.length;
        continue;
      }
      const two = src.slice(i, i + 2);
      if (two === "<>" || two === "<=" || two === ">=") { toks.push({ t: "op", v: two }); i += 2; continue; }
      if ("+-*/^&=<>%(),;:".includes(ch)) { toks.push({ t: "op", v: ch }); i++; continue; }
      throw err("#NAME?");
    }
    return toks;
  }

  function parse(src) {
    const toks = tokenise(src);
    let p = 0;
    const peek = () => toks[p];
    const isOp = v => toks[p] && toks[p].t === "op" && toks[p].v === v;
    const take = v => { if (!isOp(v)) throw err("#VALUE!"); p++; };

    function compare() {
      let left = concat();
      while (peek() && peek().t === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(peek().v)) {
        const op = toks[p++].v;
        left = { k: "bin", op, a: left, b: concat() };
      }
      return left;
    }
    function concat() {
      let left = additive();
      while (isOp("&")) { p++; left = { k: "bin", op: "&", a: left, b: additive() }; }
      return left;
    }
    function additive() {
      let left = term();
      while (isOp("+") || isOp("-")) { const op = toks[p++].v; left = { k: "bin", op, a: left, b: term() }; }
      return left;
    }
    function term() {
      let left = power();
      while (isOp("*") || isOp("/")) { const op = toks[p++].v; left = { k: "bin", op, a: left, b: power() }; }
      return left;
    }
    function power() {
      let left = unary();
      while (isOp("^")) { p++; left = { k: "bin", op: "^", a: left, b: unary() }; }
      return left;
    }
    function unary() {
      if (isOp("-")) { p++; return { k: "neg", a: unary() }; }
      if (isOp("+")) { p++; return unary(); }
      let node = primary();
      while (isOp("%")) { p++; node = { k: "pct", a: node }; }
      return node;
    }
    function primary() {
      const tok = toks[p];
      if (!tok) throw err("#VALUE!");
      if (tok.t === "num") { p++; return { k: "lit", v: tok.v }; }
      if (tok.t === "str") { p++; return { k: "lit", v: tok.v }; }
      if (tok.t === "ref") {
        p++;
        if (isOp(":")) {
          p++;
          const end = toks[p];
          if (!end || end.t !== "ref") throw err("#REF!");
          p++;
          return { k: "range", r1: Math.min(tok.r, end.r), c1: Math.min(tok.c, end.c),
                   r2: Math.max(tok.r, end.r), c2: Math.max(tok.c, end.c) };
        }
        return { k: "ref", r: tok.r, c: tok.c };
      }
      if (tok.t === "name") {
        p++;
        if (isOp("(")) {
          p++;
          const args = [];
          if (!isOp(")")) {
            for (;;) {
              if (isOp(",") || isOp(";") || isOp(")")) args.push({ k: "lit", v: null });
              else args.push(compare());
              if (isOp(",") || isOp(";")) { p++; continue; }
              break;
            }
          }
          take(")");
          return { k: "call", name: tok.v, args };
        }
        if (tok.v === "TRUE") return { k: "lit", v: true };
        if (tok.v === "FALSE") return { k: "lit", v: false };
        if (tok.v.startsWith("#")) throw err(tok.v);
        throw err("#NAME?");
      }
      if (isOp("(")) { p++; const e = compare(); take(")"); return e; }
      throw err("#VALUE!");
    }

    const tree = compare();
    if (p !== toks.length) throw err("#VALUE!");
    return tree;
  }

  // ---- values ------------------------------------------------------------
  function toNumber(v) {
    if (isErr(v)) throw v;
    if (v === null || v === "") return 0;
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v ? 1 : 0;
    const n = literal(v);
    if (typeof n === "number") return n;
    throw err("#VALUE!");
  }
  const toText = v => (v === null ? "" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v));
  function toBool(v) {
    if (isErr(v)) throw v;
    if (typeof v === "boolean") return v;
    if (v === null) return false;
    if (typeof v === "number") return v !== 0;
    const up = String(v).toUpperCase();
    if (up === "TRUE") return true;
    if (up === "FALSE") return false;
    throw err("#VALUE!");
  }

  // Numbers taken from a function's arguments the way a spreadsheet takes them:
  // from a range only real numbers count (text and blanks are skipped), while a
  // value typed straight into the call is converted.
  function numbers(args) {
    const out = [];
    for (const a of args) {
      if (a && a.range) {
        for (const v of a.values) {
          if (isErr(v)) throw v;
          if (typeof v === "number") out.push(v);
        }
      } else {
        if (isErr(a)) throw a;
        if (a === null) continue;
        out.push(toNumber(a));
      }
    }
    return out;
  }
  const flat = a => (a && a.range ? a.values : [a]);

  function criterion(c) {
    if (typeof c === "number") return v => typeof v === "number" && v === c;
    if (typeof c === "boolean") return v => v === c;
    const text = toText(c);
    const m = /^(<=|>=|<>|<|>|=)?(.*)$/.exec(text);
    const op = m[1] || "=";
    const rhs = literal(m[2]);
    if (typeof rhs === "number") {
      return v => {
        if (typeof v !== "number") return op === "<>";
        switch (op) {
          case "=": return v === rhs; case "<>": return v !== rhs;
          case "<": return v < rhs; case ">": return v > rhs;
          case "<=": return v <= rhs; default: return v >= rhs;
        }
      };
    }
    const target = fold(m[2] || "");
    if (op === "=" && target === "") return v => v === null || v === "";
    if (op === "<>" && target === "") return v => !(v === null || v === "");
    const pattern = new RegExp("^" + target.replace(/[.+^${}()|[\]\\]/g, "\\$&")
      .replace(/\*/g, ".*").replace(/\?/g, ".") + "$");
    return v => {
      const s = fold(toText(v));
      if (op === "=") return pattern.test(s);
      if (op === "<>") return !pattern.test(s);
      return false;
    };
  }

  function median(xs) {
    const s = xs.slice().sort((a, b) => a - b);
    const n = s.length;
    return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2;
  }
  // Excel's QUARTILE.INC and PERCENTILE.INC: rank (n - 1)p, interpolated.
  function quantile(xs, p) {
    if (!xs.length) throw err("#NUM!");
    if (p < 0 || p > 1) throw err("#NUM!");
    const s = xs.slice().sort((a, b) => a - b);
    const h = (s.length - 1) * p, lo = Math.floor(h);
    return lo + 1 < s.length ? s[lo] + (h - lo) * (s[lo + 1] - s[lo]) : s[lo];
  }
  function variance(xs, sample) {
    const n = xs.length;
    if (n < (sample ? 2 : 1)) throw err("#DIV/0!");
    const m = xs.reduce((a, b) => a + b, 0) / n;
    return xs.reduce((a, b) => a + (b - m) * (b - m), 0) / (sample ? n - 1 : n);
  }
  function ifs(args, reduce) {
    const range = args[0];
    const test = criterion(args[1]);
    const over = args[2] && args[2].range ? args[2] : range;
    const picked = [];
    range.values.forEach((v, i) => {
      if (isErr(v)) throw v;
      if (test(v)) { const w = over.values[i]; if (isErr(w)) throw w; if (typeof w === "number") picked.push(w); }
    });
    return reduce(picked);
  }
  const needRange = a => { if (!(a && a.range)) throw err("#VALUE!"); };
  const mean = xs => { if (!xs.length) throw err("#DIV/0!"); return xs.reduce((x, y) => x + y, 0) / xs.length; };

  const FUNCTIONS = {
    SUM: a => numbers(a).reduce((x, y) => x + y, 0),
    AVERAGE: a => mean(numbers(a)),
    MEDIAN: a => { const xs = numbers(a); if (!xs.length) throw err("#NUM!"); return median(xs); },
    QUARTILE: a => { const q = Math.trunc(toNumber(a[1])); if (q < 0 || q > 4) throw err("#NUM!"); return quantile(numbers([a[0]]), q / 4); },
    "QUARTILE.INC": a => { const q = Math.trunc(toNumber(a[1])); if (q < 0 || q > 4) throw err("#NUM!"); return quantile(numbers([a[0]]), q / 4); },
    PERCENTILE: a => quantile(numbers([a[0]]), toNumber(a[1])),
    "PERCENTILE.INC": a => quantile(numbers([a[0]]), toNumber(a[1])),
    MIN: a => { const xs = numbers(a); return xs.length ? Math.min(...xs) : 0; },
    MAX: a => { const xs = numbers(a); return xs.length ? Math.max(...xs) : 0; },
    STDEV: a => Math.sqrt(variance(numbers(a), true)),
    "STDEV.S": a => Math.sqrt(variance(numbers(a), true)),
    "STDEV.P": a => Math.sqrt(variance(numbers(a), false)),
    STDEVP: a => Math.sqrt(variance(numbers(a), false)),
    VAR: a => variance(numbers(a), true),
    "VAR.S": a => variance(numbers(a), true),
    "VAR.P": a => variance(numbers(a), false),
    COUNT: a => a.flatMap(flat).filter(v => typeof v === "number").length,
    COUNTA: a => a.flatMap(flat).filter(v => v !== null && v !== "").length,
    COUNTBLANK: a => a.flatMap(flat).filter(v => v === null || v === "").length,
    COUNTIF: a => { needRange(a[0]); const t = criterion(a[1]); return a[0].values.filter(v => !isErr(v) && t(v)).length; },
    SUMIF: a => { needRange(a[0]); return ifs(a, xs => xs.reduce((x, y) => x + y, 0)); },
    AVERAGEIF: a => { needRange(a[0]); return ifs(a, mean); },
    ROUND: a => { const x = toNumber(a[0]), d = a.length > 1 ? toNumber(a[1]) : 0; const f = Math.pow(10, d); return Math.round((x + Math.sign(x) * 1e-12) * f) / f; },
    ABS: a => Math.abs(toNumber(a[0])),
    SQRT: a => { const x = toNumber(a[0]); if (x < 0) throw err("#NUM!"); return Math.sqrt(x); },
    POWER: a => Math.pow(toNumber(a[0]), toNumber(a[1])),
    IF: () => { throw err("#VALUE!"); },  // evaluated lazily in Sheet.evaluate
    AND: a => a.flatMap(flat).every(v => toBool(v)),
    OR: a => a.flatMap(flat).some(v => toBool(v)),
    NOT: a => !toBool(a[0]),
    ISBLANK: a => a[0] === null || a[0] === "",
    ISNUMBER: a => typeof a[0] === "number",
  };
  const FUNCTION_NAMES = Object.keys(FUNCTIONS).sort();

  function compareValues(op, a, b) {
    let x = a, y = b;
    if (typeof x === "string" || typeof y === "string") { x = fold(toText(x)); y = fold(toText(y)); }
    else { x = toNumber(x); y = toNumber(y); }
    switch (op) {
      case "=": return x === y; case "<>": return x !== y;
      case "<": return x < y; case ">": return x > y;
      case "<=": return x <= y; default: return x >= y;
    }
  }

  // ---- the sheet ---------------------------------------------------------
  class Sheet {
    constructor(rows) {
      this.cells = rows.map(r => r.map(v => String(v)));
      this.normalise();
    }
    get rows() { return this.cells.length; }
    get cols() { return this.cells.reduce((m, r) => Math.max(m, r.length), 0); }
    normalise() {
      const w = this.cols;
      for (const r of this.cells) while (r.length < w) r.push("");
    }
    raw(r, c) { return (this.cells[r] && this.cells[r][c]) || ""; }
    set(r, c, text) {
      while (this.cells.length <= r) this.cells.push([]);
      while (this.cells[r].length <= c) this.cells[r].push("");
      this.cells[r][c] = text;
      this.normalise();
    }

    // Every cell's value, computed once per call. A formula that reaches
    // itself is #CIRC!, and any other failure is the spreadsheet error a
    // student would recognise.
    evaluate() {
      const memo = new Map();
      const busy = new Set();
      const scalar = v => (v && v.range ? (v.values.length === 1 ? v.values[0] : err("#VALUE!")) : v);
      const value = (r, c) => {
        if (r < 0 || c < 0) return err("#REF!");
        const key = r * 100000 + c;
        if (memo.has(key)) return memo.get(key);
        const raw = this.raw(r, c);
        let v;
        if (!raw.startsWith("=") || raw.length === 1) v = literal(raw);
        else if (busy.has(key)) return err("#CIRC!");
        else {
          busy.add(key);
          try {
            v = scalar(run(parse(normalise(raw.slice(1)))));
            if (typeof v === "number" && !isFinite(v)) v = err("#NUM!");
          } catch (e) {
            v = isErr(e) ? e : err("#VALUE!");
          }
          busy.delete(key);
        }
        memo.set(key, v);
        return v;
      };
      const run = node => {
        switch (node.k) {
          case "lit": return node.v;
          case "ref": return value(node.r, node.c);
          case "range": {
            const values = [];
            for (let r = node.r1; r <= node.r2; r++) for (let c = node.c1; c <= node.c2; c++) values.push(value(r, c));
            return { range: true, values };
          }
          case "neg": return -toNumber(scalar(run(node.a)));
          case "pct": return toNumber(scalar(run(node.a))) / 100;
          case "bin": {
            const a = scalar(run(node.a)), b = scalar(run(node.b));
            if (isErr(a)) throw a;
            if (isErr(b)) throw b;
            switch (node.op) {
              case "+": return toNumber(a) + toNumber(b);
              case "-": return toNumber(a) - toNumber(b);
              case "*": return toNumber(a) * toNumber(b);
              case "/": { const d = toNumber(b); if (d === 0) throw err("#DIV/0!"); return toNumber(a) / d; }
              case "^": return Math.pow(toNumber(a), toNumber(b));
              case "&": return toText(a) + toText(b);
              default: return compareValues(node.op, a, b);
            }
          }
          case "call": {
            const fn = FUNCTIONS[node.name];
            if (!fn) throw err("#NAME?");
            if (node.name === "IF") {
              const cond = toBool(scalar(run(node.args[0])));
              const branch = cond ? node.args[1] : node.args[2];
              if (!branch) return cond;
              const out = scalar(run(branch));
              if (isErr(out)) throw out;
              return out;
            }
            return fn(node.args.map(run));
          }
        }
        throw err("#VALUE!");
      };
      const grid = [];
      for (let r = 0; r < this.rows; r++) {
        const row = [];
        for (let c = 0; c < this.cols; c++) row.push(value(r, c));
        grid.push(row);
      }
      return grid;
    }

    // Rows and columns in and out, with every formula's references moved the
    // way a spreadsheet moves them: cells below or right of the change follow
    // it, a range that spans it grows or shrinks, a deleted cell is #REF!.
    remap(fn) {
      this.cells = this.cells.map(row => row.map(t => (t.startsWith("=") ? rewriteRefs(t, fn) : t)));
    }
    deleteRow(r) {
      this.cells.splice(r, 1);
      this.remap(ref => deleted(ref, r, "r"));
    }
    deleteColumn(c) {
      for (const row of this.cells) row.splice(c, 1);
      this.remap(ref => deleted(ref, c, "c"));
    }
    insertRow(at) {
      this.cells.splice(at, 0, new Array(this.cols).fill(""));
      this.remap(ref => inserted(ref, at, "r"));
    }
    insertColumn(at) {
      for (const row of this.cells) row.splice(at, 0, "");
      this.remap(ref => inserted(ref, at, "c"));
    }
    addRow() { this.insertRow(this.rows); }
    addColumn() { this.insertColumn(this.cols); }
  }

  // ---- moving formulas ---------------------------------------------------
  // Every reference in a formula, outside quotes, passed through `fn`, which
  // returns the replacement text. A range arrives as one {start, end} pair so
  // a deleted row can shrink it rather than break it.
  const REF_TEXT = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?/y;
  function rewriteRefs(formula, fn) {
    const src = "=" + normalise(formula.slice(1));
    let out = "";
    let i = 0;
    let quoted = false;
    while (i < src.length) {
      const ch = src[i];
      if (ch === '"') { quoted = !quoted; out += ch; i++; continue; }
      const prev = src[i - 1] || "";
      if (!quoted && /[A-Za-z$]/.test(ch) && !/[A-Za-z0-9_.]/.test(prev)) {
        REF_TEXT.lastIndex = i;
        const m = REF_TEXT.exec(src);
        if (m && !/[A-Za-z0-9(.]/.test(src[i + m[0].length] || "")) {
          const one = (d1, l, d2, n) => ({ colAbs: d1 === "$", c: colIndex(l), rowAbs: d2 === "$", r: Number(n) - 1 });
          const start = one(m[1], m[2], m[3], m[4]);
          const end = m[6] ? one(m[5], m[6], m[7], m[8]) : null;
          out += fn({ start, end });
          i += m[0].length;
          continue;
        }
        const word = /^[A-Za-z0-9_.$]+/.exec(src.slice(i));
        out += word[0];
        i += word[0].length;
        continue;
      }
      out += ch;
      i++;
    }
    return out;
  }
  const refText = p => (p.colAbs ? "$" : "") + colName(p.c) + (p.rowAbs ? "$" : "") + (p.r + 1);

  // The formula as it reads when copied `dr` rows down and `dc` columns across:
  // relative parts move, "$" parts stay, and anything pushed off the sheet is #REF!.
  function shift(formula, dr, dc) {
    if (!formula.startsWith("=")) return formula;
    return rewriteRefs(formula, ({ start, end }) => {
      const move = p => ({ ...p, r: p.rowAbs ? p.r : p.r + dr, c: p.colAbs ? p.c : p.c + dc });
      const a = move(start), b = end && move(end);
      if (a.r < 0 || a.c < 0 || (b && (b.r < 0 || b.c < 0))) return "#REF!";
      return refText(a) + (b ? ":" + refText(b) : "");
    });
  }

  // `k` is "r" or "c": the axis the row or column was removed from.
  function deleted({ start, end }, at, k) {
    if (!end) {
      if (start[k] === at) return "#REF!";
      return refText(start[k] > at ? { ...start, [k]: start[k] - 1 } : start);
    }
    let a = start, b = end;
    if (a[k] > b[k]) [a, b] = [b, a];
    if (a[k] === at && b[k] === at) return "#REF!";
    return refText({ ...a, [k]: a[k] > at ? a[k] - 1 : a[k] }) + ":" + refText({ ...b, [k]: b[k] >= at ? b[k] - 1 : b[k] });
  }
  function inserted({ start, end }, at, k) {
    const move = p => (p[k] >= at ? { ...p, [k]: p[k] + 1 } : p);
    if (!end) return refText(move(start));
    let a = start, b = end;
    if (a[k] > b[k]) [a, b] = [b, a];
    return refText(move(a)) + ":" + refText(move(b));
  }

  // ---- CSV ---------------------------------------------------------------
  function parseCsv(text) {
    if (text.startsWith(BOM)) text = text.slice(1);
    const rows = [];
    let row = [], field = "", i = 0, quoted = false;
    while (i < text.length) {
      const ch = text[i];
      if (quoted) {
        if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i += 2; continue; } quoted = false; i++; continue; }
        field += ch; i++; continue;
      }
      if (ch === '"') { quoted = true; i++; continue; }
      if (ch === ",") { row.push(field); field = ""; i++; continue; }
      if (ch === "\r") { i++; continue; }
      if (ch === "\n") { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
      field += ch; i++;
    }
    if (field !== "" || row.length) { row.push(field); rows.push(row); }
    return rows;
  }
  // Values, not formulas, as a spreadsheet saves CSV; the byte order mark makes
  // Excel read the Greek headers as UTF-8.
  function toCsv(grid) {
    const cell = v => {
      const s = isErr(v) ? v.code : toText(v);
      return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    return BOM + grid.map(r => r.map(cell).join(",")).join("\r\n") + "\r\n";
  }

  // What a cell shows: numbers without float noise, errors by their code.
  function display(v) {
    if (v === null || v === undefined) return "";
    if (isErr(v)) return v.code;
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (typeof v === "number") {
      if (Number.isInteger(v)) return String(v);
      return String(Math.round(v * 1e6) / 1e6);
    }
    return String(v);
  }

  const api = { quantile: (xs, p) => quantile(xs, p), Sheet, SheetError, isErr, parseCsv, toCsv, display, shift, colName, colIndex,
                cellName, literal, normalise, FUNCTION_NAMES };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.SheetEngine = api;
})(typeof window !== "undefined" ? window : this);
