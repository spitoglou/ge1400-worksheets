// Open a lab's companion PDFs (resources, quick reference, lecture slides) in a
// window over the page, so the worksheet stays where the student left it.
//
// The PDF is drawn with pdf.js rather than the browser's own viewer, for two
// reasons. A link inside the browser's viewer loads its target in the window
// itself, and the sites the materials point at (the WHO ICD browser, PubMed)
// refuse to be shown inside another page, so the student got a blank frame;
// here every link is an ordinary <a target="_blank">. And most phones have no
// viewer that shows a PDF inside a page at all, which pdf.js does not need.
// pdf.js is fetched the first time a file is opened, not with the page.
(function () {
  "use strict";
  const dialog = document.getElementById("material");
  if (!dialog) return;
  const title = dialog.querySelector(".material-title");
  const viewer = dialog.querySelector(".material-viewer");
  const openLink = dialog.querySelector(".material-open");
  const saveLink = dialog.querySelector(".material-save");
  let pdfjs = null;
  let current = 0;  // which open() is live, so a closed or replaced file stops drawing

  function loadPdfjs() {
    if (pdfjs) return pdfjs;
    pdfjs = new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "assets/pdf.min.js";
      s.onload = () => {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = "assets/pdf.worker.min.js";
        resolve(window.pdfjsLib);
      };
      s.onerror = () => { pdfjs = null; reject(new Error("pdf.js")); };
      document.head.appendChild(s);
    });
    return pdfjs;
  }

  function note(text) {
    viewer.replaceChildren();
    const p = document.createElement("p");
    p.className = "material-note";
    p.textContent = text;
    viewer.appendChild(p);
  }

  async function drawPage(doc, n, run) {
    const page = await doc.getPage(n);
    if (run !== current) return;
    const width = Math.min(viewer.clientWidth - 24, 1000);
    const base = page.getViewport({ scale: 1 });
    const vp = page.getViewport({ scale: width / base.width });
    const ratio = window.devicePixelRatio || 1;

    const box = document.createElement("div");
    box.className = "material-page";
    box.style.width = vp.width + "px";
    box.style.height = vp.height + "px";
    const canvas = document.createElement("canvas");
    canvas.width = Math.floor(vp.width * ratio);
    canvas.height = Math.floor(vp.height * ratio);
    box.appendChild(canvas);
    viewer.appendChild(box);
    await page.render({
      canvasContext: canvas.getContext("2d"),
      viewport: vp,
      transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : null,
    }).promise;

    // Links: one transparent <a> over each link rectangle, opening a new tab.
    for (const a of await page.getAnnotations({ intent: "display" })) {
      if (a.subtype !== "Link" || !a.url) continue;
      const [x1, y1, x2, y2] = vp.convertToViewportRectangle(a.rect);
      const link = document.createElement("a");
      link.href = a.url;
      link.target = "_blank";
      link.rel = "noopener";
      link.title = a.url;
      link.style.left = Math.min(x1, x2) + "px";
      link.style.top = Math.min(y1, y2) + "px";
      link.style.width = Math.abs(x2 - x1) + "px";
      link.style.height = Math.abs(y2 - y1) + "px";
      box.appendChild(link);
    }
  }

  async function show(href, label) {
    const run = ++current;
    title.textContent = label;
    openLink.href = href;
    saveLink.href = href;
    saveLink.setAttribute("download", decodeURIComponent(href.split("/").pop()));
    dialog.showModal();
    viewer.scrollTop = 0;
    note("Φόρτωση…");
    try {
      const lib = await loadPdfjs();
      const doc = await lib.getDocument(href).promise;
      if (run !== current) return;
      viewer.replaceChildren();
      for (let n = 1; n <= doc.numPages; n++) {
        await drawPage(doc, n, run);
        if (run !== current) return;
      }
    } catch (err) {
      if (run === current) {
        note("Το αρχείο δεν εμφανίζεται εδώ. Χρησιμοποιήστε το «Άνοιγμα σε νέα καρτέλα» ή τη «Λήψη» παραπάνω.");
      }
    }
  }

  document.querySelectorAll("a.material-link").forEach(a => {
    a.addEventListener("click", ev => {
      if (ev.ctrlKey || ev.metaKey || ev.shiftKey || ev.button === 1) return;
      ev.preventDefault();
      show(a.getAttribute("href"), a.dataset.label || a.textContent.trim());
    });
  });

  dialog.querySelector(".material-close").addEventListener("click", () => dialog.close());
  // a click on the backdrop, outside the window's box, closes it too
  dialog.addEventListener("click", ev => {
    if (ev.target === dialog) dialog.close();
  });
  dialog.addEventListener("close", () => { current++; viewer.replaceChildren(); });
})();
