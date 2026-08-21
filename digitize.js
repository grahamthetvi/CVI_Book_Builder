import { t, applyDomTranslations } from "./i18n.js";
import { EpubError, isEpubFile, renderEpubToPages } from "./epub-pages.js";

const PDFJS_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs";
const PDFJS_WORKER_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs";
const PADDLE_OCR_URL = "https://cdn.jsdelivr.net/npm/ppu-paddle-ocr@6.3.0/web/+esm";

const MAX_PAGES = 40;
const MAX_CROPS_PER_PAGE = 4;
const OCR_RENDER_SCALE = 2;

/** @typedef {{ id: string, name: string, blob: Blob, objectUrl: string, ocrText: string, ocrConfidence: number|null, ocrStatus: 'idle'|'pending'|'done'|'error', ocrError?: string, textSource?: 'epub'|'ocr'|null, crops: { id: string, name: string, file: File, objectUrl: string }[] }} DigitizePage */

/** @type {DigitizePage[]} */
let pages = [];
let selectedPageId = null;
let ocrServicePromise = null;
let ocrServiceModelKey = "";
let pdfjsPromise = null;

/** Crop drag state in image natural coordinates */
let cropDrag = null;

/** @type {null | {
 *   isolateBlob: (blob: Blob) => Promise<Blob>,
 *   rebuildSpreads: (spreads: { storyText: string, oddText: string, salientFeatures?: string, imageFiles?: File[] }[]) => void,
 *   setStatus: (text: string, isError?: boolean) => void,
 *   ensureCompatibleImage: (file: File) => Promise<File>,
 * }} */
let deps = null;

function el(id) {
  return document.getElementById(id);
}

function setDigitizeStatus(text, isError = false) {
  const status = el("digitizeStatus");
  if (!status) return;
  status.textContent = text || "";
  status.classList.toggle("error", Boolean(isError));
}

function newId(prefix) {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return `${prefix}-${crypto.randomUUID()}`;
  return `${prefix}-${Date.now()}-${Math.floor(Math.random() * 1e9)}`;
}

function revokePageUrls(page) {
  if (page.objectUrl) URL.revokeObjectURL(page.objectUrl);
  for (const crop of page.crops || []) {
    if (crop.objectUrl) URL.revokeObjectURL(crop.objectUrl);
  }
}

function clearPages() {
  pages.forEach(revokePageUrls);
  pages = [];
  selectedPageId = null;
  cropDrag = null;
}

function getSelectedPage() {
  return pages.find((p) => p.id === selectedPageId) || null;
}

function getModelKey() {
  const select = el("digitizeOcrModel");
  return select?.value || "v6-small";
}

async function getPdfjs() {
  if (pdfjsPromise) return pdfjsPromise;
  pdfjsPromise = import(PDFJS_URL).then((mod) => {
    const lib = mod.default || mod;
    if (lib.GlobalWorkerOptions) {
      lib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
    }
    return lib;
  }).catch((err) => {
    pdfjsPromise = null;
    throw err;
  });
  return pdfjsPromise;
}

async function getOcrService() {
  const modelKey = getModelKey();
  if (ocrServicePromise && ocrServiceModelKey === modelKey) return ocrServicePromise;

  ocrServiceModelKey = modelKey;
  ocrServicePromise = (async () => {
    setDigitizeStatus(t("javascriptStrings.digitize.loadingOcr"));
    const mod = await import(PADDLE_OCR_URL);
    const {
      PaddleOcrService,
      V6_SMALL_MODEL,
      V6_TINY_MODEL,
      V5_EN_MOBILE_MODEL,
      V5_ARABIC_MOBILE_MODEL
    } = mod;

    const modelMap = {
      "v6-small": V6_SMALL_MODEL,
      "v6-tiny": V6_TINY_MODEL,
      "v5-en-mobile": V5_EN_MOBILE_MODEL,
      "v5-arabic-mobile": V5_ARABIC_MOBILE_MODEL
    };
    const model = modelMap[modelKey] || V6_SMALL_MODEL;
    const service = new PaddleOcrService({ model, debugging: { verbose: false } });
    await service.initialize();
    return service;
  })().catch((err) => {
    ocrServicePromise = null;
    ocrServiceModelKey = "";
    throw err;
  });

  return ocrServicePromise;
}

function blobToArrayBuffer(blob) {
  return blob.arrayBuffer();
}

async function runOcrOnBlob(blob) {
  const service = await getOcrService();
  const buffer = await blobToArrayBuffer(blob);
  const result = await service.recognize(buffer, { flatten: false });
  const text = (result?.text || "").trim();
  const confidence = typeof result?.confidence === "number" ? result.confidence : null;
  return { text, confidence };
}

async function renderPdfToPages(file) {
  const pdfjs = await getPdfjs();
  const data = await file.arrayBuffer();
  const pdf = await pdfjs.getDocument({ data }).promise;
  const pageCount = Math.min(pdf.numPages, MAX_PAGES);
  const out = [];

  for (let i = 1; i <= pageCount; i += 1) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: OCR_RENDER_SCALE });
    const canvas = document.createElement("canvas");
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    await page.render({ canvasContext: ctx, viewport }).promise;
    const blob = await new Promise((resolve, reject) => {
      canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("PDF page render failed"))), "image/png");
    });
    const objectUrl = URL.createObjectURL(blob);
    out.push({
      id: newId("page"),
      name: `${file.name.replace(/\.[^/.]+$/, "")}-p${i}.png`,
      blob,
      objectUrl,
      ocrText: "",
      ocrConfidence: null,
      ocrStatus: "idle",
      crops: []
    });
  }

  if (pdf.numPages > MAX_PAGES) {
    setDigitizeStatus(t("javascriptStrings.digitize.truncatedPages", { max: MAX_PAGES, total: pdf.numPages }), false);
  }

  return out;
}

function epubErrorMessage(err) {
  if (err instanceof EpubError) {
    if (err.code === "drm") return t("javascriptStrings.digitize.epubDrm");
    if (err.code === "noPages") return t("javascriptStrings.digitize.epubNoPages");
    return t("javascriptStrings.digitize.epubInvalid");
  }
  return err?.message || t("javascriptStrings.errors.unknownFallback");
}

async function renderEpubFileToPages(file) {
  setDigitizeStatus(t("javascriptStrings.digitize.loadingEpub"));
  const { pages: raw, total } = await renderEpubToPages(file, {
    maxPages: MAX_PAGES,
    renderScale: OCR_RENDER_SCALE,
    ensureCompatibleImage: deps?.ensureCompatibleImage
  });
  if (total > MAX_PAGES) {
    setDigitizeStatus(t("javascriptStrings.digitize.truncatedPages", { max: MAX_PAGES, total }), false);
  }
  return raw.map((p) => ({
    id: newId("page"),
    name: p.name,
    blob: p.blob,
    objectUrl: URL.createObjectURL(p.blob),
    ocrText: p.ocrText || "",
    ocrConfidence: null,
    ocrStatus: p.textSource === "epub" && p.ocrText ? "done" : "idle",
    textSource: p.textSource || null,
    crops: []
  }));
}

async function filesToPages(fileList) {
  const files = Array.from(fileList || []);
  const out = [];

  for (const file of files) {
    if (out.length >= MAX_PAGES) break;
    const lower = (file.name || "").toLowerCase();
    if (file.type === "application/pdf" || lower.endsWith(".pdf")) {
      const pdfPages = await renderPdfToPages(file);
      for (const p of pdfPages) {
        if (out.length >= MAX_PAGES) break;
        out.push(p);
      }
      continue;
    }

    if (isEpubFile(file)) {
      const epubPages = await renderEpubFileToPages(file);
      for (const p of epubPages) {
        if (out.length >= MAX_PAGES) break;
        out.push(p);
      }
      continue;
    }

    if (!deps?.ensureCompatibleImage) continue;
    try {
      const ready = await deps.ensureCompatibleImage(file);
      const blob = ready;
      const objectUrl = URL.createObjectURL(blob);
      out.push({
        id: newId("page"),
        name: ready.name || file.name || `page-${out.length + 1}.png`,
        blob,
        objectUrl,
        ocrText: "",
        ocrConfidence: null,
        ocrStatus: "idle",
        crops: []
      });
    } catch (err) {
      console.error(err);
      setDigitizeStatus(
        `${t("javascriptStrings.digitize.couldNotLoadPage")} ${file.name}: ${err.message || t("javascriptStrings.errors.unknownFallback")}`,
        true
      );
    }
  }

  return out;
}

function deriveOddText(storyText) {
  const cleaned = (storyText || "").replace(/\s+/g, " ").trim();
  if (!cleaned) return "";
  const words = cleaned.split(" ").filter(Boolean);
  const phrase = words.slice(0, 3).join(" ");
  return phrase.length > 28 ? phrase.slice(0, 28).trim() : phrase;
}

function renderPageStrip() {
  const strip = el("digitizePageStrip");
  if (!strip) return;
  strip.innerHTML = "";

  if (!pages.length) {
    strip.innerHTML = `<p class="hint">${t("javascriptStrings.digitize.noPagesYet")}</p>`;
    return;
  }

  pages.forEach((page, index) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "digitize-page-thumb" + (page.id === selectedPageId ? " is-selected" : "");
    btn.setAttribute("aria-pressed", page.id === selectedPageId ? "true" : "false");

    const img = document.createElement("img");
    img.src = page.objectUrl;
    img.alt = t("javascriptStrings.digitize.pageThumbAlt", { n: index + 1 });

    const label = document.createElement("span");
    label.className = "digitize-page-thumb-label";
    const statusMark =
      page.ocrStatus === "done" ? "✓" :
      page.ocrStatus === "pending" ? "…" :
      page.ocrStatus === "error" ? "!" : "";
    label.textContent = `${index + 1}${statusMark ? ` ${statusMark}` : ""}`;

    btn.appendChild(img);
    btn.appendChild(label);
    btn.addEventListener("click", () => {
      selectedPageId = page.id;
      renderAll();
    });
    strip.appendChild(btn);
  });
}

function syncCropOverlaySize() {
  const img = el("digitizePageImage");
  const canvas = el("digitizeCropCanvas");
  if (!img || !canvas || !img.naturalWidth) return;
  const rect = img.getBoundingClientRect();
  canvas.width = Math.max(1, Math.round(rect.width));
  canvas.height = Math.max(1, Math.round(rect.height));
  canvas.style.width = `${rect.width}px`;
  canvas.style.height = `${rect.height}px`;
  drawCropOverlay();
}

function drawCropOverlay() {
  const canvas = el("digitizeCropCanvas");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  if (!cropDrag) return;

  const { x0, y0, x1, y1 } = cropDrag;
  const left = Math.min(x0, x1);
  const top = Math.min(y0, y1);
  const w = Math.abs(x1 - x0);
  const h = Math.abs(y1 - y0);
  if (w < 2 || h < 2) return;

  ctx.fillStyle = "rgba(0, 0, 0, 0.35)";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.clearRect(left, top, w, h);
  ctx.strokeStyle = "#7eb8ff";
  ctx.lineWidth = 2;
  ctx.setLineDash([6, 4]);
  ctx.strokeRect(left + 0.5, top + 0.5, w, h);
}

function canvasPointFromEvent(e) {
  const canvas = el("digitizeCropCanvas");
  const rect = canvas.getBoundingClientRect();
  return {
    x: Math.min(Math.max(0, e.clientX - rect.left), rect.width),
    y: Math.min(Math.max(0, e.clientY - rect.top), rect.height)
  };
}

function getCropRectInNaturalPixels() {
  const img = el("digitizePageImage");
  const canvas = el("digitizeCropCanvas");
  if (!img || !canvas || !cropDrag || !img.naturalWidth) return null;
  const { x0, y0, x1, y1 } = cropDrag;
  const left = Math.min(x0, x1);
  const top = Math.min(y0, y1);
  const right = Math.max(x0, x1);
  const bottom = Math.max(y0, y1);
  const w = right - left;
  const h = bottom - top;
  if (w < 8 || h < 8) return null;

  const scaleX = img.naturalWidth / canvas.width;
  const scaleY = img.naturalHeight / canvas.height;
  return {
    sx: Math.round(left * scaleX),
    sy: Math.round(top * scaleY),
    sw: Math.round(w * scaleX),
    sh: Math.round(h * scaleY)
  };
}

async function cropSelectionToBlob() {
  const page = getSelectedPage();
  const rect = getCropRectInNaturalPixels();
  if (!page || !rect) return null;

  const img = new Image();
  img.decoding = "async";
  const loaded = new Promise((resolve, reject) => {
    img.onload = () => resolve();
    img.onerror = () => reject(new Error("Could not load page image for crop."));
  });
  img.src = page.objectUrl;
  await loaded;

  const canvas = document.createElement("canvas");
  canvas.width = rect.sw;
  canvas.height = rect.sh;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(img, rect.sx, rect.sy, rect.sw, rect.sh, 0, 0, rect.sw, rect.sh);
  return new Promise((resolve, reject) => {
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("Crop failed."))), "image/png");
  });
}

function renderCropsList() {
  const list = el("digitizeCropsList");
  if (!list) return;
  list.innerHTML = "";
  const page = getSelectedPage();
  if (!page || !page.crops.length) {
    list.innerHTML = `<p class="hint">${t("javascriptStrings.digitize.noCropsYet")}</p>`;
    return;
  }

  page.crops.forEach((crop, index) => {
    const item = document.createElement("div");
    item.className = "digitize-crop-item";
    const img = document.createElement("img");
    img.src = crop.objectUrl;
    img.alt = crop.name;
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "danger";
    removeBtn.textContent = "✖";
    removeBtn.title = t("javascriptStrings.digitize.removeCrop");
    removeBtn.addEventListener("click", () => {
      URL.revokeObjectURL(crop.objectUrl);
      page.crops.splice(index, 1);
      renderCropsList();
    });
    item.appendChild(img);
    item.appendChild(removeBtn);
    list.appendChild(item);
  });
}

function renderWorkspace() {
  const empty = el("digitizeWorkspaceEmpty");
  const workspace = el("digitizeWorkspace");
  const img = el("digitizePageImage");
  const textArea = el("digitizeOcrText");
  const conf = el("digitizeOcrConfidence");
  const page = getSelectedPage();

  if (!page) {
    if (empty) empty.hidden = false;
    if (workspace) workspace.hidden = true;
    return;
  }

  if (empty) empty.hidden = true;
  if (workspace) workspace.hidden = false;

  if (img) {
    img.onload = () => syncCropOverlaySize();
    img.src = page.objectUrl;
  }
  if (textArea) {
    textArea.value = page.ocrText || "";
  }
  if (conf) {
    if (page.ocrStatus === "pending") {
      conf.textContent = t("javascriptStrings.digitize.ocrRunning");
    } else if (page.ocrStatus === "error") {
      conf.textContent = page.ocrError || t("javascriptStrings.digitize.ocrFailed");
    } else if (page.textSource === "epub") {
      conf.textContent = t("javascriptStrings.digitize.textFromEpub");
    } else if (page.ocrConfidence != null) {
      conf.textContent = t("javascriptStrings.digitize.ocrConfidence", {
        pct: Math.round(page.ocrConfidence * 100)
      });
    } else if (page.ocrStatus === "done") {
      conf.textContent = t("javascriptStrings.digitize.ocrDoneNoConfidence");
    } else {
      conf.textContent = t("javascriptStrings.digitize.ocrNotRun");
    }
  }

  cropDrag = null;
  renderCropsList();
  requestAnimationFrame(() => syncCropOverlaySize());
}

function renderAll() {
  renderPageStrip();
  renderWorkspace();
  const buildBtn = el("digitizeBuildBookBtn");
  const runOcrBtn = el("digitizeRunOcrBtn");
  const clearBtn = el("digitizeClearBtn");
  if (buildBtn) buildBtn.disabled = !pages.length;
  if (runOcrBtn) runOcrBtn.disabled = !pages.length;
  if (clearBtn) clearBtn.disabled = !pages.length;
}

async function handleUpload(fileList) {
  if (!fileList || !fileList.length) return;
  setDigitizeStatus(t("javascriptStrings.digitize.loadingPages"));
  try {
    const newPages = await filesToPages(fileList);
    if (!newPages.length) {
      setDigitizeStatus(t("javascriptStrings.digitize.noValidPages"), true);
      return;
    }
    if (pages.length + newPages.length > MAX_PAGES) {
      const room = Math.max(0, MAX_PAGES - pages.length);
      pages.push(...newPages.slice(0, room));
      setDigitizeStatus(t("javascriptStrings.digitize.truncatedPages", { max: MAX_PAGES, total: pages.length + newPages.length }), false);
    } else {
      pages.push(...newPages);
      setDigitizeStatus(t("javascriptStrings.digitize.pagesLoaded", { n: pages.length }));
    }
    if (!selectedPageId && pages.length) selectedPageId = pages[0].id;
    renderAll();
  } catch (err) {
    console.error(err);
    setDigitizeStatus(
      `${t("javascriptStrings.digitize.loadFailed")}${epubErrorMessage(err)}`,
      true
    );
  }
}

async function runOcrOnAllPages() {
  if (!pages.length) return;
  const runBtn = el("digitizeRunOcrBtn");
  if (runBtn) runBtn.disabled = true;

  try {
    await getOcrService();
    for (let i = 0; i < pages.length; i += 1) {
      const page = pages[i];
      page.ocrStatus = "pending";
      page.ocrError = undefined;
      renderAll();
      setDigitizeStatus(t("javascriptStrings.digitize.ocrProgress", { current: i + 1, total: pages.length }));
      try {
        const { text, confidence } = await runOcrOnBlob(page.blob);
        page.ocrText = text;
        page.ocrConfidence = confidence;
        page.ocrStatus = "done";
        page.textSource = "ocr";
      } catch (err) {
        console.error(err);
        page.ocrStatus = "error";
        page.ocrError = err.message || t("javascriptStrings.errors.unknownFallback");
      }
      if (page.id === selectedPageId) {
        const textArea = el("digitizeOcrText");
        if (textArea) textArea.value = page.ocrText || "";
      }
      renderAll();
    }
    const errors = pages.filter((p) => p.ocrStatus === "error").length;
    if (errors) {
      setDigitizeStatus(t("javascriptStrings.digitize.ocrFinishedWithErrors", { n: errors }), true);
    } else {
      setDigitizeStatus(t("javascriptStrings.digitize.ocrFinished"));
    }
  } catch (err) {
    console.error(err);
    setDigitizeStatus(
      `${t("javascriptStrings.digitize.ocrInitFailed")}${err.message || t("javascriptStrings.errors.unknownFallback")}`,
      true
    );
  } finally {
    if (runBtn) runBtn.disabled = !pages.length;
  }
}

async function isolateCurrentCrop() {
  const page = getSelectedPage();
  if (!page) return;
  if (page.crops.length >= MAX_CROPS_PER_PAGE) {
    setDigitizeStatus(t("javascriptStrings.digitize.maxCrops"), true);
    return;
  }
  if (!getCropRectInNaturalPixels()) {
    setDigitizeStatus(t("javascriptStrings.digitize.drawCropFirst"), true);
    return;
  }

  const isolateBtn = el("digitizeIsolateCropBtn");
  if (isolateBtn) isolateBtn.disabled = true;
  setDigitizeStatus(t("javascriptStrings.digitize.isolatingCrop"));

  try {
    const cropBlob = await cropSelectionToBlob();
    if (!cropBlob) throw new Error(t("javascriptStrings.digitize.drawCropFirst"));
    if (!deps?.isolateBlob) throw new Error("Background remover is not ready.");
    const isolated = await deps.isolateBlob(cropBlob);
    const file = new File(
      [isolated],
      `${page.name.replace(/\.[^/.]+$/, "")}-crop${page.crops.length + 1}.png`,
      { type: "image/png" }
    );
    page.crops.push({
      id: newId("crop"),
      name: file.name,
      file,
      objectUrl: URL.createObjectURL(file)
    });
    cropDrag = null;
    drawCropOverlay();
    renderCropsList();
    setDigitizeStatus(t("javascriptStrings.digitize.cropIsolated"));
  } catch (err) {
    console.error(err);
    setDigitizeStatus(
      `${t("javascriptStrings.digitize.isolateFailed")}${err.message || t("javascriptStrings.errors.unknownFallback")}`,
      true
    );
  } finally {
    if (isolateBtn) isolateBtn.disabled = false;
  }
}

function buildBookFromPages() {
  if (!pages.length || !deps?.rebuildSpreads) return;

  const hasAnyText = pages.some((p) => (p.ocrText || "").trim());
  if (!hasAnyText) {
    const ok = window.confirm(t("javascriptStrings.digitize.confirmBuildWithoutOcr"));
    if (!ok) return;
  } else {
    const ok = window.confirm(t("javascriptStrings.digitize.confirmBuildReplace"));
    if (!ok) return;
  }

  // Persist any in-progress textarea edit
  const page = getSelectedPage();
  const textArea = el("digitizeOcrText");
  if (page && textArea) page.ocrText = textArea.value;

  const spreads = pages.map((p) => {
    const storyText = (p.ocrText || "").trim();
    return {
      storyText,
      oddText: deriveOddText(storyText),
      salientFeatures: "",
      imageFiles: p.crops.map((c) => c.file).slice(0, 4)
    };
  });

  deps.rebuildSpreads(spreads);
  setDigitizeStatus(t("javascriptStrings.digitize.bookBuilt", { n: spreads.length }));
  deps.setStatus?.(t("javascriptStrings.digitize.bookBuilt", { n: spreads.length }));
}

function initCropInteraction() {
  const canvas = el("digitizeCropCanvas");
  if (!canvas) return;

  const onDown = (e) => {
    if (!getSelectedPage()) return;
    e.preventDefault();
    const pt = canvasPointFromEvent(e);
    cropDrag = { x0: pt.x, y0: pt.y, x1: pt.x, y1: pt.y };
    drawCropOverlay();
  };
  const onMove = (e) => {
    if (!cropDrag) return;
    e.preventDefault();
    const pt = canvasPointFromEvent(e);
    cropDrag.x1 = pt.x;
    cropDrag.y1 = pt.y;
    drawCropOverlay();
  };
  const onUp = () => {
    if (!cropDrag) return;
    const w = Math.abs(cropDrag.x1 - cropDrag.x0);
    const h = Math.abs(cropDrag.y1 - cropDrag.y0);
    if (w < 8 || h < 8) {
      cropDrag = null;
      drawCropOverlay();
    }
  };

  canvas.addEventListener("pointerdown", onDown);
  window.addEventListener("pointermove", onMove);
  window.addEventListener("pointerup", onUp);
  window.addEventListener("resize", () => syncCropOverlaySize());
}

/**
 * @param {{
 *   isolateBlob: (blob: Blob) => Promise<Blob>,
 *   rebuildSpreads: (spreads: { storyText: string, oddText: string, salientFeatures?: string, imageFiles?: File[] }[]) => void,
 *   setStatus: (text: string, isError?: boolean) => void,
 *   ensureCompatibleImage: (file: File) => Promise<File>,
 * }} options
 */
export function initDigitizeBook(options) {
  deps = options;
  const panel = el("digitizeBookPanel");
  if (!panel) return;

  const pdfInput = el("digitizePdfInput");
  const imagesInput = el("digitizeImagesInput");
  const runOcrBtn = el("digitizeRunOcrBtn");
  const isolateBtn = el("digitizeIsolateCropBtn");
  const clearCropBtn = el("digitizeClearCropBtn");
  const buildBtn = el("digitizeBuildBookBtn");
  const clearBtn = el("digitizeClearBtn");
  const textArea = el("digitizeOcrText");
  const modelSelect = el("digitizeOcrModel");

  if (pdfInput) {
    pdfInput.addEventListener("change", async () => {
      await handleUpload(pdfInput.files);
      pdfInput.value = "";
    });
  }
  if (imagesInput) {
    imagesInput.addEventListener("change", async () => {
      await handleUpload(imagesInput.files);
      imagesInput.value = "";
    });
  }
  if (runOcrBtn) runOcrBtn.addEventListener("click", () => runOcrOnAllPages());
  if (isolateBtn) isolateBtn.addEventListener("click", () => isolateCurrentCrop());
  if (clearCropBtn) {
    clearCropBtn.addEventListener("click", () => {
      cropDrag = null;
      drawCropOverlay();
    });
  }
  if (buildBtn) buildBtn.addEventListener("click", () => buildBookFromPages());
  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      if (!pages.length) return;
      if (!window.confirm(t("javascriptStrings.digitize.confirmClearPages"))) return;
      clearPages();
      renderAll();
      setDigitizeStatus(t("javascriptStrings.digitize.pagesCleared"));
    });
  }
  if (textArea) {
    textArea.addEventListener("input", () => {
      const page = getSelectedPage();
      if (!page) return;
      page.ocrText = textArea.value;
    });
  }
  if (modelSelect) {
    modelSelect.addEventListener("change", () => {
      ocrServicePromise = null;
      ocrServiceModelKey = "";
      setDigitizeStatus(t("javascriptStrings.digitize.modelChanged"));
    });
  }

  initCropInteraction();
  renderAll();
  applyDomTranslations(panel);
}

export function refreshDigitizeLocale() {
  const panel = el("digitizeBookPanel");
  if (panel) applyDomTranslations(panel);
  renderAll();
}
