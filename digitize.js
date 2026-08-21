import { t, applyDomTranslations } from "./i18n.js";

const PDFJS_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs";
const PDFJS_WORKER_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs";
const PADDLE_OCR_URL = "https://cdn.jsdelivr.net/npm/ppu-paddle-ocr@6.3.0/web/+esm";

const MAX_PAGES = 40;
const MAX_CROPS_PER_PAGE = 4;
const OCR_RENDER_SCALE = 2;
const OCR_MIN_SIDE = 1400;
const OCR_MAX_SIDE = 2200;
/** How strongly saturated (colored) pixels are darkened so yellow/cyan/pink text reads as ink. */
const OCR_SATURATION_DARKEN = 0.84;
/** Mean luminance below this is treated as light-on-dark and inverted before OCR. */
const OCR_INVERT_LUMA_THRESHOLD = 118;

/** @typedef {{ id: string, name: string, blob: Blob, objectUrl: string, ocrText: string, ocrConfidence: number|null, ocrStatus: 'idle'|'pending'|'done'|'error', ocrError?: string, crops: { id: string, name: string, file: File, objectUrl: string }[] }} DigitizePage */

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

function scoreOcrResult(result) {
  const text = (result?.text || "").trim();
  const confidence = typeof result?.confidence === "number" ? result.confidence : null;
  let letters = 0;
  try {
    letters = (text.match(/\p{L}|\p{N}/gu) || []).length;
  } catch {
    letters = (text.match(/[A-Za-z0-9\u0600-\u06FF]/g) || []).length;
  }
  const conf = confidence == null ? 0.55 : confidence;
  return { text, confidence, letters, score: letters * (0.35 + conf) };
}

/**
 * Darken colored (high-saturation) pixels and invert light-on-dark pages.
 * Children's-book yellow/cyan/pink letters often have too little luminance
 * contrast for PaddleOCR; treating saturation as ink recovers them.
 */
function enhanceImageDataForColoredText(imageData) {
  const { data, width, height } = imageData;
  const pixelCount = width * height;
  let lumaSum = 0;
  for (let i = 0; i < data.length; i += 4) {
    lumaSum += 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
  }
  const invert = lumaSum / pixelCount < OCR_INVERT_LUMA_THRESHOLD;

  const outR = new Uint8ClampedArray(pixelCount);
  const outG = new Uint8ClampedArray(pixelCount);
  const outB = new Uint8ClampedArray(pixelCount);
  const luma = new Uint8ClampedArray(pixelCount);
  const hist = new Uint32Array(256);

  for (let i = 0, p = 0; i < data.length; i += 4, p += 1) {
    let r = data[i];
    let g = data[i + 1];
    let b = data[i + 2];
    if (invert) {
      r = 255 - r;
      g = 255 - g;
      b = 255 - b;
    }
    const maxc = Math.max(r, g, b);
    const minc = Math.min(r, g, b);
    const chroma = maxc - minc;
    const sat = maxc === 0 ? 0 : chroma / maxc;
    const factor = 1 - sat * OCR_SATURATION_DARKEN;
    const nr = r * factor;
    const ng = g * factor;
    const nb = b * factor;
    outR[p] = nr;
    outG[p] = ng;
    outB[p] = nb;
    const y = Math.round(0.299 * nr + 0.587 * ng + 0.114 * nb);
    luma[p] = y;
    hist[y] += 1;
  }

  const loCut = pixelCount * 0.015;
  const hiCut = pixelCount * 0.985;
  let acc = 0;
  let lo = 0;
  let hi = 255;
  for (let v = 0; v < 256; v += 1) {
    acc += hist[v];
    if (lo === 0 && acc >= loCut) lo = v;
    if (acc >= hiCut) {
      hi = v;
      break;
    }
  }
  const range = hi - lo;
  const skipStretch = range < 8;

  for (let p = 0, i = 0; p < pixelCount; p += 1, i += 4) {
    if (skipStretch) {
      data[i] = outR[p];
      data[i + 1] = outG[p];
      data[i + 2] = outB[p];
      continue;
    }
    const stretched = ((luma[p] - lo) / range) * 255;
    const gain = luma[p] <= 0 ? 1 : stretched / luma[p];
    data[i] = Math.max(0, Math.min(255, outR[p] * gain));
    data[i + 1] = Math.max(0, Math.min(255, outG[p] * gain));
    data[i + 2] = Math.max(0, Math.min(255, outB[p] * gain));
  }
}

async function enhanceBlobForOcr(blob) {
  const bitmap = await createImageBitmap(blob);
  const srcW = bitmap.width;
  const srcH = bitmap.height;
  const minSide = Math.min(srcW, srcH);
  const maxSide = Math.max(srcW, srcH);
  let scale = 1;
  if (minSide > 0 && minSide < OCR_MIN_SIDE) scale = OCR_MIN_SIDE / minSide;
  if (maxSide * scale > OCR_MAX_SIDE) scale = OCR_MAX_SIDE / maxSide;

  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(srcW * scale));
  canvas.height = Math.max(1, Math.round(srcH * scale));
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  if (typeof bitmap.close === "function") bitmap.close();

  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  enhanceImageDataForColoredText(imageData);
  ctx.putImageData(imageData, 0, 0);

  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (b) => (b ? resolve(b) : reject(new Error("OCR preprocess failed"))),
      "image/png"
    );
  });
}

async function recognizeBlob(service, blob) {
  const buffer = await blobToArrayBuffer(blob);
  return service.recognize(buffer, { flatten: false });
}

async function runOcrOnBlob(blob) {
  const service = await getOcrService();
  let enhanced;
  try {
    enhanced = await enhanceBlobForOcr(blob);
  } catch (err) {
    console.warn("OCR color preprocess failed; using original page image.", err);
  }

  const primaryBlob = enhanced || blob;
  const primary = scoreOcrResult(await recognizeBlob(service, primaryBlob));
  const primaryLooksOk = primary.letters >= 3 && (primary.confidence == null || primary.confidence >= 0.45);

  if (primaryLooksOk || !enhanced) {
    return { text: primary.text, confidence: primary.confidence };
  }

  const fallback = scoreOcrResult(await recognizeBlob(service, blob));
  const best = fallback.score >= primary.score ? fallback : primary;
  return { text: best.text, confidence: best.confidence };

  const fallback = scoreOcrResult(await recognizeBlob(service, blob));
  const best = fallback.score >= primary.score ? fallback : primary;
  return { text: best.text, confidence: best.confidence };
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
  const hasPages = Boolean(pages.length);
  const buildBtn = el("digitizeBuildBookBtn");
  const runOcrBtn = el("digitizeRunOcrBtn");
  const clearBtn = el("digitizeClearBtn");
  const copyPageBtn = el("digitizeCopyPageTextBtn");
  const copyAiBtn = el("digitizeCopyAiReviewBtn");
  const applyAiBtn = el("digitizeApplyAiReviewBtn");
  if (buildBtn) buildBtn.disabled = !hasPages;
  if (runOcrBtn) runOcrBtn.disabled = !hasPages;
  if (clearBtn) clearBtn.disabled = !hasPages;
  if (copyPageBtn) copyPageBtn.disabled = !getSelectedPage();
  if (copyAiBtn) copyAiBtn.disabled = !hasPages;
  if (applyAiBtn) applyAiBtn.disabled = !hasPages;
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
      `${t("javascriptStrings.digitize.loadFailed")}${err.message || t("javascriptStrings.errors.unknownFallback")}`,
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

function persistCurrentOcrText() {
  const page = getSelectedPage();
  const textArea = el("digitizeOcrText");
  if (page && textArea) page.ocrText = textArea.value;
}

function copyTextToClipboard(text, successMsg, failMsg) {
  const done = (ok) => {
    setDigitizeStatus(ok ? successMsg : failMsg, !ok);
    if (ok) deps?.setStatus?.(successMsg);
    else deps?.setStatus?.(failMsg, true);
  };

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(
      () => done(true),
      () => done(false)
    );
    return;
  }

  try {
    const textArea = document.createElement("textarea");
    textArea.value = text;
    textArea.style.top = "0";
    textArea.style.left = "0";
    textArea.style.position = "fixed";
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    const successful = document.execCommand("copy");
    document.body.removeChild(textArea);
    done(Boolean(successful));
  } catch (err) {
    console.error(err);
    done(false);
  }
}

function getBookSetupValues() {
  return {
    title: (el("bookTitle")?.value || "").trim(),
    eccArea: (el("eccArea")?.value || "").trim(),
    activityPrompt: (el("activityPrompt")?.value || "").trim()
  };
}

function applyBookSetupValues({ title, eccArea, activityPrompt }) {
  if (title && el("bookTitle")) el("bookTitle").value = title;
  if (eccArea && el("eccArea")) el("eccArea").value = eccArea;
  if (activityPrompt && el("activityPrompt")) el("activityPrompt").value = activityPrompt;
}

function parseSpreadTaggedText(raw) {
  const titleMatch = raw.match(/^\s*TITLE:\s*(.+)\s*$/im);
  const title = titleMatch ? titleMatch[1].trim() : "";
  const eccMatch = raw.match(/^\s*ECC_AREA:\s*(.+)\s*$/im);
  const eccArea = eccMatch ? eccMatch[1].trim() : "";
  const activityMatch = raw.match(/^\s*ACTIVITY_PROMPT:\s*(.+)\s*$/im);
  const activityPrompt = activityMatch ? activityMatch[1].trim() : "";

  const chunks = raw.split(/^\s*SPREAD:\s*$/gim).map((c) => c.trim()).filter(Boolean);
  const spreads = [];

  for (const chunk of chunks) {
    const storyMatch = chunk.match(/^\s*STORY:\s*([\s\S]*?)(?=^\s*SALIENT_FEATURES:|^\s*ODD_TEXT:|^\s*IMAGE_PROMPT:|\s*$)/im);
    const salientMatch = chunk.match(/^\s*SALIENT_FEATURES:\s*([\s\S]*?)(?=^\s*ODD_TEXT:|^\s*IMAGE_PROMPT:|\s*$)/im);
    const oddMatch = chunk.match(/^\s*ODD_TEXT:\s*(.+)\s*$/im);
    const imagePromptMatch = chunk.match(/^\s*IMAGE_PROMPT:\s*([\s\S]*?)(?=^\s*SPREAD:|\s*$)/im);
    const storyText = storyMatch ? storyMatch[1].trim() : "";
    const salientFeatures = salientMatch ? salientMatch[1].trim() : "";
    const oddText = oddMatch ? oddMatch[1].trim() : "";
    const imagePrompt = imagePromptMatch ? imagePromptMatch[1].trim() : "";
    if (storyText || oddText || salientFeatures || imagePrompt) {
      spreads.push({ storyText, salientFeatures, oddText, imagePrompt });
    }
  }

  return { title, eccArea, activityPrompt, spreads };
}

function parsePageTaggedText(raw) {
  const re = /^\s*PAGE\s+(\d+)\s*:\s*$/gim;
  const starts = [];
  let match;
  while ((match = re.exec(raw))) {
    starts.push({ n: Number(match[1]), at: match.index, len: match[0].length });
  }
  if (!starts.length) return [];

  const byIndex = [];
  for (let i = 0; i < starts.length; i += 1) {
    const start = starts[i].at + starts[i].len;
    const end = i + 1 < starts.length ? starts[i + 1].at : raw.length;
    const text = raw.slice(start, end).trim();
    const idx = Math.max(0, starts[i].n - 1);
    byIndex[idx] = {
      storyText: text,
      oddText: deriveOddText(text),
      salientFeatures: "",
      imagePrompt: ""
    };
  }
  return byIndex;
}

function parseDigitizeAiText(raw) {
  const fenced = raw.match(/```(?:[\w-]+)?\s*([\s\S]*?)```/);
  const body = fenced ? fenced[1].trim() : raw;
  const tagged = parseSpreadTaggedText(body);
  if (tagged.spreads.length) return tagged;
  const pageSpreads = parsePageTaggedText(body);
  return {
    title: tagged.title,
    eccArea: tagged.eccArea,
    activityPrompt: tagged.activityPrompt,
    spreads: pageSpreads
  };
}

function buildPageCopyPrompt(page, index) {
  const body = (page?.ocrText || "").trim() || t("javascriptStrings.digitize.emptyPageOcr");
  return `${t("javascriptStrings.digitize.pageCopyPrompt")}

--- ${t("javascriptStrings.digitize.pageHeading", { n: index + 1 })} ---
${body}`.trim();
}

function buildAllPagesAiReviewPrompt() {
  persistCurrentOcrText();
  const setup = getBookSetupValues();
  const pageBlocks = pages.map((page, index) => {
    const body = (page.ocrText || "").trim() || t("javascriptStrings.digitize.emptyPageOcr");
    return `--- ${t("javascriptStrings.digitize.pageHeading", { n: index + 1 })} ---
${body}`;
  }).join("\n\n");

  return `${t("javascriptStrings.digitize.aiReviewIntro")}

${t("javascriptStrings.digitize.aiReviewSetupHeading")}
- TITLE: ${setup.title || "[Book title]"}
- ECC_AREA: ${setup.eccArea || "[ECC area]"}
- ACTIVITY_PROMPT: ${setup.activityPrompt || "[One-sentence sensory activity]"}

${t("javascriptStrings.digitize.aiReviewTemplate")}

${t("javascriptStrings.digitize.aiReviewPagesHeading")}

${pageBlocks}`.trim();
}

function copySelectedPageText() {
  persistCurrentOcrText();
  const page = getSelectedPage();
  if (!page) {
    setDigitizeStatus(t("javascriptStrings.digitize.selectPageToCopy"), true);
    return;
  }
  const index = pages.indexOf(page);
  const text = (page.ocrText || "").trim();
  if (!text) {
    setDigitizeStatus(t("javascriptStrings.digitize.noPageTextToCopy"), true);
    return;
  }
  copyTextToClipboard(
    buildPageCopyPrompt(page, index < 0 ? 0 : index),
    t("javascriptStrings.digitize.pageTextCopied"),
    t("javascriptStrings.digitize.copyFailed")
  );
}

function copyAllPagesForAiReview() {
  if (!pages.length) {
    setDigitizeStatus(t("javascriptStrings.digitize.noPagesYet"), true);
    return;
  }
  persistCurrentOcrText();
  const hasAnyText = pages.some((p) => (p.ocrText || "").trim());
  if (!hasAnyText) {
    setDigitizeStatus(t("javascriptStrings.digitize.noAllTextToCopy"), true);
    return;
  }
  copyTextToClipboard(
    buildAllPagesAiReviewPrompt(),
    t("javascriptStrings.digitize.aiReviewCopied"),
    t("javascriptStrings.digitize.copyFailed")
  );
}

function applyAiReviewToSpreads() {
  if (!pages.length || !deps?.rebuildSpreads) return;
  persistCurrentOcrText();

  const raw = (el("digitizeAiReviewInput")?.value || "").trim();
  if (!raw) {
    setDigitizeStatus(t("javascriptStrings.digitize.pasteAiReviewFirst"), true);
    return;
  }

  const parsed = parseDigitizeAiText(raw);
  if (!parsed.spreads.length) {
    setDigitizeStatus(t("javascriptStrings.digitize.aiReviewNoSpreads"), true);
    return;
  }

  if (!window.confirm(t("javascriptStrings.digitize.confirmApplyAiReview"))) return;

  applyBookSetupValues(parsed);

  const count = Math.max(pages.length, parsed.spreads.length);
  const spreads = [];
  for (let i = 0; i < count; i += 1) {
    const page = pages[i];
    const ai = parsed.spreads[i];
    const storyText = (ai?.storyText || page?.ocrText || "").trim();
    if (page && ai?.storyText) page.ocrText = ai.storyText;
    spreads.push({
      storyText,
      oddText: (ai?.oddText || deriveOddText(storyText)).trim(),
      salientFeatures: ai?.salientFeatures || "",
      imagePrompt: ai?.imagePrompt || "",
      imageFiles: page ? page.crops.map((c) => c.file).slice(0, 4) : []
    });
  }

  deps.rebuildSpreads(spreads);
  renderAll();
  setDigitizeStatus(t("javascriptStrings.digitize.aiReviewApplied", { n: spreads.length }));
  deps.setStatus?.(t("javascriptStrings.digitize.aiReviewApplied", { n: spreads.length }));
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

  persistCurrentOcrText();

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
  const copyPageBtn = el("digitizeCopyPageTextBtn");
  const copyAiBtn = el("digitizeCopyAiReviewBtn");
  const applyAiBtn = el("digitizeApplyAiReviewBtn");

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
  if (copyPageBtn) copyPageBtn.addEventListener("click", () => copySelectedPageText());
  if (copyAiBtn) copyAiBtn.addEventListener("click", () => copyAllPagesForAiReview());
  if (applyAiBtn) applyAiBtn.addEventListener("click", () => applyAiReviewToSpreads());
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
