import { t, applyDomTranslations } from "./i18n.js";
import { EpubError, isEpubFile, renderEpubToPages } from "./epub-pages.js";
import { joinPdfTextItems, pairNookSpreads } from "./nook-pdf.js";
import { extractObjectFromImage, normalizedPointOnImage, objectExtractUserMessage } from "./object-extract.js";

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

/** @typedef {{ x0: number, y0: number, x1: number, y1: number }} NormalizedCrop */
/** @typedef {{ id: string, name: string, blob: Blob, objectUrl: string, ocrText: string, ocrConfidence: number|null, ocrStatus: 'idle'|'pending'|'done'|'error', ocrError?: string, textSource?: 'epub'|'ocr'|null, pendingCrop?: NormalizedCrop|null, crops: { id: string, name: string, file: File, objectUrl: string }[] }} DigitizePage */

/** @type {DigitizePage[]} */
let pages = [];
let selectedPageId = null;
let ocrServicePromise = null;
let ocrServiceModelKey = "";
let pdfjsPromise = null;

/** Crop drag in normalized image coordinates (0–1), so it survives overlay resizes. */
let cropDrag = null;
let cropDragging = false;
let cropResizeObserver = null;
let extractBusy = false;

/** @type {null | {
 *   isolateBlob: (blob: Blob) => Promise<Blob>,
 *   decorateCutout?: (blob: Blob) => Promise<Blob>,
 *   rebuildSpreads: (spreads: { storyText: string, oddText: string, salientFeatures?: string, imageFiles?: File[] }[]) => void,
 *   setStatus: (text: string, isError?: boolean) => void,
 *   ensureCompatibleImage: (file: File) => Promise<File>,
 *   setBookTitle?: (title: string) => void,
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
  cropDragging = false;
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
}

async function renderPdfPageToCanvas(page, scale = OCR_RENDER_SCALE) {
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.floor(viewport.width);
  canvas.height = Math.floor(viewport.height);
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  await page.render({ canvasContext: ctx, viewport }).promise;
  return canvas;
}

function canvasToPngBlob(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("Could not encode image."))), "image/png");
  });
}

async function getPdfPageText(page) {
  const content = await page.getTextContent();
  return joinPdfTextItems(content?.items || []);
}

const CONTENT_COLOR_THRESHOLD = 28;

function medianChannel(values) {
  if (!values.length) return 255;
  const sorted = values.slice().sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)];
}

/**
 * Margin color from the middle of each edge, inset past a thin frame.
 * Corner samples are skipped so a border or page number does not become the background.
 */
function sampleMarginBackground(data, width, height) {
  const rgb = [];
  let transparent = 0;
  let total = 0;
  const insets = [6, 14, 28, 48].filter((v) => v < width / 5 && v < height / 5);
  const bands = insets.length ? insets : [0];
  const x0 = Math.floor(width * 0.3);
  const x1 = Math.ceil(width * 0.7);
  const y0 = Math.floor(height * 0.3);
  const y1 = Math.ceil(height * 0.7);

  const consider = (x, y) => {
    if (x < 0 || y < 0 || x >= width || y >= height) return;
    const i = (y * width + x) * 4;
    total += 1;
    if (data[i + 3] < 16) {
      transparent += 1;
      return;
    }
    rgb.push([data[i], data[i + 1], data[i + 2]]);
  };

  for (const inset of bands) {
    for (let x = x0; x < x1; x += 4) {
      consider(x, inset);
      consider(x, height - 1 - inset);
    }
    for (let y = y0; y < y1; y += 4) {
      consider(inset, y);
      consider(width - 1 - inset, y);
    }
  }

  if (!total || transparent > total * 0.6) {
    return { r: 0, g: 0, b: 0, transparent: true };
  }
  return {
    r: medianChannel(rgb.map((p) => p[0])),
    g: medianChannel(rgb.map((p) => p[1])),
    b: medianChannel(rgb.map((p) => p[2])),
    transparent: false
  };
}

function isMarginPixel(data, index, bg) {
  if (data[index + 3] < 16) return true;
  if (bg.transparent) return false;
  return (
    Math.abs(data[index] - bg.r) <= CONTENT_COLOR_THRESHOLD &&
    Math.abs(data[index + 1] - bg.g) <= CONTENT_COLOR_THRESHOLD &&
    Math.abs(data[index + 2] - bg.b) <= CONTENT_COLOR_THRESHOLD
  );
}

/** Bounding span of the photo, ignoring a thin border or page number. */
function spanOfSubstantialContent(counts, length, crossLength) {
  const minHits = Math.max(10, Math.floor(crossLength * 0.06));
  const minBand = Math.max(20, Math.floor(length * 0.035));
  const gapMerge = Math.max(12, Math.floor(length * 0.02));
  const content = [];
  for (let i = 0; i < length; i += 1) content.push(counts[i] >= minHits);

  const runs = [];
  let index = 0;
  while (index < length) {
    if (!content[index]) {
      index += 1;
      continue;
    }
    let end = index + 1;
    while (end < length && content[end]) end += 1;
    runs.push({ start: index, end: end - 1 });
    index = end;
  }
  if (!runs.length) return { start: 0, end: length - 1 };

  const merged = [];
  for (const run of runs) {
    const prev = merged[merged.length - 1];
    if (prev && run.start - prev.end - 1 <= gapMerge) prev.end = run.end;
    else merged.push({ start: run.start, end: run.end });
  }

  const substantial = merged.filter((run) => run.end - run.start + 1 >= minBand);
  const pool = substantial.length ? substantial : merged;
  const chosen = pool.reduce((best, run) => (
    run.end - run.start > best.end - best.start ? run : best
  ));
  return { start: chosen.start, end: chosen.end };
}

function findContentBounds(imageData, width, height) {
  const data = imageData.data;
  const bg = sampleMarginBackground(data, width, height);
  const rowCounts = new Uint32Array(height);
  const colCounts = new Uint32Array(width);

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const i = (y * width + x) * 4;
      if (!isMarginPixel(data, i, bg)) {
        rowCounts[y] += 1;
        colCounts[x] += 1;
      }
    }
  }

  const rows = spanOfSubstantialContent(rowCounts, height, width);
  const cols = spanOfSubstantialContent(colCounts, width, height);
  const pad = 8;
  const top = Math.max(0, rows.start - pad);
  const left = Math.max(0, cols.start - pad);
  const bottom = Math.min(height - 1, rows.end + pad);
  const right = Math.min(width - 1, cols.end + pad);
  const w = right - left + 1;
  const h = bottom - top + 1;

  if (w < 16 || h < 16 || w * h < width * height * 0.02) {
    return { x: 0, y: 0, w: width, h: height };
  }
  return { x: left, y: top, w, h };
}

function cropCanvasToContent(sourceCanvas) {
  const ctx = sourceCanvas.getContext("2d", { willReadFrequently: true });
  const { width, height } = sourceCanvas;
  if (!width || !height) return sourceCanvas;
  const imageData = ctx.getImageData(0, 0, width, height);
  const bounds = findContentBounds(imageData, width, height);
  if (bounds.w >= width - 2 && bounds.h >= height - 2) return sourceCanvas;
  const out = document.createElement("canvas");
  out.width = bounds.w;
  out.height = bounds.h;
  out.getContext("2d").drawImage(
    sourceCanvas,
    bounds.x,
    bounds.y,
    bounds.w,
    bounds.h,
    0,
    0,
    bounds.w,
    bounds.h
  );
  return out;
}

async function renderPdfToPages(file) {
  const pdfjs = await getPdfjs();
  const data = await file.arrayBuffer();
  const pdf = await pdfjs.getDocument({ data }).promise;
  const pageCount = Math.min(pdf.numPages, MAX_PAGES);
  const out = [];

  for (let i = 1; i <= pageCount; i += 1) {
    const page = await pdf.getPage(i);
    const canvas = await renderPdfPageToCanvas(page);
    const blob = await canvasToPngBlob(canvas);
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
  // A hidden drawer reports 0. Writing that size pins the overlay at 0×0 and drops every drag.
  if (rect.width < 2 || rect.height < 2) return;
  const dpr = window.devicePixelRatio || 1;
  const w = Math.max(1, Math.round(rect.width * dpr));
  const h = Math.max(1, Math.round(rect.height * dpr));
  if (canvas.width !== w || canvas.height !== h) {
    canvas.width = w;
    canvas.height = h;
  }
  canvas.style.width = "";
  canvas.style.height = "";
  drawCropOverlay();
}

function drawCropOverlay() {
  const canvas = el("digitizeCropCanvas");
  const img = el("digitizePageImage");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  if (!cropDrag || !img) return;

  const rect = img.getBoundingClientRect();
  const scaleX = rect.width > 1 ? canvas.width / rect.width : canvas.width;
  const scaleY = rect.height > 1 ? canvas.height / rect.height : canvas.height;
  const left = Math.min(cropDrag.x0, cropDrag.x1) * rect.width * scaleX;
  const top = Math.min(cropDrag.y0, cropDrag.y1) * rect.height * scaleY;
  const w = Math.abs(cropDrag.x1 - cropDrag.x0) * rect.width * scaleX;
  const h = Math.abs(cropDrag.y1 - cropDrag.y0) * rect.height * scaleY;
  if (w < 2 || h < 2) return;

  ctx.fillStyle = "rgba(0, 0, 0, 0.35)";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.clearRect(left, top, w, h);
  ctx.strokeStyle = "#7eb8ff";
  ctx.lineWidth = Math.max(2, Math.round((window.devicePixelRatio || 1) * 1.5));
  ctx.setLineDash([6, 4]);
  ctx.strokeRect(left + 0.5, top + 0.5, w, h);
}

function imageNormFromEvent(e) {
  const img = el("digitizePageImage");
  if (!img || !img.naturalWidth) return null;
  return normalizedPointOnImage(img, e.clientX, e.clientY);
}

function clickExtractEnabled() {
  return Boolean(el("digitizeClickExtract")?.checked);
}

function syncClickExtractFrame() {
  const frame = el("digitizePageImage")?.closest(".digitize-crop-frame");
  if (frame) frame.classList.toggle("is-click-extract", clickExtractEnabled());
}

function cropSelectionIsBigEnough(norm = cropDrag) {
  const img = el("digitizePageImage");
  if (!img || !norm || !img.naturalWidth) return false;
  const display = img.getBoundingClientRect();
  const w = Math.abs(norm.x1 - norm.x0);
  const h = Math.abs(norm.y1 - norm.y0);
  if (display.width >= 2 && display.height >= 2) {
    return w * display.width >= 8 && h * display.height >= 8;
  }
  return w * img.naturalWidth >= 8 && h * img.naturalHeight >= 8;
}

function getCropRectInNaturalPixels() {
  const img = el("digitizePageImage");
  if (!img || !cropDrag || !img.naturalWidth || !cropSelectionIsBigEnough()) return null;
  const left = Math.min(cropDrag.x0, cropDrag.x1);
  const top = Math.min(cropDrag.y0, cropDrag.y1);
  const width = Math.abs(cropDrag.x1 - cropDrag.x0);
  const height = Math.abs(cropDrag.y1 - cropDrag.y0);
  let sx = Math.round(left * img.naturalWidth);
  let sy = Math.round(top * img.naturalHeight);
  let sw = Math.round(width * img.naturalWidth);
  let sh = Math.round(height * img.naturalHeight);
  sx = Math.max(0, Math.min(sx, img.naturalWidth - 1));
  sy = Math.max(0, Math.min(sy, img.naturalHeight - 1));
  sw = Math.max(1, Math.min(sw, img.naturalWidth - sx));
  sh = Math.max(1, Math.min(sh, img.naturalHeight - sy));
  return { sx, sy, sw, sh };
}

function loadImageElement(url) {
  const img = new Image();
  img.decoding = "async";
  const loaded = new Promise((resolve, reject) => {
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Could not load page image for crop."));
  });
  img.src = url;
  return loaded;
}

async function cropNormalizedRegionToBlob(page, norm) {
  if (!page || !norm) return null;
  const img = await loadImageElement(page.objectUrl);
  const left = Math.min(norm.x0, norm.x1);
  const top = Math.min(norm.y0, norm.y1);
  const width = Math.abs(norm.x1 - norm.x0);
  const height = Math.abs(norm.y1 - norm.y0);
  let sx = Math.round(left * img.naturalWidth);
  let sy = Math.round(top * img.naturalHeight);
  let sw = Math.round(width * img.naturalWidth);
  let sh = Math.round(height * img.naturalHeight);
  sx = Math.max(0, Math.min(sx, Math.max(0, img.naturalWidth - 1)));
  sy = Math.max(0, Math.min(sy, Math.max(0, img.naturalHeight - 1)));
  sw = Math.max(1, Math.min(sw, img.naturalWidth - sx));
  sh = Math.max(1, Math.min(sh, img.naturalHeight - sy));
  const canvas = document.createElement("canvas");
  canvas.width = sw;
  canvas.height = sh;
  canvas.getContext("2d").drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh);
  return canvasToPngBlob(canvas);
}

async function cropSelectionToBlob() {
  const page = getSelectedPage();
  if (!page || !cropDrag) return null;
  return cropNormalizedRegionToBlob(page, cropDrag);
}

function rememberPendingCrop() {
  const page = getSelectedPage();
  if (!page) return;
  page.pendingCrop = cropDrag && cropSelectionIsBigEnough()
    ? { x0: cropDrag.x0, y0: cropDrag.y0, x1: cropDrag.x1, y1: cropDrag.y1 }
    : null;
}

async function runPool(items, limit, worker) {
  if (!items.length) return;
  let cursor = 0;
  const size = Math.max(1, Math.min(limit, items.length));
  const runners = Array.from({ length: size }, async () => {
    while (cursor < items.length) {
      const index = cursor;
      cursor += 1;
      await worker(items[index], index);
    }
  });
  await Promise.all(runners);
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

  cropDrag = page.pendingCrop
    ? { x0: page.pendingCrop.x0, y0: page.pendingCrop.y0, x1: page.pendingCrop.x1, y1: page.pendingCrop.y1 }
    : null;
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
  const isolateAllBtn = el("digitizeIsolateAllPagesBtn");
  if (buildBtn) buildBtn.disabled = !hasPages;
  if (runOcrBtn) runOcrBtn.disabled = !hasPages;
  if (isolateAllBtn) isolateAllBtn.disabled = !hasPages;
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
    page.pendingCrop = null;
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

function objectExtractMessages() {
  return {
    empty: t("javascriptStrings.objectExtract.empty"),
    tooBroad: t("javascriptStrings.objectExtract.tooBroad"),
    outside: t("javascriptStrings.objectExtract.outside"),
    failed: t("javascriptStrings.objectExtract.failed")
  };
}

async function extractAtPoint(point) {
  const page = getSelectedPage();
  if (!page || !point || extractBusy) return;
  if (page.crops.length >= MAX_CROPS_PER_PAGE) {
    setDigitizeStatus(t("javascriptStrings.digitize.maxCrops"), true);
    return;
  }

  let box = null;
  if (cropDrag && cropSelectionIsBigEnough()) {
    const left = Math.min(cropDrag.x0, cropDrag.x1);
    const top = Math.min(cropDrag.y0, cropDrag.y1);
    const right = Math.max(cropDrag.x0, cropDrag.x1);
    const bottom = Math.max(cropDrag.y0, cropDrag.y1);
    const inside = point.x >= left && point.x <= right && point.y >= top && point.y <= bottom;
    if (inside) box = { x0: left, y0: top, x1: right, y1: bottom };
  }

  const pageId = page.id;
  extractBusy = true;
  const isolateBtn = el("digitizeIsolateCropBtn");
  if (isolateBtn) isolateBtn.disabled = true;
  setDigitizeStatus(t("javascriptStrings.objectExtract.extracting"));

  try {
    const img = await loadImageElement(page.objectUrl);
    const cutout = await extractObjectFromImage(img, point, {
      box,
      onStage: (stage) => {
        if (stage === "loading-model") setDigitizeStatus(t("javascriptStrings.objectExtract.loadingModel"));
      }
    });
    const finished = deps?.decorateCutout ? await deps.decorateCutout(cutout) : cutout;
    const target = pages.find((item) => item.id === pageId);
    if (!target) return;
    if (target.crops.length >= MAX_CROPS_PER_PAGE) {
      setDigitizeStatus(t("javascriptStrings.digitize.maxCrops"), true);
      return;
    }
    const file = new File(
      [finished],
      `${target.name.replace(/\.[^/.]+$/, "")}-object${target.crops.length + 1}.png`,
      { type: "image/png" }
    );
    target.crops.push({
      id: newId("crop"),
      name: file.name,
      file,
      objectUrl: URL.createObjectURL(file)
    });
    if (selectedPageId === pageId) renderCropsList();
    setDigitizeStatus(t("javascriptStrings.digitize.objectExtracted"));
  } catch (err) {
    console.error(err);
    setDigitizeStatus(objectExtractUserMessage(err, objectExtractMessages()), true);
  } finally {
    extractBusy = false;
    if (isolateBtn) isolateBtn.disabled = false;
  }
}

async function isolateAllPageImages() {
  if (!pages.length) return;
  if (!deps?.isolateBlob) {
    setDigitizeStatus(t("javascriptStrings.digitize.backgroundRemoverNotReady"), true);
    return;
  }
  const hasCrops = pages.some((p) => (p.crops || []).length);
  if (hasCrops && !window.confirm(t("javascriptStrings.digitize.confirmIsolateAllReplace"))) return;

  const isolateAllBtn = el("digitizeIsolateAllPagesBtn");
  const isolateBtn = el("digitizeIsolateCropBtn");
  if (isolateAllBtn) isolateAllBtn.disabled = true;
  if (isolateBtn) isolateBtn.disabled = true;

  let okCount = 0;
  let failCount = 0;
  let finished = 0;
  setDigitizeStatus(t("javascriptStrings.digitize.isolatingAllPages", { current: 0, total: pages.length }));

  try {
    await runPool(pages, Math.min(4, pages.length), async (page) => {
      try {
        const source = page.pendingCrop && cropSelectionIsBigEnough(page.pendingCrop)
          ? await cropNormalizedRegionToBlob(page, page.pendingCrop)
          : page.blob;
        if (!source) throw new Error(t("javascriptStrings.digitize.drawCropFirst"));
        const isolated = await deps.isolateBlob(source);
        const file = new File(
          [isolated],
          `${page.name.replace(/\.[^/.]+$/, "")}-isolated.png`,
          { type: "image/png" }
        );
        for (const crop of page.crops || []) {
          if (crop.objectUrl) URL.revokeObjectURL(crop.objectUrl);
        }
        page.crops = [{
          id: newId("crop"),
          name: file.name,
          file,
          objectUrl: URL.createObjectURL(file)
        }];
        page.pendingCrop = null;
        if (page.id === selectedPageId) cropDrag = null;
        okCount += 1;
      } catch (err) {
        console.error(err);
        failCount += 1;
      } finally {
        finished += 1;
        setDigitizeStatus(t("javascriptStrings.digitize.isolatingAllPages", {
          current: finished,
          total: pages.length
        }));
      }
    });
    drawCropOverlay();
    renderCropsList();
    if (failCount) {
      setDigitizeStatus(
        t("javascriptStrings.digitize.allPagesIsolatedWithErrors", { ok: okCount, n: failCount }),
        true
      );
    } else {
      setDigitizeStatus(t("javascriptStrings.digitize.allPagesIsolated", { n: okCount }));
    }
  } finally {
    if (isolateBtn) isolateBtn.disabled = false;
    renderAll();
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

function getDigitizeMode() {
  const checked = document.querySelector('input[name="digitizeBookType"]:checked');
  return checked?.value === "nook" ? "nook" : "normal";
}

function applyDigitizeMode() {
  const isNook = getDigitizeMode() === "nook";
  const normal = el("digitizeNormalMode");
  const nook = el("digitizeNookMode");
  const hintNormal = el("digitizeHintNormal");
  const hintNook = el("digitizeHintNook");
  if (normal) normal.hidden = isNook;
  if (nook) nook.hidden = !isNook;
  if (hintNormal) hintNormal.hidden = isNook;
  if (hintNook) hintNook.hidden = !isNook;
  if (isNook) {
    setDigitizeStatus(t("digitizeBook.nookInitialStatus"));
  } else if (!pages.length) {
    setDigitizeStatus(t("digitizeBook.initialStatus"));
  }
}

function setNookImportBusy(busy) {
  const input = el("digitizeNookPdfInput");
  const isolate = el("digitizeNookIsolate");
  const isolateAll = el("digitizeNookIsolateAll");
  document.querySelectorAll('input[name="digitizeBookType"]').forEach((radio) => {
    radio.disabled = busy;
  });
  if (input) input.disabled = busy;
  if (isolate) isolate.disabled = busy;
  if (isolateAll) isolateAll.disabled = busy;
}

async function renderNookPhotoFile(pdf, pageNumber, baseName) {
  const page = await pdf.getPage(pageNumber);
  const canvas = cropCanvasToContent(await renderPdfPageToCanvas(page));
  const blob = await canvasToPngBlob(canvas);
  return new File([blob], `${baseName}-p${pageNumber}.png`, { type: "image/png" });
}

async function importCviBookNookPdf(file) {
  if (!file || !deps?.rebuildSpreads) return;
  const lower = (file.name || "").toLowerCase();
  if (file.type !== "application/pdf" && !lower.endsWith(".pdf")) {
    setDigitizeStatus(t("javascriptStrings.digitize.nookNoPdf"), true);
    return;
  }

  setNookImportBusy(true);
  setDigitizeStatus(t("javascriptStrings.digitize.nookLoading"));

  try {
    const pdfjs = await getPdfjs();
    const data = await file.arrayBuffer();
    const pdf = await pdfjs.getDocument({ data }).promise;
    const pageCount = Math.min(pdf.numPages, MAX_PAGES);
    const truncated = pdf.numPages > MAX_PAGES;

    const textPages = [];
    for (let i = 1; i <= pageCount; i += 1) {
      const page = await pdf.getPage(i);
      textPages.push({
        pageNumber: i,
        text: await getPdfPageText(page)
      });
    }

    const parsed = pairNookSpreads(textPages);
    if (!parsed.spreads.length) {
      setDigitizeStatus(t("javascriptStrings.digitize.nookNotFormat"), true);
      return;
    }

    const ok = window.confirm(t("javascriptStrings.digitize.nookConfirmReplace"));
    if (!ok) {
      setDigitizeStatus(t("digitizeBook.nookInitialStatus"));
      return;
    }

    const isolateAll = Boolean(el("digitizeNookIsolateAll")?.checked);
    const isolate = isolateAll || Boolean(el("digitizeNookIsolate")?.checked);
    const baseName = file.name.replace(/\.[^/.]+$/, "") || "nook";
    const photoIndexes = parsed.spreads
      .map((s) => s.photoPageNumber)
      .filter((n) => typeof n === "number");
    let photoDone = 0;

    const jobs = [];
    for (const spread of parsed.spreads) {
      /** @type {File|null} */
      let imageFile = null;
      if (spread.photoPageNumber) {
        photoDone += 1;
        setDigitizeStatus(
          t("javascriptStrings.digitize.nookProgress", { current: photoDone, total: photoIndexes.length || 1 })
        );
        try {
          imageFile = await renderNookPhotoFile(pdf, spread.photoPageNumber, baseName);
        } catch (err) {
          console.error(err);
        }
      }
      jobs.push({ spread, imageFile });
    }

    if (isolate && deps?.isolateBlob) {
      const withPhotos = jobs.filter((job) => job.imageFile);
      let isolatedDone = 0;
      const concurrency = isolateAll ? Math.min(4, withPhotos.length || 1) : 1;
      await runPool(withPhotos, concurrency, async (job) => {
        try {
          const isolated = await deps.isolateBlob(job.imageFile);
          job.imageFile = new File([isolated], job.imageFile.name, { type: "image/png" });
        } catch (err) {
          console.error(err);
        } finally {
          isolatedDone += 1;
          const progressKey = isolateAll
            ? "javascriptStrings.digitize.nookIsolatingAll"
            : "javascriptStrings.digitize.nookIsolating";
          setDigitizeStatus(t(progressKey, {
            current: isolatedDone,
            total: withPhotos.length || 1
          }));
        }
      });
    }

    const spreads = jobs.map((job) => ({
      storyText: job.spread.storyText,
      oddText: job.spread.oddText,
      salientFeatures: job.spread.salientFeatures,
      imageFiles: job.imageFile ? [job.imageFile] : []
    }));

    if (parsed.title) deps.setBookTitle?.(parsed.title);
    deps.rebuildSpreads(spreads);
    const doneMsg = t("javascriptStrings.digitize.nookImported", { n: spreads.length });
    setDigitizeStatus(truncated
      ? `${doneMsg} ${t("javascriptStrings.digitize.truncatedPages", { max: MAX_PAGES, total: pdf.numPages })}`
      : doneMsg);
    deps.setStatus?.(doneMsg);
  } catch (err) {
    console.error(err);
    setDigitizeStatus(
      `${t("javascriptStrings.digitize.loadFailed")}${err.message || t("javascriptStrings.errors.unknownFallback")}`,
      true
    );
  } finally {
    setNookImportBusy(false);
    const nookInput = el("digitizeNookPdfInput");
    if (nookInput) nookInput.value = "";
  }
}

function initCropInteraction() {
  const canvas = el("digitizeCropCanvas");
  if (!canvas) return;
  let press = null;

  const beginDrag = (pt) => {
    cropDragging = true;
    cropDrag = { x0: pt.x, y0: pt.y, x1: pt.x, y1: pt.y };
    rememberPendingCrop();
    drawCropOverlay();
  };

  const onDown = (e) => {
    if (!getSelectedPage() || extractBusy) return;
    if (e.button != null && e.button !== 0) return;
    syncCropOverlaySize();
    const pt = imageNormFromEvent(e);
    if (!pt) return;
    e.preventDefault();
    try {
      if (canvas.setPointerCapture) canvas.setPointerCapture(e.pointerId);
    } catch (err) {
      /* The pointer can already be gone; later move events still update the box. */
    }
    press = { x: pt.x, y: pt.y, dragged: false };
    if (!clickExtractEnabled()) beginDrag(pt);
  };
  const onMove = (e) => {
    if (press && !press.dragged) {
      const pt = imageNormFromEvent(e);
      const img = el("digitizePageImage");
      const rect = img?.getBoundingClientRect();
      if (pt && rect) {
        const dx = (pt.x - press.x) * rect.width;
        const dy = (pt.y - press.y) * rect.height;
        if (Math.hypot(dx, dy) > 6) {
          press.dragged = true;
          if (clickExtractEnabled() && !cropDragging) beginDrag(press);
        }
      }
    }
    if (!cropDragging || !cropDrag) return;
    const pt = imageNormFromEvent(e);
    if (!pt) return;
    e.preventDefault();
    cropDrag.x1 = pt.x;
    cropDrag.y1 = pt.y;
    rememberPendingCrop();
    drawCropOverlay();
  };
  const onUp = (e) => {
    const wasClick = Boolean(press && !press.dragged && clickExtractEnabled());
    const clickPoint = press ? { x: press.x, y: press.y } : null;
    press = null;
    if (wasClick && clickPoint) {
      cropDragging = false;
      extractAtPoint(clickPoint);
      return;
    }
    if (!cropDragging) return;
    cropDragging = false;
    const pt = imageNormFromEvent(e);
    if (pt && cropDrag) {
      cropDrag.x1 = pt.x;
      cropDrag.y1 = pt.y;
    }
    if (!cropSelectionIsBigEnough()) cropDrag = null;
    rememberPendingCrop();
    drawCropOverlay();
  };

  canvas.addEventListener("pointerdown", onDown);
  canvas.addEventListener("pointermove", onMove);
  canvas.addEventListener("pointerup", onUp);
  canvas.addEventListener("pointercancel", onUp);
  window.addEventListener("pointermove", onMove);
  window.addEventListener("pointerup", onUp);
  window.addEventListener("pointercancel", onUp);
  window.addEventListener("resize", () => syncCropOverlaySize());

  const frame = canvas.parentElement;
  if (typeof ResizeObserver !== "undefined" && frame && !cropResizeObserver) {
    cropResizeObserver = new ResizeObserver(() => syncCropOverlaySize());
    cropResizeObserver.observe(frame);
  }
}

/**
 * @param {{
 *   isolateBlob: (blob: Blob) => Promise<Blob>,
 *   decorateCutout?: (blob: Blob) => Promise<Blob>,
 *   rebuildSpreads: (spreads: { storyText: string, oddText: string, salientFeatures?: string, imageFiles?: File[] }[]) => void,
 *   setStatus: (text: string, isError?: boolean) => void,
 *   ensureCompatibleImage: (file: File) => Promise<File>,
 *   setBookTitle?: (title: string) => void,
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
  const isolateAllPagesBtn = el("digitizeIsolateAllPagesBtn");
  const clearCropBtn = el("digitizeClearCropBtn");
  const buildBtn = el("digitizeBuildBookBtn");
  const clearBtn = el("digitizeClearBtn");
  const textArea = el("digitizeOcrText");
  const modelSelect = el("digitizeOcrModel");
  const nookPdfInput = el("digitizeNookPdfInput");
  const copyPageBtn = el("digitizeCopyPageTextBtn");
  const copyAiBtn = el("digitizeCopyAiReviewBtn");
  const applyAiBtn = el("digitizeApplyAiReviewBtn");

  document.querySelectorAll('input[name="digitizeBookType"]').forEach((radio) => {
    radio.addEventListener("change", () => applyDigitizeMode());
  });

  if (nookPdfInput) {
    nookPdfInput.addEventListener("change", async () => {
      const file = nookPdfInput.files && nookPdfInput.files[0];
      if (!file) return;
      await importCviBookNookPdf(file);
    });
  }

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
  const clickExtract = el("digitizeClickExtract");
  if (clickExtract) {
    clickExtract.addEventListener("change", () => syncClickExtractFrame());
  }
  syncClickExtractFrame();

  if (isolateBtn) isolateBtn.addEventListener("click", () => isolateCurrentCrop());
  if (isolateAllPagesBtn) isolateAllPagesBtn.addEventListener("click", () => isolateAllPageImages());
  if (clearCropBtn) {
    clearCropBtn.addEventListener("click", () => {
      cropDrag = null;
      const page = getSelectedPage();
      if (page) page.pendingCrop = null;
      drawCropOverlay();
    });
  }

  const nookIsolate = el("digitizeNookIsolate");
  const nookIsolateAll = el("digitizeNookIsolateAll");
  if (nookIsolate && nookIsolateAll) {
    nookIsolate.addEventListener("change", () => {
      if (!nookIsolate.checked) nookIsolateAll.checked = false;
    });
    nookIsolateAll.addEventListener("change", () => {
      if (nookIsolateAll.checked) nookIsolate.checked = true;
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
  applyDomTranslations(panel);
  applyDigitizeMode();
  renderAll();
}

export function setDigitizeMode(mode) {
  const id = mode === "nook" ? "digitizeTypeNook" : "digitizeTypeNormal";
  const radio = el(id);
  if (!radio) return;
  radio.checked = true;
  applyDigitizeMode();
}

export function setDigitizeTypeLocked(locked) {
  const fieldset = document.querySelector("#digitizeBookPanel .digitize-book-type");
  if (fieldset) fieldset.hidden = Boolean(locked);
}

export function refreshDigitizeLocale() {
  const panel = el("digitizeBookPanel");
  if (panel) applyDomTranslations(panel);
  applyDigitizeMode();
  renderAll();
}
