/**
 * Click-to-extract one character or object from a page image.
 *
 * Flat illustrations (the usual children's-book case) are cut out with a color
 * select on this device, so they do not wait on a model download. Shaded
 * pictures and photos fall through to MediaPipe's MagicTouch interactive
 * segmenter, which also stays on device after the model download.
 */

import { FilesetResolver, InteractiveSegmenter } from "https://cdn.jsdelivr.net/npm/@mediapipe/tasks-vision@0.10.32/vision_bundle.mjs";

const VISION_WASM_URL = "https://cdn.jsdelivr.net/npm/@mediapipe/tasks-vision@0.10.32/wasm";
const MAGIC_TOUCH_URL = "https://storage.googleapis.com/mediapipe-models/interactive_segmenter/magic_touch/float32/1/magic_touch.tflite";

const MAX_SIDE = 1280;
const FLAT_COLOR_STD = 26;
const FLOOD_TOLERANCE = 34;
const FLOOD_TOLERANCE_RETRY = 50;
const MIN_OBJECT_PIXELS = 48;
const MAX_OBJECT_RATIO = 0.9;

/** @typedef {"empty" | "too-broad" | "outside" | "failed"} ObjectExtractFailure */

export class ObjectExtractError extends Error {
  /**
   * @param {"empty" | "too-broad" | "outside"} code
   */
  constructor(code) {
    super(code);
    this.name = "ObjectExtractError";
    this.code = code;
  }
}

/**
 * @param {unknown} err
 * @param {{ empty: string, tooBroad: string, outside: string, failed: string }} messages
 */
export function objectExtractUserMessage(err, messages) {
  const code = err instanceof ObjectExtractError ? err.code : "failed";
  switch (code) {
    case "empty":
      return messages.empty;
    case "too-broad":
      return messages.tooBroad;
    case "outside":
      return messages.outside;
    case "failed": {
      const detail = err instanceof Error && !(err instanceof ObjectExtractError) ? err.message : "";
      return `${messages.failed}${detail}`;
    }
    default: {
      const unexpected = code;
      return `${messages.failed}${String(unexpected)}`;
    }
  }
}

/**
 * Map a click to normalized image coordinates, including object-fit letterboxing.
 * Returns null when the click misses the painted pixels.
 * @param {HTMLImageElement} img
 * @param {number} clientX
 * @param {number} clientY
 * @returns {{ x: number, y: number } | null}
 */
export function normalizedPointOnImage(img, clientX, clientY) {
  const rect = img.getBoundingClientRect();
  const nw = img.naturalWidth || img.width;
  const nh = img.naturalHeight || img.height;
  if (rect.width < 2 || rect.height < 2 || !nw || !nh) return null;

  let drawW = rect.width;
  let drawH = rect.height;
  let offsetX = 0;
  let offsetY = 0;
  const fit = getComputedStyle(img).objectFit;
  const elementAspect = rect.width / rect.height;
  const imageAspect = nw / nh;
  if (fit === "contain" || fit === "scale-down") {
    if (imageAspect > elementAspect) {
      drawH = rect.width / imageAspect;
      offsetY = (rect.height - drawH) / 2;
    } else {
      drawW = rect.height * imageAspect;
      offsetX = (rect.width - drawW) / 2;
    }
  } else if (fit === "cover") {
    if (imageAspect > elementAspect) {
      drawW = rect.height * imageAspect;
      offsetX = (rect.width - drawW) / 2;
    } else {
      drawH = rect.width / imageAspect;
      offsetY = (rect.height - drawH) / 2;
    }
  }

  const x = (clientX - rect.left - offsetX) / drawW;
  const y = (clientY - rect.top - offsetY) / drawH;
  if (x < 0 || y < 0 || x > 1 || y > 1) return null;
  return { x, y };
}

let segmenterPromise = null;

function getSegmenter() {
  if (!segmenterPromise) {
    segmenterPromise = createSegmenter().catch((err) => {
      segmenterPromise = null;
      throw err;
    });
  }
  return segmenterPromise;
}

async function createSegmenter() {
  const vision = await FilesetResolver.forVisionTasks(VISION_WASM_URL);
  return InteractiveSegmenter.createFromOptions(vision, {
    baseOptions: {
      modelAssetPath: MAGIC_TOUCH_URL,
      delegate: "CPU"
    },
    outputConfidenceMasks: true,
    outputCategoryMask: false
  });
}

/**
 * @param {Blob | HTMLImageElement | HTMLCanvasElement} source
 * @param {{ x: number, y: number }} point normalized 0–1
 * @param {{ box?: { x0: number, y0: number, x1: number, y1: number } | null, onStage?: (stage: "extracting" | "loading-model") => void }} [options]
 * @returns {Promise<Blob>}
 */
export async function extractObjectFromImage(source, point, options = {}) {
  if (!point || point.x < 0 || point.y < 0 || point.x > 1 || point.y > 1) {
    throw new ObjectExtractError("outside");
  }
  options.onStage?.("extracting");
  const working = await drawWorkingCanvas(source);
  const { canvas, width, height } = working;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  if (!ctx) throw new Error("Could not read the image.");
  const image = ctx.getImageData(0, 0, width, height);
  const sx = clampPixel(point.x, width);
  const sy = clampPixel(point.y, height);
  const box = normalizedBoxToPixels(options.box, width, height);
  const seed = nearestOpaquePixel(image, width, height, sx, sy, box);
  if (!seed) throw new ObjectExtractError("empty");

  const std = localColorStd(image, width, height, seed.x, seed.y);
  let alpha = null;
  if (std <= FLAT_COLOR_STD) {
    alpha = floodAlpha(image, width, height, seed, box, FLOOD_TOLERANCE);
    if (!maskIsUseful(alpha, width, height, box)) {
      alpha = floodAlpha(image, width, height, seed, box, FLOOD_TOLERANCE_RETRY);
    }
  }

  if (!alpha || !maskIsUseful(alpha, width, height, box)) {
    if (alpha && std <= FLAT_COLOR_STD && classifyMask(alpha, width, height, box) === "too-broad") {
      throw new ObjectExtractError("too-broad");
    }
    let modelError = null;
    try {
      alpha = await magicTouchAlpha(canvas, point, width, height, box, options.onStage);
    } catch (err) {
      modelError = err;
      alpha = null;
    }
    if (!alpha || !maskIsUseful(alpha, width, height, box)) {
      const fallback = floodAlpha(image, width, height, seed, box, FLOOD_TOLERANCE_RETRY);
      if (maskIsUseful(fallback, width, height, box)) {
        alpha = fallback;
      } else if (modelError) {
        throw modelError;
      }
    }
  }

  assertUseful(alpha, width, height, box);
  return cutoutBlob(image, alpha, width, height);
}

function clampPixel(unit, size) {
  return Math.min(size - 1, Math.max(0, Math.round(unit * (size - 1))));
}

/**
 * @param {Blob | HTMLImageElement | HTMLCanvasElement} source
 */
async function drawWorkingCanvas(source) {
  const image = await sourceToDrawable(source);
  const nw = image instanceof HTMLCanvasElement ? image.width : image.naturalWidth || image.width;
  const nh = image instanceof HTMLCanvasElement ? image.height : image.naturalHeight || image.height;
  if (!nw || !nh) throw new Error("Could not read the image.");
  const scale = Math.min(1, MAX_SIDE / Math.max(nw, nh));
  const width = Math.max(1, Math.round(nw * scale));
  const height = Math.max(1, Math.round(nh * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Could not read the image.");
  ctx.imageSmoothingEnabled = scale < 1;
  ctx.drawImage(image, 0, 0, width, height);
  return { canvas, width, height };
}

/**
 * @param {Blob | HTMLImageElement | HTMLCanvasElement} source
 * @returns {Promise<CanvasImageSource>}
 */
async function sourceToDrawable(source) {
  if (source instanceof HTMLCanvasElement) return source;
  if (source instanceof Blob) return loadImageElement(source);
  if (source instanceof HTMLImageElement) {
    if (!source.complete || !source.naturalWidth) {
      await source.decode();
    }
    return source;
  }
  throw new Error("Unsupported image.");
}

function loadImageElement(blob) {
  const url = URL.createObjectURL(blob);
  const img = new Image();
  return new Promise((resolve, reject) => {
    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve(img);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Could not read the image."));
    };
    img.src = url;
  });
}

/**
 * @param {{ x0: number, y0: number, x1: number, y1: number } | null | undefined} box
 */
function normalizedBoxToPixels(box, width, height) {
  if (!box) return null;
  const x0 = clampPixel(Math.min(box.x0, box.x1), width);
  const y0 = clampPixel(Math.min(box.y0, box.y1), height);
  const x1 = clampPixel(Math.max(box.x0, box.x1), width);
  const y1 = clampPixel(Math.max(box.y0, box.y1), height);
  if (x1 - x0 < 2 || y1 - y0 < 2) return null;
  return { x0, y0, x1: x1 + 1, y1: y1 + 1 };
}

function nearestOpaquePixel(image, width, height, x, y, box) {
  const x0 = box ? box.x0 : 0;
  const y0 = box ? box.y0 : 0;
  const x1 = box ? box.x1 : width;
  const y1 = box ? box.y1 : height;
  const data = image.data;
  for (let radius = 0; radius <= 8; radius += 1) {
    for (let dy = -radius; dy <= radius; dy += 1) {
      for (let dx = -radius; dx <= radius; dx += 1) {
        if (Math.max(Math.abs(dx), Math.abs(dy)) !== radius) continue;
        const px = x + dx;
        const py = y + dy;
        if (px < x0 || py < y0 || px >= x1 || py >= y1) continue;
        if (data[(py * width + px) * 4 + 3] > 16) return { x: px, y: py };
      }
    }
  }
  return null;
}

function localColorStd(image, width, height, x, y) {
  const data = image.data;
  let n = 0;
  let sr = 0;
  let sg = 0;
  let sb = 0;
  let sr2 = 0;
  let sg2 = 0;
  let sb2 = 0;
  for (let dy = -4; dy <= 4; dy += 1) {
    for (let dx = -4; dx <= 4; dx += 1) {
      const px = Math.min(width - 1, Math.max(0, x + dx));
      const py = Math.min(height - 1, Math.max(0, y + dy));
      const i = (py * width + px) * 4;
      const r = data[i];
      const g = data[i + 1];
      const b = data[i + 2];
      sr += r;
      sg += g;
      sb += b;
      sr2 += r * r;
      sg2 += g * g;
      sb2 += b * b;
      n += 1;
    }
  }
  const vr = sr2 / n - (sr / n) * (sr / n);
  const vg = sg2 / n - (sg / n) * (sg / n);
  const vb = sb2 / n - (sb / n) * (sb / n);
  return Math.sqrt(Math.max(0, vr + vg + vb));
}

function floodAlpha(image, width, height, seed, box, tolerance) {
  const data = image.data;
  const x0 = box ? box.x0 : 0;
  const y0 = box ? box.y0 : 0;
  const x1 = box ? box.x1 : width;
  const y1 = box ? box.y1 : height;
  const region = Math.max(1, (x1 - x0) * (y1 - y0));
  const stopAt = Math.floor(region * MAX_OBJECT_RATIO) + 1;
  const seedIndex = (seed.y * width + seed.x) * 4;
  const sr = data[seedIndex];
  const sg = data[seedIndex + 1];
  const sb = data[seedIndex + 2];
  const tol2 = tolerance * tolerance;
  const seen = new Uint8Array(width * height);
  const alpha = new Float32Array(width * height);
  const qx = new Int32Array(width * height);
  const qy = new Int32Array(width * height);
  let head = 0;
  let tail = 0;
  let count = 0;
  const enqueue = (nx, ny) => {
    const ni = ny * width + nx;
    if (seen[ni]) return;
    seen[ni] = 1;
    qx[tail] = nx;
    qy[tail] = ny;
    tail += 1;
  };
  enqueue(seed.x, seed.y);

  while (head < tail) {
    const x = qx[head];
    const y = qy[head];
    head += 1;
    const i = y * width + x;
    const p = i * 4;
    const dr = data[p] - sr;
    const dg = data[p + 1] - sg;
    const db = data[p + 2] - sb;
    if (dr * dr + dg * dg + db * db > tol2) continue;
    alpha[i] = 1;
    count += 1;
    if (count > stopAt) {
      alpha.fill(1);
      return alpha;
    }
    if (x + 1 < x1) enqueue(x + 1, y);
    if (x - 1 >= x0) enqueue(x - 1, y);
    if (y + 1 < y1) enqueue(x, y + 1);
    if (y - 1 >= y0) enqueue(x, y - 1);
  }

  return fillPinholes(alpha, width, height, box);
}

function fillPinholes(alpha, width, height, box) {
  const x0 = box ? box.x0 : 0;
  const y0 = box ? box.y0 : 0;
  const x1 = box ? box.x1 : width;
  const y1 = box ? box.y1 : height;
  const out = new Float32Array(alpha);
  for (let y = Math.max(1, y0); y < Math.min(height - 1, y1); y += 1) {
    for (let x = Math.max(1, x0); x < Math.min(width - 1, x1); x += 1) {
      const i = y * width + x;
      if (alpha[i] > 0) continue;
      const neighbors = alpha[i - 1] + alpha[i + 1] + alpha[i - width] + alpha[i + width];
      if (neighbors >= 3) out[i] = 1;
    }
  }
  return out;
}

async function magicTouchAlpha(canvas, point, width, height, box, onStage) {
  onStage?.("loading-model");
  const segmenter = await getSegmenter();
  onStage?.("extracting");
  const result = segmenter.segment(canvas, {
    keypoint: { x: point.x, y: point.y }
  });
  try {
    const mask = result.confidenceMasks && result.confidenceMasks[0];
    if (!mask) throw new Error("Object extractor returned no mask.");
    const data = mask.getAsFloat32Array();
    const mw = mask.width;
    const mh = mask.height;
    if (!mw || !mh || data.length < mw * mh) throw new Error("Object extractor returned an empty mask.");
    const alpha = new Float32Array(width * height);
    const x0 = box ? box.x0 : 0;
    const y0 = box ? box.y0 : 0;
    const x1 = box ? box.x1 : width;
    const y1 = box ? box.y1 : height;
    for (let y = y0; y < y1; y += 1) {
      const my = Math.min(mh - 1, Math.max(0, Math.round(((y + 0.5) * mh) / height - 0.5)));
      for (let x = x0; x < x1; x += 1) {
        const mx = Math.min(mw - 1, Math.max(0, Math.round(((x + 0.5) * mw) / width - 0.5)));
        alpha[y * width + x] = confidenceToAlpha(data[my * mw + mx]);
      }
    }
    return alpha;
  } finally {
    if (typeof result.close === "function") result.close();
  }
}

function confidenceToAlpha(confidence) {
  if (confidence <= 0.35) return 0;
  if (confidence >= 0.62) return 1;
  return (confidence - 0.35) / 0.27;
}

/** Background clicks flood to every edge. A character usually does not. */
function touchesAllEdges(alpha, width, height) {
  let top = false;
  let bottom = false;
  let left = false;
  let right = false;
  for (let x = 0; x < width; x += 1) {
    if (alpha[x] > 0.2) top = true;
    if (alpha[(height - 1) * width + x] > 0.2) bottom = true;
  }
  for (let y = 0; y < height; y += 1) {
    if (alpha[y * width] > 0.2) left = true;
    if (alpha[y * width + width - 1] > 0.2) right = true;
  }
  return top && bottom && left && right;
}

function classifyMask(alpha, width, height, box) {
  const stats = maskStats(alpha, width, height, box);
  const ratio = stats.count / stats.region;
  if (stats.count < MIN_OBJECT_PIXELS || ratio < 0.002) return "empty";
  if (ratio > MAX_OBJECT_RATIO) return "too-broad";
  if (!box && ratio > 0.5 && touchesAllEdges(alpha, width, height)) return "too-broad";
  return "ok";
}

function maskIsUseful(alpha, width, height, box) {
  return classifyMask(alpha, width, height, box) === "ok";
}

function assertUseful(alpha, width, height, box) {
  const kind = classifyMask(alpha, width, height, box);
  if (kind === "ok") return;
  throw new ObjectExtractError(kind);
}

function maskStats(alpha, width, height, box) {
  const x0 = box ? box.x0 : 0;
  const y0 = box ? box.y0 : 0;
  const x1 = box ? box.x1 : width;
  const y1 = box ? box.y1 : height;
  let count = 0;
  for (let y = y0; y < y1; y += 1) {
    const row = y * width;
    for (let x = x0; x < x1; x += 1) {
      if (alpha[row + x] > 0.2) count += 1;
    }
  }
  return { count, region: Math.max(1, (x1 - x0) * (y1 - y0)) };
}

function cutoutBlob(image, alpha, width, height) {
  const pixels = image.data;
  let minX = width;
  let minY = height;
  let maxX = -1;
  let maxY = -1;
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const i = y * width + x;
      const a = Math.round(Math.max(0, Math.min(1, alpha[i])) * 255);
      const p = i * 4;
      pixels[p + 3] = Math.min(pixels[p + 3], a);
      if (a > 16) {
        if (x < minX) minX = x;
        if (y < minY) minY = y;
        if (x > maxX) maxX = x;
        if (y > maxY) maxY = y;
      }
    }
  }
  if (maxX < 0) throw new ObjectExtractError("empty");
  const pad = 2;
  minX = Math.max(0, minX - pad);
  minY = Math.max(0, minY - pad);
  maxX = Math.min(width - 1, maxX + pad);
  maxY = Math.min(height - 1, maxY + pad);
  const cropW = maxX - minX + 1;
  const cropH = maxY - minY + 1;
  const full = document.createElement("canvas");
  full.width = width;
  full.height = height;
  const fullCtx = full.getContext("2d");
  if (!fullCtx) throw new Error("Could not build the cutout.");
  fullCtx.putImageData(image, 0, 0);
  const out = document.createElement("canvas");
  out.width = cropW;
  out.height = cropH;
  const outCtx = out.getContext("2d");
  if (!outCtx) throw new Error("Could not build the cutout.");
  outCtx.drawImage(full, minX, minY, cropW, cropH, 0, 0, cropW, cropH);
  return new Promise((resolve, reject) => {
    out.toBlob((blob) => {
      if (!blob) {
        reject(new Error("Could not build the cutout."));
        return;
      }
      resolve(blob);
    }, "image/png");
  });
}
