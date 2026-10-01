/**
 * Background removal for the Image Isolator and Digitize Book.
 *
 * BiRefNet-lite at 512×512 is the in-browser matte that stays accurate on both
 * photos and illustrated objects. It runs on WebGPU when the GPU supports
 * 16-bit shaders, and on WASM otherwise. IMG.LY's ISNet remover is the fallback
 * if that model cannot run. The outline is a uniform ring around every
 * foreground object, including when several objects share one cutout.
 *
 * Transformers.js is imported only when removal starts so the page does not
 * download that library on open.
 */

const TRANSFORMERS_URL = "https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.8.1/dist/transformers.web.js";
const MODEL_ID = "studioludens/birefnet-lite-512";
const IMGLY_URL = "https://cdn.jsdelivr.net/npm/@imgly/background-removal/+esm";

const FOREGROUND_ALPHA = 24;
const HAZE_ALPHA = 12;

let transformersPromise = null;
let sessionPromise = null;
let birefnetDisabled = false;
let imglyPromise = null;
let removalQueue = Promise.resolve();

function enqueueRemoval(task) {
  const run = removalQueue.then(task, task);
  removalQueue = run.then(() => undefined, () => undefined);
  return run;
}

function loadTransformers() {
  if (!transformersPromise) {
    transformersPromise = import(TRANSFORMERS_URL).then((mod) => {
      if (mod.env) {
        mod.env.allowLocalModels = false;
        if (mod.env.backends?.onnx?.wasm) {
          mod.env.backends.onnx.wasm.numThreads = 1;
        }
      }
      return mod;
    }).catch((err) => {
      transformersPromise = null;
      throw err;
    });
  }
  return transformersPromise;
}

async function preferredBackend() {
  if (typeof navigator !== "undefined" && navigator.gpu && typeof navigator.gpu.requestAdapter === "function") {
    try {
      const adapter = await navigator.gpu.requestAdapter();
      if (adapter && adapter.features && adapter.features.has("shader-f16")) {
        return { device: "webgpu", dtype: "fp16" };
      }
    } catch (err) {
      console.warn("WebGPU is unavailable; background removal will use the CPU.", err);
    }
  }
  return { device: "wasm", dtype: "fp32" };
}

async function openSession(backend) {
  const { AutoModel, AutoProcessor } = await loadTransformers();
  const processor = await AutoProcessor.from_pretrained(MODEL_ID);
  const model = await AutoModel.from_pretrained(MODEL_ID, backend);
  return { model, processor };
}

function getSession(onStage) {
  if (!sessionPromise) {
    sessionPromise = (async () => {
      onStage?.("loading-model");
      const backend = await preferredBackend();
      try {
        return await openSession(backend);
      } catch (err) {
        if (backend.device === "webgpu") {
          console.warn("WebGPU background removal failed; retrying on the CPU.", err);
          return openSession({ device: "wasm", dtype: "fp32" });
        }
        throw err;
      }
    })().catch((err) => {
      sessionPromise = null;
      throw err;
    });
  }
  return sessionPromise;
}

/**
 * @param {Blob} blob
 * @param {{ onStage?: (stage: "loading-model" | "removing") => void }} [options]
 * @returns {Promise<Blob>}
 */
export function removeBackgroundFromBlob(blob, options = {}) {
  return enqueueRemoval(() => removeBackgroundNow(blob, options));
}

async function removeBackgroundNow(blob, options) {
  if (!birefnetDisabled) {
    try {
      return await removeWithBiRefNet(blob, options);
    } catch (err) {
      console.warn("BiRefNet background removal failed; using the previous remover.", err);
      birefnetDisabled = true;
      sessionPromise = null;
    }
  }
  options.onStage?.("removing");
  return removeWithImgly(blob);
}

async function removeWithBiRefNet(blob, options) {
  const session = await getSession(options.onStage);
  options.onStage?.("removing");
  const image = await blobToImage(blob);
  const width = image.naturalWidth || image.width;
  const height = image.naturalHeight || image.height;
  const source = drawImage(image, width, height);
  const { RawImage } = await loadTransformers();
  const url = URL.createObjectURL(blob);
  try {
    const raw = await RawImage.fromURL(url);
    const { pixel_values: pixelValues } = await session.processor(raw);
    const outputs = await session.model({ input_image: pixelValues });
    const mask = maskFromOutputs(outputs);
    return cutoutFromMask(source, mask, width, height);
  } finally {
    URL.revokeObjectURL(url);
  }
}

function maskFromOutputs(outputs) {
  const tensor = outputs.logits ?? outputs.output_image ?? outputs.output;
  if (!tensor || typeof tensor.sigmoid !== "function") {
    const keys = outputs ? Object.keys(outputs).join(", ") : "";
    throw new Error(`Background model returned no mask (${keys}).`);
  }
  let matte = tensor.sigmoid();
  while (matte.dims && matte.dims.length > 2) {
    matte = matte.squeeze(0);
  }
  const uint8 = matte.mul(255).to("uint8");
  const height = uint8.dims[uint8.dims.length - 2];
  const width = uint8.dims[uint8.dims.length - 1];
  if (!width || !height || uint8.data.length < width * height) {
    throw new Error("Background model returned an empty mask.");
  }
  return { data: uint8.data, width, height };
}

function cutoutFromMask(sourceCanvas, mask, width, height) {
  const ctx = sourceCanvas.getContext("2d", { willReadFrequently: true });
  const image = ctx.getImageData(0, 0, width, height);
  const pixels = image.data;
  const maskCanvas = document.createElement("canvas");
  maskCanvas.width = mask.width;
  maskCanvas.height = mask.height;
  const maskCtx = maskCanvas.getContext("2d");
  const maskImage = maskCtx.createImageData(mask.width, mask.height);
  for (let i = 0; i < mask.width * mask.height; i += 1) {
    const alpha = cleanAlpha(mask.data[i]);
    const p = i * 4;
    maskImage.data[p + 3] = alpha;
  }
  maskCtx.putImageData(maskImage, 0, 0);
  const scaled = document.createElement("canvas");
  scaled.width = width;
  scaled.height = height;
  const scaledCtx = scaled.getContext("2d");
  scaledCtx.imageSmoothingEnabled = true;
  scaledCtx.drawImage(maskCanvas, 0, 0, width, height);
  const scaledMask = scaledCtx.getImageData(0, 0, width, height).data;
  for (let i = 0; i < width * height; i += 1) {
    pixels[i * 4 + 3] = Math.min(pixels[i * 4 + 3], cleanAlpha(scaledMask[i * 4 + 3]));
  }
  ctx.putImageData(image, 0, 0);
  return canvasToPng(sourceCanvas);
}

function cleanAlpha(value) {
  if (value <= HAZE_ALPHA) return 0;
  if (value >= 243) return 255;
  return value;
}

async function removeWithImgly(blob) {
  const removeBackground = await getImglyRemover();
  const config = await imglyConfig();
  const result = await removeBackground(blob, config);
  if (result instanceof Blob) return result;
  if (result && result.blob instanceof Blob) return result.blob;
  if (result instanceof ArrayBuffer) return new Blob([result], { type: "image/png" });
  if (result && typeof result.arrayBuffer === "function") {
    const buffer = await result.arrayBuffer();
    return new Blob([buffer], { type: result.type || "image/png" });
  }
  throw new Error("Unexpected output from background remover.");
}

function getImglyRemover() {
  if (!imglyPromise) {
    imglyPromise = import(IMGLY_URL).then((mod) => {
      const candidates = [mod.default, mod.removeBackground, mod.removeBg, mod.default && mod.default.removeBackground];
      const fn = candidates.find((candidate) => typeof candidate === "function");
      if (!fn) throw new Error(`Could not find remover function. Module keys: ${Object.keys(mod).join(", ")}`);
      return fn;
    }).catch((err) => {
      imglyPromise = null;
      throw err;
    });
  }
  return imglyPromise;
}

async function imglyConfig() {
  const backend = await preferredBackend();
  return { device: backend.device === "webgpu" ? "gpu" : "cpu" };
}

/**
 * Draw a solid outline around each foreground object. Separate objects each
 * get their own ring; the ring sits outside the pixels that were kept.
 * @param {Blob} blob
 * @param {string} color
 * @param {number} thickness
 */
export async function outlineCutoutBlob(blob, color, thickness) {
  const image = await blobToImage(blob);
  const width = image.naturalWidth || image.width;
  const height = image.naturalHeight || image.height;
  const radius = Math.max(1, Math.round(Number(thickness) || 1));
  const pad = radius + 2;
  const source = drawImage(image, width, height);
  const srcCtx = source.getContext("2d", { willReadFrequently: true });
  const src = srcCtx.getImageData(0, 0, width, height);
  const outW = width + pad * 2;
  const outH = height + pad * 2;
  const foreground = new Uint8Array(outW * outH);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      if (src.data[(y * width + x) * 4 + 3] > FOREGROUND_ALPHA) {
        foreground[(y + pad) * outW + (x + pad)] = 1;
      }
    }
  }
  const distance = chamferDistance(foreground, outW, outH);
  const [red, green, blue] = hexToRgb(color);
  const out = document.createElement("canvas");
  out.width = outW;
  out.height = outH;
  const outCtx = out.getContext("2d");
  const painted = outCtx.createImageData(outW, outH);
  const limit = radius * 3;
  for (let i = 0; i < outW * outH; i += 1) {
    const d = distance[i];
    if (d > 0 && d <= limit) {
      const p = i * 4;
      painted.data[p] = red;
      painted.data[p + 1] = green;
      painted.data[p + 2] = blue;
      painted.data[p + 3] = 255;
    }
  }
  outCtx.putImageData(painted, 0, 0);
  outCtx.drawImage(source, pad, pad);
  return canvasToPng(out);
}

/** Integer chamfer distance (costs 3 and 4) to the nearest foreground pixel. */
function chamferDistance(foreground, width, height) {
  const inf = (width + height) * 3;
  const dist = new Int32Array(width * height);
  for (let i = 0; i < dist.length; i += 1) dist[i] = foreground[i] ? 0 : inf;

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const i = y * width + x;
      let d = dist[i];
      if (x > 0) d = Math.min(d, dist[i - 1] + 3);
      if (y > 0) d = Math.min(d, dist[i - width] + 3);
      if (x > 0 && y > 0) d = Math.min(d, dist[i - width - 1] + 4);
      if (x + 1 < width && y > 0) d = Math.min(d, dist[i - width + 1] + 4);
      dist[i] = d;
    }
  }
  for (let y = height - 1; y >= 0; y -= 1) {
    for (let x = width - 1; x >= 0; x -= 1) {
      const i = y * width + x;
      let d = dist[i];
      if (x + 1 < width) d = Math.min(d, dist[i + 1] + 3);
      if (y + 1 < height) d = Math.min(d, dist[i + width] + 3);
      if (x + 1 < width && y + 1 < height) d = Math.min(d, dist[i + width + 1] + 4);
      if (x > 0 && y + 1 < height) d = Math.min(d, dist[i + width - 1] + 4);
      dist[i] = d;
    }
  }
  return dist;
}

function hexToRgb(hex) {
  let h = String(hex || "#FFFF00").trim().replace("#", "");
  if (h.length === 3) h = h.split("").map((channel) => channel + channel).join("");
  const value = Number.parseInt(h, 16);
  if (!Number.isFinite(value)) return [255, 255, 0];
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255];
}

function blobToImage(blob) {
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

function drawImage(image, width, height) {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Could not read the image.");
  ctx.drawImage(image, 0, 0, width, height);
  return canvas;
}

function canvasToPng(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("Could not build the cutout."));
        return;
      }
      resolve(blob);
    }, "image/png");
  });
}
