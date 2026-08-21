const JSZIP_URL = "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";
const HTML2CANVAS_URL = "https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/+esm";

const XLINK_NS = "http://www.w3.org/1999/xlink";
const DEFAULT_PAGE_WIDTH = 1100;
const MAX_RASTER_HEIGHT = 4000;
const RASTER_EXT = /\.(jpe?g|png|gif|webp|bmp)$/i;

let jszipPromise = null;
let html2canvasPromise = null;

export class EpubError extends Error {
  /**
   * @param {'drm'|'invalid'|'noPages'} code
   * @param {string} message
   */
  constructor(code, message) {
    super(message);
    this.name = "EpubError";
    this.code = code;
  }
}

export function isEpubFile(file) {
  if (!file) return false;
  const lower = (file.name || "").toLowerCase();
  const type = (file.type || "").toLowerCase();
  if (lower.endsWith(".epub")) return true;
  return type === "application/epub+zip";
}

function pickExport(mod) {
  return mod && (mod.default || mod);
}

async function getJSZip() {
  if (jszipPromise) return jszipPromise;
  jszipPromise = import(JSZIP_URL)
    .then((mod) => {
      const JSZip = pickExport(mod);
      if (typeof JSZip !== "function") throw new Error("JSZip did not load correctly.");
      return JSZip;
    })
    .catch((err) => {
      jszipPromise = null;
      throw err;
    });
  return jszipPromise;
}

async function getHtml2Canvas() {
  if (html2canvasPromise) return html2canvasPromise;
  html2canvasPromise = import(HTML2CANVAS_URL)
    .then((mod) => {
      const fn = pickExport(mod);
      if (typeof fn !== "function") throw new Error("html2canvas did not load correctly.");
      return fn;
    })
    .catch((err) => {
      html2canvasPromise = null;
      throw err;
    });
  return html2canvasPromise;
}

function localNameOf(el) {
  return (el.localName || el.tagName || "").toLowerCase();
}

function xmlElements(xml, localName) {
  const want = localName.toLowerCase();
  return [...xml.getElementsByTagName("*")].filter((el) => localNameOf(el) === want);
}

function parseXml(text, label) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (xmlElements(doc, "parsererror").length) {
    throw new EpubError("invalid", `Could not parse ${label}.`);
  }
  return doc;
}

function dirname(path) {
  const i = path.lastIndexOf("/");
  return i <= 0 ? "" : path.slice(0, i);
}

function normalizeZipPath(path) {
  const parts = [];
  for (const p of String(path || "").replace(/\\/g, "/").split("/")) {
    if (!p || p === ".") continue;
    if (p === "..") parts.pop();
    else parts.push(p);
  }
  return parts.join("/");
}

function decodeHref(href) {
  const raw = String(href || "").trim().split("#")[0].split("?")[0];
  if (!raw) return "";
  try {
    return decodeURIComponent(raw);
  } catch {
    return raw;
  }
}

function resolveZipPath(baseFile, relativeHref) {
  const rel = decodeHref(relativeHref);
  if (!rel) return "";
  if (/^(https?:|data:|blob:|mailto:|javascript:)/i.test(rel)) return "";
  if (rel.startsWith("/")) return normalizeZipPath(rel.slice(1));
  const dir = dirname(baseFile);
  return normalizeZipPath(dir ? `${dir}/${rel}` : rel);
}

function buildZipIndex(zip) {
  /** @type {Map<string, string>} */
  const index = new Map();
  for (const name of Object.keys(zip.files || {})) {
    if (zip.files[name].dir) continue;
    index.set(normalizeZipPath(name).toLowerCase(), name);
  }
  return index;
}

function zipEntry(zip, index, path) {
  const normalized = normalizeZipPath(path);
  if (!normalized) return null;
  if (zip.files[normalized] && !zip.files[normalized].dir) return zip.files[normalized];
  const actual = index.get(normalized.toLowerCase());
  if (actual && zip.files[actual] && !zip.files[actual].dir) return zip.files[actual];
  return null;
}

function guessMediaType(path, declared) {
  const d = (declared || "").toLowerCase().trim();
  if (d) return d;
  const lower = (path || "").toLowerCase();
  if (lower.endsWith(".xhtml") || lower.endsWith(".html") || lower.endsWith(".htm")) return "application/xhtml+xml";
  if (lower.endsWith(".svg")) return "image/svg+xml";
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".gif")) return "image/gif";
  if (lower.endsWith(".webp")) return "image/webp";
  if (lower.endsWith(".css")) return "text/css";
  return "";
}

function isHtmlMediaType(mt) {
  const m = (mt || "").toLowerCase();
  return (
    m === "application/xhtml+xml" ||
    m === "text/html" ||
    m === "application/html+xml" ||
    m === "application/x-dtbook+xml"
  );
}

function isImageMediaType(mt) {
  return (mt || "").toLowerCase().startsWith("image/");
}

function isRasterPath(path) {
  return RASTER_EXT.test(path || "");
}

function parseViewport(content) {
  if (!content) return null;
  const widthMatch = String(content).match(/width\s*=\s*(\d+(?:\.\d+)?)/i);
  const heightMatch = String(content).match(/height\s*=\s*(\d+(?:\.\d+)?)/i);
  const width = widthMatch ? Number(widthMatch[1]) : NaN;
  const height = heightMatch ? Number(heightMatch[1]) : NaN;
  if (!Number.isFinite(width) || width < 32) return null;
  return {
    width: Math.round(width),
    height: Number.isFinite(height) && height >= 32 ? Math.round(height) : null
  };
}

function canvasToPngBlob(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("Could not rasterize EPUB page."))), "image/png");
  });
}

function itemProperties(el) {
  return (el.getAttribute("properties") || "").toLowerCase().split(/\s+/).filter(Boolean);
}

/**
 * @param {import("jszip")} zip
 * @param {Map<string, string>} index
 */
async function readOpf(zip, index) {
  const containerEntry = zipEntry(zip, index, "META-INF/container.xml");
  if (!containerEntry) throw new EpubError("invalid", "Missing META-INF/container.xml.");
  const containerXml = await containerEntry.async("text");
  const containerDoc = parseXml(containerXml, "container.xml");
  const rootfile = xmlElements(containerDoc, "rootfile")[0];
  const opfPath = normalizeZipPath(rootfile?.getAttribute("full-path") || "");
  if (!opfPath) throw new EpubError("invalid", "EPUB is missing a package document.");
  const opfEntry = zipEntry(zip, index, opfPath);
  if (!opfEntry) throw new EpubError("invalid", "Could not find the EPUB package document.");
  const opfXml = await opfEntry.async("text");
  return { opfPath, opfDoc: parseXml(opfXml, "package document") };
}

function readSpineItems(opfDoc, opfPath) {
  /** @type {Map<string, { id: string, href: string, mediaType: string, properties: string[] }>} */
  const manifest = new Map();
  for (const item of xmlElements(opfDoc, "item")) {
    const id = item.getAttribute("id");
    const href = item.getAttribute("href");
    if (!id || !href) continue;
    const path = resolveZipPath(opfPath, href);
    manifest.set(id, {
      id,
      href: path,
      mediaType: guessMediaType(path, item.getAttribute("media-type")),
      properties: itemProperties(item)
    });
  }

  const spine = [];
  for (const itemref of xmlElements(opfDoc, "itemref")) {
    if ((itemref.getAttribute("linear") || "yes").toLowerCase() === "no") continue;
    const idref = itemref.getAttribute("idref");
    const item = idref ? manifest.get(idref) : null;
    if (!item) continue;
    if (item.properties.includes("nav")) continue;
    spine.push(item);
  }
  return spine;
}

function sanitizeHtmlDocument(doc) {
  doc.querySelectorAll("script, iframe, object, embed, form, link[rel='preload'], link[rel='prefetch']").forEach((el) => el.remove());
  for (const el of [...doc.querySelectorAll("*")]) {
    for (const attr of [...el.attributes]) {
      const name = attr.name.toLowerCase();
      if (name.startsWith("on")) el.removeAttribute(attr.name);
      if ((name === "href" || name === "src") && /^\s*javascript:/i.test(attr.value)) {
        el.removeAttribute(attr.name);
      }
    }
  }
}

function extractText(doc) {
  const text = (doc.body?.innerText || doc.documentElement?.textContent || "").replace(/\s+/g, " ").trim();
  return text;
}

function attrHref(el) {
  return (
    el.getAttribute("src") ||
    el.getAttribute("href") ||
    el.getAttributeNS(XLINK_NS, "href") ||
    ""
  );
}

function singleLocalRasterHref(doc, htmlPath) {
  const nodes = [...doc.querySelectorAll("img, image")];
  const local = [];
  for (const el of nodes) {
    const href = attrHref(el);
    if (!href || /^(https?:|data:|blob:)/i.test(href)) continue;
    const resolved = resolveZipPath(htmlPath, href);
    if (resolved && isRasterPath(resolved)) local.push(resolved);
  }
  if (local.length !== 1) return null;
  return local[0];
}

function rewriteCssUrls(cssText, cssPath, toBlobUrl) {
  return String(cssText || "").replace(/url\(\s*(['"]?)([^'")]+)\1\s*\)/gi, (match, quote, rawUrl) => {
    const url = String(rawUrl || "").trim();
    if (!url || /^(https?:|data:|blob:|#)/i.test(url)) return match;
    const blobUrl = toBlobUrl(resolveZipPath(cssPath, url));
    if (!blobUrl) return match;
    return `url(${quote}${blobUrl}${quote})`;
  });
}

function collectCssUrlPaths(cssText, basePath) {
  const needed = [];
  String(cssText || "").replace(/url\(\s*(['"]?)([^'")]+)\1\s*\)/gi, (_, __, rawUrl) => {
    const url = String(rawUrl || "").trim();
    if (url && !/^(https?:|data:|blob:|#)/i.test(url)) needed.push(resolveZipPath(basePath, url));
    return _;
  });
  return needed;
}

/**
 * @param {object} opts
 * @param {*} opts.zip
 * @param {Map<string, string>} opts.index
 * @param {Map<string, string>} opts.blobUrlCache
 * @param {string[]} opts.createdUrls
 */
function createBlobUrlResolver({ zip, index, blobUrlCache, createdUrls }) {
  /** @type {Map<string, Promise<string|null>>} */
  const inflight = new Map();

  const resolve = (path, mimeHint = "") => {
    const normalized = normalizeZipPath(path);
    if (!normalized) return Promise.resolve(null);
    if (blobUrlCache.has(normalized)) return Promise.resolve(blobUrlCache.get(normalized));
    if (inflight.has(normalized)) return inflight.get(normalized);

    const work = (async () => {
      const entry = zipEntry(zip, index, normalized);
      if (!entry) return null;
      let blob = await entry.async("blob");
      const type = guessMediaType(normalized, mimeHint || blob.type);
      if (type && blob.type !== type) blob = new Blob([blob], { type });
      if (type === "text/css" || normalized.toLowerCase().endsWith(".css")) {
        const cssText = await blob.text();
        await Promise.all(collectCssUrlPaths(cssText, normalized).map((p) => resolve(p)));
        const css2 = rewriteCssUrls(cssText, normalized, (p) => blobUrlCache.get(p) || "");
        blob = new Blob([css2], { type: "text/css" });
      }
      const url = URL.createObjectURL(blob);
      blobUrlCache.set(normalized, url);
      createdUrls.push(url);
      return url;
    })();

    inflight.set(normalized, work);
    return work;
  };

  return resolve;
}

async function rewriteHtmlResources(doc, htmlPath, resolveBlobUrl) {
  const jobs = [];

  const setUrlAttr = (el, attr, ns, path) => {
    jobs.push(
      resolveBlobUrl(path).then((url) => {
        if (!url) return;
        if (ns) el.setAttributeNS(ns, attr, url);
        else el.setAttribute(attr, url);
      })
    );
  };

  for (const el of doc.querySelectorAll("img, source, video, audio, input[type='image']")) {
    const src = el.getAttribute("src");
    if (src) setUrlAttr(el, "src", null, resolveZipPath(htmlPath, src));
    const poster = el.getAttribute("poster");
    if (poster) setUrlAttr(el, "poster", null, resolveZipPath(htmlPath, poster));
    const srcset = el.getAttribute("srcset");
    if (srcset) {
      jobs.push(
        (async () => {
          const parts = srcset.split(",");
          const rewritten = [];
          for (const part of parts) {
            const trimmed = part.trim();
            if (!trimmed) continue;
            const bits = trimmed.split(/\s+/);
            const url = await resolveBlobUrl(resolveZipPath(htmlPath, bits[0]));
            bits[0] = url || bits[0];
            rewritten.push(bits.join(" "));
          }
          el.setAttribute("srcset", rewritten.join(", "));
        })()
      );
    }
  }

  for (const el of doc.querySelectorAll("image")) {
    const href = el.getAttribute("href") || el.getAttributeNS(XLINK_NS, "href");
    if (href) {
      setUrlAttr(el, "href", null, resolveZipPath(htmlPath, href));
      setUrlAttr(el, "href", XLINK_NS, resolveZipPath(htmlPath, href));
    }
  }

  for (const el of doc.querySelectorAll("link[href]")) {
    const rel = (el.getAttribute("rel") || "").toLowerCase();
    if (!/\bstylesheet\b/.test(rel) && el.getAttribute("as") !== "style") continue;
    const href = el.getAttribute("href");
    if (href) setUrlAttr(el, "href", null, resolveZipPath(htmlPath, href));
  }

  for (const el of doc.querySelectorAll("style")) {
    jobs.push(Promise.all(collectCssUrlPaths(el.textContent || "", htmlPath).map((p) => resolveBlobUrl(p))));
  }

  await Promise.all(jobs);
}

function rewriteInlineCssWithCache(doc, htmlPath, blobUrlCache) {
  const lookup = (p) => blobUrlCache.get(p) || "";
  for (const el of doc.querySelectorAll("style")) {
    el.textContent = rewriteCssUrls(el.textContent || "", htmlPath, lookup);
  }
  for (const el of doc.querySelectorAll("[style]")) {
    el.setAttribute("style", rewriteCssUrls(el.getAttribute("style") || "", htmlPath, lookup));
  }
}

async function waitForDocumentImages(doc, timeoutMs = 12000) {
  const imgs = [...(doc.images || [])];
  await Promise.all(
    imgs.map(
      (img) =>
        new Promise((resolve) => {
          if (img.complete) {
            resolve();
            return;
          }
          const t = setTimeout(resolve, timeoutMs);
          img.addEventListener(
            "load",
            () => {
              clearTimeout(t);
              resolve();
            },
            { once: true }
          );
          img.addEventListener(
            "error",
            () => {
              clearTimeout(t);
              resolve();
            },
            { once: true }
          );
        })
    )
  );
}

async function rasterizeHtml(html, { width, height, scale }) {
  const html2canvas = await getHtml2Canvas();
  const iframe = document.createElement("iframe");
  iframe.setAttribute("sandbox", "allow-same-origin");
  iframe.setAttribute("scrolling", "no");
  iframe.setAttribute("aria-hidden", "true");
  Object.assign(iframe.style, {
    position: "fixed",
    left: "0",
    top: "0",
    width: `${width}px`,
    height: `${height}px`,
    opacity: "0",
    pointerEvents: "none",
    border: "0",
    zIndex: "-1"
  });
  document.body.appendChild(iframe);

  try {
    const doc = iframe.contentDocument;
    if (!doc) throw new Error("Could not open an EPUB preview frame.");
    doc.open();
    doc.write(html);
    doc.close();

    await new Promise((resolve) => {
      if (doc.readyState === "complete") resolve();
      else iframe.addEventListener("load", () => resolve(), { once: true });
      setTimeout(resolve, 1500);
    });
    await waitForDocumentImages(doc);

    if (doc.body) {
      doc.body.style.margin = "0";
      doc.body.style.width = `${width}px`;
    }
    const contentH = Math.max(
      doc.documentElement?.scrollHeight || 0,
      doc.body?.scrollHeight || 0,
      height
    );
    const finalH = Math.min(Math.max(contentH, height), MAX_RASTER_HEIGHT);
    iframe.style.height = `${finalH}px`;
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));

    const target = doc.body || doc.documentElement;
    const canvas = await html2canvas(target, {
      scale,
      width,
      height: finalH,
      windowWidth: width,
      windowHeight: finalH,
      useCORS: true,
      allowTaint: true,
      backgroundColor: "#ffffff",
      logging: false,
      imageTimeout: 8000
    });
    return canvasToPngBlob(canvas);
  } finally {
    iframe.remove();
  }
}

async function blobFromZip(zip, index, path, mediaType) {
  const entry = zipEntry(zip, index, path);
  if (!entry) return null;
  let blob = await entry.async("blob");
  const type = guessMediaType(path, mediaType || blob.type);
  if (type && blob.type !== type) blob = new Blob([blob], { type });
  return blob;
}

async function rasterizeSvgBlob(blob, scale) {
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.decoding = "async";
    const loaded = new Promise((resolve, reject) => {
      img.onload = () => resolve();
      img.onerror = () => reject(new Error("Could not render SVG page."));
    });
    img.src = url;
    await loaded;
    const width = Math.max(1, img.naturalWidth || DEFAULT_PAGE_WIDTH);
    const height = Math.max(1, img.naturalHeight || Math.round(width * 0.75));
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(width * scale);
    canvas.height = Math.round(height * scale);
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    return canvasToPngBlob(canvas);
  } finally {
    URL.revokeObjectURL(url);
  }
}

/**
 * @param {File|Blob} file
 * @param {{
 *   maxPages?: number,
 *   renderScale?: number,
 *   ensureCompatibleImage?: (file: File) => Promise<File>,
 *   JSZip?: typeof import("jszip"),
 * }} [options]
 * @returns {Promise<{ pages: { blob: Blob, name: string, ocrText: string, textSource: 'epub'|null }[], total: number }>}
 */
export async function renderEpubToPages(file, options = {}) {
  const maxPages = options.maxPages || 40;
  const renderScale = options.renderScale || 2;
  const ensureCompatibleImage = options.ensureCompatibleImage;
  const JSZip = options.JSZip || (await getJSZip());
  const data = typeof file.arrayBuffer === "function" ? await file.arrayBuffer() : file;
  const zip = await JSZip.loadAsync(data);
  const index = buildZipIndex(zip);

  if (zipEntry(zip, index, "META-INF/encryption.xml")) {
    throw new EpubError("drm", "This EPUB is encrypted (DRM) and cannot be opened in the browser.");
  }

  const { opfPath, opfDoc } = await readOpf(zip, index);
  const spine = readSpineItems(opfDoc, opfPath);
  if (!spine.length) throw new EpubError("noPages", "No readable pages were found in this EPUB.");

  const createdUrls = [];
  const blobUrlCache = new Map();
  const resolveBlobUrl = createBlobUrlResolver({ zip, index, blobUrlCache, createdUrls });
  const baseName = (file.name || "book").replace(/\.[^/.]+$/, "");
  const pages = [];

  try {
    const limit = Math.min(spine.length, maxPages);
    for (let i = 0; i < limit; i += 1) {
      const item = spine[i];
      const pageName = `${baseName}-p${i + 1}.png`;

      if (isImageMediaType(item.mediaType) && item.mediaType !== "image/svg+xml") {
        let blob = await blobFromZip(zip, index, item.href, item.mediaType);
        if (!blob) continue;
        if (ensureCompatibleImage) {
          const ready = await ensureCompatibleImage(new File([blob], item.href.split("/").pop() || pageName, { type: blob.type }));
          blob = ready;
        }
        pages.push({ blob, name: pageName, ocrText: "", textSource: null });
        continue;
      }

      if (item.mediaType === "image/svg+xml") {
        const svgBlob = await blobFromZip(zip, index, item.href, item.mediaType);
        if (!svgBlob) continue;
        const blob = await rasterizeSvgBlob(svgBlob, renderScale);
        pages.push({ blob, name: pageName, ocrText: "", textSource: null });
        continue;
      }

      if (!isHtmlMediaType(item.mediaType)) continue;

      const htmlEntry = zipEntry(zip, index, item.href);
      if (!htmlEntry) continue;
      const htmlText = await htmlEntry.async("text");
      const doc = new DOMParser().parseFromString(htmlText, "text/html");
      sanitizeHtmlDocument(doc);
      const ocrText = extractText(doc);
      const textSource = ocrText ? "epub" : null;

      const rasterHref = singleLocalRasterHref(doc, item.href);
      const rasterFallbacks = [...doc.querySelectorAll("img, image")]
        .map((el) => resolveZipPath(item.href, attrHref(el)))
        .filter((p) => p && isRasterPath(p));
      if (rasterHref) {
        let blob = await blobFromZip(zip, index, rasterHref);
        if (blob) {
          if (ensureCompatibleImage) {
            const ready = await ensureCompatibleImage(
              new File([blob], rasterHref.split("/").pop() || pageName, { type: blob.type || "image/jpeg" })
            );
            blob = ready;
          }
          pages.push({ blob, name: pageName, ocrText, textSource });
          continue;
        }
      }

      const viewportEl = doc.querySelector('meta[name="viewport"]');
      const viewport = parseViewport(viewportEl?.getAttribute("content"));
      const width = viewport?.width || DEFAULT_PAGE_WIDTH;
      const height = viewport?.height || Math.round(width * 0.75);

      await rewriteHtmlResources(doc, item.href, resolveBlobUrl);
      rewriteInlineCssWithCache(doc, item.href, blobUrlCache);

      const serialized = `<!DOCTYPE html>${doc.documentElement.outerHTML}`;
      let blob = null;
      try {
        blob = await rasterizeHtml(serialized, { width, height, scale: renderScale });
      } catch (err) {
        console.warn("EPUB HTML rasterize failed; trying image fallback.", err);
        const fallbackPath = rasterFallbacks[0];
        if (fallbackPath) blob = await blobFromZip(zip, index, fallbackPath);
        if (!blob) throw err;
        if (ensureCompatibleImage) {
          const ready = await ensureCompatibleImage(
            new File([blob], fallbackPath.split("/").pop() || pageName, { type: blob.type || "image/jpeg" })
          );
          blob = ready;
        }
      }
      pages.push({ blob, name: pageName, ocrText, textSource });
    }
  } finally {
    for (const url of createdUrls) URL.revokeObjectURL(url);
  }

  if (!pages.length) throw new EpubError("noPages", "No readable pages were found in this EPUB.");
  return { pages, total: spine.length };
}
