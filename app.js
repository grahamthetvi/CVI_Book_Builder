const spreadsContainer = document.getElementById("spreadsContainer");
const spreadTemplate = document.getElementById("spreadTemplate");
const addSpreadButton = document.getElementById("addSpreadButton");
const parseAiButton = document.getElementById("parseAiButton");
const exportPptxButton = document.getElementById("exportPptxButton");
const refreshPreviewButton = document.getElementById("refreshPreviewButton");
const printPdfButton = document.getElementById("printPdfButton");
const statusMessage = document.getElementById("statusMessage");
const previewContainer = document.getElementById("previewContainer");

const bookTitleInput = document.getElementById("bookTitle");
const oddTextPositionInput = document.getElementById("oddTextPosition");
const oddTextColorInput = document.getElementById("oddTextColor");
const oddTextSizeInput = document.getElementById("oddTextSize");
const oddBorderColorInput = document.getElementById("oddBorderColor");
const oddBorderSizeInput = document.getElementById("oddBorderSize");
const oddBgColorInput = document.getElementById("oddBgColor");
const storyTextColorInput = document.getElementById("storyTextColor");
const visualComplexityInput = document.getElementById("visualComplexity");
const includeTeachingActivitiesInput = document.getElementById("includeTeachingActivities");
const eccAreaInput = document.getElementById("eccArea");
const activityPromptInput = document.getElementById("activityPrompt");
const studentCviTipsInput = document.getElementById("studentCviTips");
const copyFvaLmaButton = document.getElementById("copyFvaLmaButton");
const copyAiPromptButton = document.getElementById("copyAiPromptButton");
const presetMaxContrastButton = document.getElementById("presetMaxContrast");
const presetHighContrastButton = document.getElementById("presetHighContrast");
const presetStandardPrintButton = document.getElementById("presetStandardPrint");
const aiInput = document.getElementById("aiInput");
const localImageInput = document.getElementById("local-image-input");
const wikimediaQueryInput = document.getElementById("wikimedia-query");
const wikimediaSearchButton = document.getElementById("wikimedia-search-btn");
const wikimediaResults = document.getElementById("wikimedia-results");
const outlineEnabledInput = document.getElementById("outline-enabled");
const outlineColorInput = document.getElementById("outline-color");
const outlineThicknessInput = document.getElementById("outline-thickness");
const outlineThicknessValue = document.getElementById("outline-thickness-value");
const processImageButton = document.getElementById("process-image-btn");
const downloadProcessedButton = document.getElementById("download-processed-btn");
const imageToolStatus = document.getElementById("image-tool-status");
const sourcePreviewImage = document.getElementById("source-preview");
const sourcePlaceholder = document.getElementById("source-placeholder");
const processedPreviewImage = document.getElementById("processed-preview");
const processedPlaceholder = document.getElementById("processed-placeholder");
const welcomeModal = document.getElementById("welcomeModal");
const closeWelcomeButton = document.getElementById("closeWelcomeButton");
const printablePreviewSection = previewContainer.closest(".panel");

let previewObjectUrls = [];
let livePreviewTimer = null;
let currentSourceUrl = "";
let currentSourceObjectUrl = null;
let currentSourceName = "";
let processedResultObjectUrl = null;
let removeBackgroundFnPromise = null;

function scheduleLivePreview(immediate = false) {
  if (immediate) {
    if (livePreviewTimer) clearTimeout(livePreviewTimer);
    livePreviewTimer = null;
    renderPreview();
    return;
  }
  if (livePreviewTimer) clearTimeout(livePreviewTimer);
  livePreviewTimer = setTimeout(() => {
    livePreviewTimer = null;
    renderPreview();
  }, 200);
}

/* Color name lookup for colorblind-friendly display (hex -> name) */
const COLOR_NAMES = {
  "#FFFFFF": "White", "#FFF": "White", "#000000": "Black", "#000": "Black",
  "#FFFF00": "Yellow", "#FF0": "Yellow", "#FF0000": "Red", "#F00": "Red",
  "#00FF00": "Green", "#0F0": "Green", "#0000FF": "Blue", "#00F": "Blue",
  "#FF00FF": "Magenta", "#F0F": "Magenta", "#00FFFF": "Cyan", "#0FF": "Cyan",
  "#111111": "Very dark gray", "#222222": "Dark gray", "#333333": "Gray",
  "#666666": "Medium gray", "#999999": "Light gray", "#CCCCCC": "Light gray",
  "#C6D0DD": "Light gray-blue", "#FF9E9E": "Light red", "#B9C2CE": "Gray-blue",
  "#2D5BD1": "Blue", "#4C5D75": "Slate", "#384353": "Dark slate"
};

function hexToColorName(hex) {
  if (!hex) return "Unknown";
  let h = hex.toUpperCase().trim();
  if (!h.startsWith("#")) h = "#" + h;
  if (h.length === 4) h = "#" + h.slice(1).split("").map((c) => c + c).join("");
  return COLOR_NAMES[h] || `Custom (${h})`;
}

function syncColorDisplay(colorInputId) {
  const colorInput = document.getElementById(colorInputId);
  const presetSelect = document.getElementById(colorInputId + "Preset");
  const nameSpan = document.getElementById(colorInputId + "-name");
  if (!colorInput || !nameSpan) return;
  const hex = colorInput.value.toUpperCase();
  const withHash = hex.startsWith("#") ? hex : "#" + hex;
  let label = hexToColorName(withHash);

  if (presetSelect) {
    const exactOption = Array.from(presetSelect.options).find((opt) => opt.value.toUpperCase() === withHash);
    if (exactOption) {
      presetSelect.value = exactOption.value;
      if (exactOption.value !== "custom") {
        label = exactOption.textContent;
      }
    } else {
      presetSelect.value = "custom";
    }
  }

  nameSpan.textContent = label;
  nameSpan.setAttribute("aria-label", `Color name: ${label}`);
}

function setSidebarColorName(colorInputId, nameSpanId) {
  const colorInput = document.getElementById(colorInputId);
  const presetSelect = document.getElementById(colorInputId + "Preset");
  const nameSpan = document.getElementById(nameSpanId);
  if (!colorInput || !nameSpan) return;
  const hex = colorInput.value.toUpperCase();
  const withHash = hex.startsWith("#") ? hex : "#" + hex;
  let label = hexToColorName(withHash);
  if (presetSelect) {
    const exactOption = Array.from(presetSelect.options).find((opt) => opt.value.toUpperCase() === withHash);
    if (exactOption && exactOption.value !== "custom") label = exactOption.textContent;
    else if (exactOption) presetSelect.value = exactOption.value;
  }
  nameSpan.textContent = label;
}

const MAIN_TO_SIDEBAR_IDS = [
  ["oddTextColor", "oddTextColorSidebar", "oddTextColorNameSidebar", "oddTextColorPresetSidebar"],
  ["oddBorderColor", "oddBorderColorSidebar", "oddBorderColorNameSidebar", "oddBorderColorPresetSidebar"],
  ["oddBgColor", "oddBgColorSidebar", "oddBgColorNameSidebar", "oddBgColorPresetSidebar"],
  ["storyTextColor", "storyTextColorSidebar", "storyTextColorNameSidebar", "storyTextColorPresetSidebar"]
];

const MAIN_TO_SIDEBAR_NUMBER_IDS = [
  ["oddTextSize", "oddTextSizeSidebar"],
  ["oddBorderSize", "oddBorderSizeSidebar"]
];

function syncMainToSidebar() {
  MAIN_TO_SIDEBAR_IDS.forEach(([mainId, sidebarId, nameSpanId, presetSidebarId]) => {
    const main = document.getElementById(mainId);
    const sidebar = document.getElementById(sidebarId);
    const presetMain = document.getElementById(mainId + "Preset");
    const presetSidebar = document.getElementById(presetSidebarId);
    if (main && sidebar) {
      sidebar.value = main.value;
      if (presetMain && presetSidebar) presetSidebar.value = presetMain.value;
    }
    setSidebarColorName(mainId, nameSpanId);
  });
  MAIN_TO_SIDEBAR_NUMBER_IDS.forEach(([mainId, sidebarId]) => {
    const main = document.getElementById(mainId);
    const sidebar = document.getElementById(sidebarId);
    if (main && sidebar) sidebar.value = main.value;
  });
}

function syncSidebarToMain(sidebarColorId, mainColorId) {
  const sidebar = document.getElementById(sidebarColorId);
  const main = document.getElementById(mainColorId);
  if (!sidebar || !main) return;
  main.value = sidebar.value;
  const presetSidebar = document.getElementById(sidebarColorId.replace("Sidebar", "PresetSidebar"));
  const presetMain = document.getElementById(mainColorId + "Preset");
  if (presetSidebar && presetMain) presetMain.value = presetSidebar.value;
  syncColorDisplay(mainColorId);
  scheduleLivePreview(true);
}

function initColorPickers() {
  const colorIds = ["oddTextColor", "oddBorderColor", "oddBgColor", "storyTextColor"];
  colorIds.forEach((id) => {
    const colorInput = document.getElementById(id);
    const presetSelect = document.getElementById(id + "Preset");
    if (!colorInput) return;

    colorInput.addEventListener("input", () => {
      syncColorDisplay(id);
      scheduleLivePreview(true);
      syncMainToSidebar();
    });

    if (presetSelect) {
      presetSelect.addEventListener("change", () => {
        if (presetSelect.value !== "custom") {
          colorInput.value = presetSelect.value;
        }
        syncColorDisplay(id);
        if (presetSelect.value === "custom") {
          colorInput.focus();
        }
        scheduleLivePreview(true);
        syncMainToSidebar();
      });
    }

    syncColorDisplay(id);
  });
  syncMainToSidebar();
  initAccessibilitySidebar();
}

function initAccessibilitySidebar() {
  MAIN_TO_SIDEBAR_IDS.forEach(([mainId, sidebarId, nameSpanId, presetSidebarId]) => {
    const sidebarColor = document.getElementById(sidebarId);
    const sidebarPreset = document.getElementById(presetSidebarId);
    if (!sidebarColor) return;
    sidebarColor.addEventListener("input", () => syncSidebarToMain(sidebarId, mainId));
    sidebarColor.addEventListener("change", () => syncSidebarToMain(sidebarId, mainId));
    if (sidebarPreset) {
      sidebarPreset.addEventListener("change", () => {
        if (sidebarPreset.value !== "custom") sidebarColor.value = sidebarPreset.value;
        syncSidebarToMain(sidebarId, mainId);
      });
    }
  });
  MAIN_TO_SIDEBAR_NUMBER_IDS.forEach(([mainId, sidebarId]) => {
    const main = document.getElementById(mainId);
    const sidebar = document.getElementById(sidebarId);
    if (!sidebar || !main) return;
    sidebar.addEventListener("input", () => {
      main.value = sidebar.value;
      scheduleLivePreview(true);
    });
    sidebar.addEventListener("change", () => {
      main.value = sidebar.value;
      scheduleLivePreview(true);
    });
  });
  const oddTextSizeInput = document.getElementById("oddTextSize");
  const oddBorderSizeInput = document.getElementById("oddBorderSize");
  if (oddTextSizeInput) oddTextSizeInput.addEventListener("input", syncMainToSidebar);
  if (oddTextSizeInput) oddTextSizeInput.addEventListener("change", syncMainToSidebar);
  if (oddBorderSizeInput) oddBorderSizeInput.addEventListener("input", syncMainToSidebar);
  if (oddBorderSizeInput) oddBorderSizeInput.addEventListener("change", syncMainToSidebar);
}

function setStatus(text, isError = false) {
  statusMessage.textContent = text;
  statusMessage.style.color = isError ? "#ff9e9e" : "#c6d0dd";
}

function showWelcomeModal() {
  if (!welcomeModal) return;
  welcomeModal.hidden = false;
  welcomeModal.style.display = "flex";
  welcomeModal.setAttribute("aria-hidden", "false");
  if (closeWelcomeButton) closeWelcomeButton.focus();
}

function hideWelcomeModal() {
  if (!welcomeModal) return;
  welcomeModal.hidden = true;
  welcomeModal.style.display = "none";
  welcomeModal.setAttribute("aria-hidden", "true");
}

function safePptColor(hex) {
  if (!hex) return "000000";
  return hex.replace("#", "").toUpperCase();
}

function isHeicLikeFile(file) {
  if (!file || !file.name) return false;
  const lower = file.name.toLowerCase();
  const t = (file.type || "").toLowerCase();
  return (
    lower.endsWith(".heic") ||
    lower.endsWith(".heif") ||
    t === "image/heic" ||
    t === "image/heif" ||
    t === "image/heif-sequence"
  );
}

let heic2anyLoadPromise = null;
function loadHeic2any() {
  if (!heic2anyLoadPromise) {
    heic2anyLoadPromise = import("https://cdn.jsdelivr.net/npm/heic2any@0.0.4/+esm").then((mod) => {
      const fn = mod.default;
      if (typeof fn !== "function") {
        throw new Error("HEIC converter did not load correctly.");
      }
      return fn;
    });
  }
  return heic2anyLoadPromise;
}

async function convertHeicLikeToJpegFile(file) {
  const heic2any = await loadHeic2any();
  const result = await heic2any({
    blob: file,
    toType: "image/jpeg",
    quality: 0.92
  });
  const blob = Array.isArray(result) ? result[0] : result;
  if (!blob || !(blob instanceof Blob)) {
    throw new Error("HEIC conversion produced no image data.");
  }
  const base = file.name.replace(/\.[^/.]+$/i, "") || "image";
  return new File([blob], `${base}.jpg`, { type: "image/jpeg", lastModified: Date.now() });
}

async function ensureBrowserCompatibleImageFile(file) {
  if (!isHeicLikeFile(file)) return file;
  return convertHeicLikeToJpegFile(file);
}

function setImageToolStatus(text, isError = false) {
  if (!imageToolStatus) return;
  imageToolStatus.textContent = text;
  imageToolStatus.style.color = isError ? "#ff9e9e" : "#b9c2ce";
}

function revokeCurrentSourceObjectUrl() {
  if (currentSourceObjectUrl) {
    URL.revokeObjectURL(currentSourceObjectUrl);
    currentSourceObjectUrl = null;
  }
}

function revokeProcessedResultUrl() {
  if (processedResultObjectUrl) {
    URL.revokeObjectURL(processedResultObjectUrl);
    processedResultObjectUrl = null;
  }
}

function resetProcessedPreview() {
  revokeProcessedResultUrl();
  if (processedPreviewImage) {
    processedPreviewImage.removeAttribute("src");
    processedPreviewImage.hidden = true;
  }
  if (processedPlaceholder) {
    processedPlaceholder.hidden = false;
  }
  if (downloadProcessedButton) {
    downloadProcessedButton.setAttribute("href", "#");
    downloadProcessedButton.setAttribute("aria-disabled", "true");
    downloadProcessedButton.classList.add("disabled");
  }
}

function setSourcePreview(url, options = {}) {
  const { isObjectUrl = false, statusText = "", sourceName = "" } = options;
  if (!url) return;

  revokeCurrentSourceObjectUrl();
  currentSourceUrl = url;
  currentSourceName = sourceName;
  if (isObjectUrl) {
    currentSourceObjectUrl = url;
  }

  if (sourcePreviewImage) {
    sourcePreviewImage.src = url;
    sourcePreviewImage.hidden = false;
  }
  if (sourcePlaceholder) {
    sourcePlaceholder.hidden = true;
  }

  resetProcessedPreview();
  if (statusText) setImageToolStatus(statusText);
}

function normalizeReturnedBlob(result) {
  if (result instanceof Blob) {
    return Promise.resolve(result);
  }
  if (result && result.blob instanceof Blob) {
    return Promise.resolve(result.blob);
  }
  if (result instanceof ArrayBuffer) {
    return Promise.resolve(new Blob([result], { type: "image/png" }));
  }
  if (result && typeof result.arrayBuffer === "function") {
    return result.arrayBuffer().then((ab) => new Blob([ab], { type: result.type || "image/png" }));
  }
  return Promise.reject(new Error("Unexpected output from background remover."));
}

function loadImageFromBlob(blob) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve(img);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Failed to decode image for outlining."));
    };
    img.src = url;
  });
}

async function applyOutlineToBlob(blob, color, thickness) {
  const img = await loadImageFromBlob(blob);
  const width = img.naturalWidth || img.width;
  const height = img.naturalHeight || img.height;
  const radius = Math.max(1, Number(thickness) || 1);
  const pad = radius + 2;

  const tintCanvas = document.createElement("canvas");
  tintCanvas.width = width;
  tintCanvas.height = height;
  const tintCtx = tintCanvas.getContext("2d");
  tintCtx.drawImage(img, 0, 0);
  tintCtx.globalCompositeOperation = "source-in";
  tintCtx.fillStyle = color || "#FFFF00";
  tintCtx.fillRect(0, 0, width, height);
  tintCtx.globalCompositeOperation = "source-over";

  const outCanvas = document.createElement("canvas");
  outCanvas.width = width + pad * 2;
  outCanvas.height = height + pad * 2;
  const outCtx = outCanvas.getContext("2d");
  const steps = Math.max(16, Math.ceil(2 * Math.PI * radius * 2));

  for (let i = 0; i < steps; i += 1) {
    const angle = (i / steps) * Math.PI * 2;
    const dx = Math.cos(angle) * radius;
    const dy = Math.sin(angle) * radius;
    outCtx.drawImage(tintCanvas, pad + dx, pad + dy);
  }

  outCtx.drawImage(img, pad, pad);

  return new Promise((resolve, reject) => {
    outCanvas.toBlob((finalBlob) => {
      if (!finalBlob) {
        reject(new Error("Failed to create outlined PNG."));
        return;
      }
      resolve(finalBlob);
    }, "image/png");
  });
}

async function getBackgroundRemover() {
  if (removeBackgroundFnPromise) return removeBackgroundFnPromise;

  removeBackgroundFnPromise = import("https://cdn.jsdelivr.net/npm/@imgly/background-removal/+esm")
    .then((mod) => {
      const candidates = [
        mod.default,
        mod.removeBackground,
        mod.removeBg,
        mod.default && mod.default.removeBackground
      ];
      const fn = candidates.find((candidate) => typeof candidate === "function");
      if (!fn) {
        throw new Error(`Could not find remover function. Module keys: ${Object.keys(mod).join(", ")}`);
      }
      return fn;
    })
    .catch((err) => {
      removeBackgroundFnPromise = null;
      throw err;
    });

  return removeBackgroundFnPromise;
}

function renderWikimediaResults(items) {
  if (!wikimediaResults) return;
  wikimediaResults.innerHTML = "";

  if (!items.length) {
    wikimediaResults.innerHTML = `<p class="hint">No Wikimedia image results found. Try a different search.</p>`;
    return;
  }

  items.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "wikimedia-result-card";
    const img = document.createElement("img");
    img.src = item.thumbUrl;
    img.alt = item.title;
    const title = document.createElement("div");
    title.className = "wikimedia-title";
    title.textContent = item.title;
    button.appendChild(img);
    button.appendChild(title);
    button.addEventListener("click", () => {
      setSourcePreview(item.fullUrl, { statusText: `Selected Wikimedia image: ${item.title}`, sourceName: item.title });
    });
    wikimediaResults.appendChild(button);
  });
}

async function searchWikimediaCommons() {
  if (!wikimediaQueryInput || !wikimediaSearchButton) return;
  const query = (wikimediaQueryInput.value || "").trim();
  if (!query) {
    setImageToolStatus("Enter a Wikimedia search term first.", true);
    return;
  }

  wikimediaSearchButton.disabled = true;
  setImageToolStatus("Searching Wikimedia Commons...");
  if (wikimediaResults) {
    wikimediaResults.innerHTML = "";
  }

  try {
    const url = `https://commons.wikimedia.org/w/api.php?action=query&format=json&origin=*&generator=search&gsrnamespace=6&gsrlimit=12&gsrsearch=${encodeURIComponent(query)}&prop=imageinfo&iiprop=url|mime|thumburl&iiurlwidth=320`;
    const res = await fetch(url);
    if (!res.ok) {
      throw new Error(`Wikimedia request failed (${res.status}).`);
    }
    const data = await res.json();
    const pages = Object.values(data.query?.pages || {});
    const items = pages
      .map((page) => {
        const info = page.imageinfo && page.imageinfo[0];
        if (!info || !info.url) return null;
        return {
          title: (page.title || "").replace(/^File:/i, ""),
          fullUrl: info.url,
          thumbUrl: info.thumburl || info.url
        };
      })
      .filter(Boolean);

    renderWikimediaResults(items);
    setImageToolStatus(`Found ${items.length} Wikimedia image(s). Click one to select.`);
  } catch (err) {
    console.error(err);
    setImageToolStatus(`Wikimedia search failed: ${err.message || "Unknown error"}`, true);
  } finally {
    wikimediaSearchButton.disabled = false;
  }
}

async function processCurrentImage() {
  if (!currentSourceUrl) {
    setImageToolStatus("Choose an image source first.", true);
    return;
  }

  if (!processImageButton) return;
  processImageButton.disabled = true;
  setImageToolStatus("Processing image...");

  try {
    const removeBackground = await getBackgroundRemover();
    const sourceResponse = await fetch(currentSourceUrl);
    if (!sourceResponse.ok) {
      throw new Error(`Failed to load source image (${sourceResponse.status}).`);
    }
    const sourceBlob = await sourceResponse.blob();
    let finalBlob = await normalizeReturnedBlob(await removeBackground(sourceBlob));

    if (outlineEnabledInput && outlineEnabledInput.checked) {
      const color = outlineColorInput ? outlineColorInput.value : "#FFFF00";
      const thickness = outlineThicknessInput ? Number(outlineThicknessInput.value) : 6;
      finalBlob = await applyOutlineToBlob(finalBlob, color, thickness);
    }

    revokeProcessedResultUrl();
    processedResultObjectUrl = URL.createObjectURL(finalBlob);

    if (processedPreviewImage) {
      processedPreviewImage.src = processedResultObjectUrl;
      processedPreviewImage.hidden = false;
    }
    if (processedPlaceholder) {
      processedPlaceholder.hidden = true;
    }
    if (downloadProcessedButton) {
      downloadProcessedButton.href = processedResultObjectUrl;
      const baseName = currentSourceName ? currentSourceName.replace(/\.[^/.]+$/, "") : "isolated-object";
      downloadProcessedButton.download = `${baseName}-isolated.png`;
      downloadProcessedButton.setAttribute("aria-disabled", "false");
      downloadProcessedButton.classList.remove("disabled");
    }

    setImageToolStatus("Done. Background removed and preview updated.");
  } catch (err) {
    console.error(err);
    setImageToolStatus(`Image processing failed: ${err.message || "Unknown error"}`, true);
  } finally {
    processImageButton.disabled = false;
  }
}

function initImageIsolator() {
  if (!localImageInput || !processImageButton) return;

  localImageInput.addEventListener("change", async () => {
    const file = localImageInput.files && localImageInput.files[0];
    if (!file) return;
    try {
      if (isHeicLikeFile(file)) {
        setImageToolStatus("Converting HEIC to JPEG…");
      }
      const readyFile = await ensureBrowserCompatibleImageFile(file);
      const objectUrl = URL.createObjectURL(readyFile);
      const note = readyFile !== file ? ` (converted from HEIC)` : "";
      setSourcePreview(objectUrl, {
        isObjectUrl: true,
        statusText: `Selected local image: ${readyFile.name}${note}`,
        sourceName: readyFile.name
      });
    } catch (err) {
      console.error(err);
      setImageToolStatus(`Could not load image: ${err.message || "Unknown error"}`, true);
    }
  });

  if (wikimediaSearchButton) {
    wikimediaSearchButton.addEventListener("click", searchWikimediaCommons);
  }

  if (wikimediaQueryInput) {
    wikimediaQueryInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        searchWikimediaCommons();
      }
    });
  }

  if (outlineThicknessInput && outlineThicknessValue) {
    outlineThicknessValue.textContent = `${outlineThicknessInput.value}px`;
    outlineThicknessInput.addEventListener("input", () => {
      outlineThicknessValue.textContent = `${outlineThicknessInput.value}px`;
    });
  }

  processImageButton.addEventListener("click", processCurrentImage);
}

function addSpread(initial = {}) {
  const fragment = spreadTemplate.content.cloneNode(true);
  const card = fragment.querySelector(".spread-card");
  const storyText = fragment.querySelector(".story-text");
  const salientFeatures = fragment.querySelector(".salient-features");
  const oddText = fragment.querySelector(".odd-text");
  const imagePromptEl = fragment.querySelector(".image-prompt");
  const imageInput = fragment.querySelector(".odd-images");
  const selectedImagesText = fragment.querySelector(".selected-images");
  const imagePreviewList = fragment.querySelector(".image-preview-list");
  const removeButton = fragment.querySelector(".remove-spread");

  card.currentFiles = [];

  storyText.value = initial.storyText || "";
  salientFeatures.value = initial.salientFeatures || "";
  oddText.value = initial.oddText || "";
  if (imagePromptEl) imagePromptEl.value = initial.imagePrompt || "";

  const renderImageList = () => {
    imagePreviewList.innerHTML = "";
    card.currentFiles.forEach((file, index) => {
      const item = document.createElement("div");
      item.className = "image-preview-item";

      const img = document.createElement("img");
      const url = URL.createObjectURL(file);
      img.src = url;
      img.onload = () => URL.revokeObjectURL(url);

      const controls = document.createElement("div");
      controls.className = "image-preview-controls";

      const leftBtn = document.createElement("button");
      leftBtn.type = "button";
      leftBtn.textContent = "◀";
      leftBtn.disabled = index === 0;
      leftBtn.onclick = () => {
        if (index > 0) {
          [card.currentFiles[index - 1], card.currentFiles[index]] = [card.currentFiles[index], card.currentFiles[index - 1]];
          renderImageList();
          scheduleLivePreview(true);
        }
      };

      const rightBtn = document.createElement("button");
      rightBtn.type = "button";
      rightBtn.textContent = "▶";
      rightBtn.disabled = index === card.currentFiles.length - 1;
      rightBtn.onclick = () => {
        if (index < card.currentFiles.length - 1) {
          [card.currentFiles[index + 1], card.currentFiles[index]] = [card.currentFiles[index], card.currentFiles[index + 1]];
          renderImageList();
          scheduleLivePreview(true);
        }
      };

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.textContent = "✖";
      removeBtn.className = "danger";
      removeBtn.onclick = () => {
        card.currentFiles.splice(index, 1);
        renderImageList();
        updateSelectedImagesText();
        scheduleLivePreview(true);
      };

      controls.appendChild(leftBtn);
      controls.appendChild(removeBtn);
      controls.appendChild(rightBtn);

      item.appendChild(img);
      item.appendChild(controls);
      imagePreviewList.appendChild(item);
    });
  };

  const updateSelectedImagesText = () => {
    selectedImagesText.textContent = card.currentFiles.length
      ? `${card.currentFiles.length} image(s) selected.`
      : "No images selected.";
  };

  imageInput.addEventListener("change", async () => {
    const newFiles = Array.from(imageInput.files || []);
    const showHeicWait = newFiles.some(isHeicLikeFile);
    if (showHeicWait) {
      setStatus("Converting HEIC…", false);
    }

    let added = 0;
    let hadError = false;

    for (const file of newFiles) {
      if (card.currentFiles.length >= 4) {
        setStatus("Only up to 4 images are allowed on odd pages.", true);
        hadError = true;
        break;
      }
      try {
        const readyFile = await ensureBrowserCompatibleImageFile(file);
        card.currentFiles.push(readyFile);
        added += 1;
      } catch (err) {
        console.error(err);
        hadError = true;
        setStatus(`Could not add ${file.name}: ${err.message || "Unknown error"}`, true);
      }
    }

    imageInput.value = "";

    if (showHeicWait && !hadError) {
      setStatus("", false);
    }

    if (added > 0) {
      renderImageList();
      updateSelectedImagesText();
      scheduleLivePreview(true);
    }
  });

  updateSelectedImagesText();

  removeButton.addEventListener("click", () => {
    card.remove();
    renumberSpreads();
    scheduleLivePreview(true);
  });

  const copyImagePromptBtn = fragment.querySelector(".copy-image-prompt");
  if (copyImagePromptBtn && imagePromptEl) {
    copyImagePromptBtn.addEventListener("click", () => {
      const text = imagePromptEl.value.trim();
      if (!text) {
        setStatus("This spread has no image prompt to copy.", true);
        return;
      }
      const toCopy = `generate image: ${text}`;
      copyToClipboardFallback(
        toCopy,
        "Image prompt copied. Paste into your AI image generator.",
        "Failed to copy. Check clipboard permissions."
      );
    });
  }

  spreadsContainer.appendChild(fragment);
  renumberSpreads();
  scheduleLivePreview(true);
}

function renumberSpreads() {
  Array.from(spreadsContainer.querySelectorAll(".spread-card")).forEach((card, idx) => {
    const num = card.querySelector(".spread-number");
    num.textContent = String(idx + 1);
  });
}

function parseAiFormattedText(raw) {
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

function collectSpreadsFromForm() {
  const cards = Array.from(spreadsContainer.querySelectorAll(".spread-card"));
  return cards.map((card, index) => {
    const storyText = card.querySelector(".story-text").value.trim();
    const salientFeatures = card.querySelector(".salient-features").value.trim();
    const oddText = card.querySelector(".odd-text").value.trim();
    const imagePromptEl = card.querySelector(".image-prompt");
    const imagePrompt = imagePromptEl ? imagePromptEl.value.trim() : "";
    const imageFiles = card.currentFiles || [];
    return {
      index,
      storyText,
      salientFeatures,
      oddText,
      imagePrompt,
      imageFiles
    };
  });
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error(`Failed to read ${file.name}`));
    reader.readAsDataURL(file);
  });
}

function getTextOutlineCss(borderSize, borderColor) {
  const size = Math.max(0, Number(borderSize) || 0);
  if (!size) return "none";
  const shadows = [];
  for (let x = -size; x <= size; x += 1) {
    for (let y = -size; y <= size; y += 1) {
      if (x === 0 && y === 0) continue;
      shadows.push(`${x}px ${y}px 0 ${borderColor}`);
    }
  }
  return shadows.join(", ");
}

function distributeImagesInRect(count, area, options = {}) {
  const imagePadding = 0.18;
  const paddedArea = {
    x: area.x + imagePadding,
    y: area.y + imagePadding,
    w: Math.max(0, area.w - imagePadding * 2),
    h: Math.max(0, area.h - imagePadding * 2)
  };
  const gap = options.highSupport
    ? Math.min(paddedArea.w, paddedArea.h) * 0.20
    : 0.12;
  const rects = [];

  if (count <= 0) return rects;
  if (count === 1) return [{ ...paddedArea }];

  if (count === 2) {
    const w = (paddedArea.w - gap) / 2;
    rects.push({ x: paddedArea.x, y: paddedArea.y, w, h: paddedArea.h });
    rects.push({ x: paddedArea.x + w + gap, y: paddedArea.y, w, h: paddedArea.h });
    return rects;
  }

  if (count === 3) {
    const topH = paddedArea.h * 0.54;
    const bottomH = paddedArea.h - topH - gap;
    rects.push({ x: paddedArea.x, y: paddedArea.y, w: paddedArea.w, h: topH });
    const bottomW = (paddedArea.w - gap) / 2;
    rects.push({ x: paddedArea.x, y: paddedArea.y + topH + gap, w: bottomW, h: bottomH });
    rects.push({ x: paddedArea.x + bottomW + gap, y: paddedArea.y + topH + gap, w: bottomW, h: bottomH });
    return rects;
  }

  const w = (paddedArea.w - gap) / 2;
  const h = (paddedArea.h - gap) / 2;
  rects.push({ x: paddedArea.x, y: paddedArea.y, w, h });
  rects.push({ x: paddedArea.x + w + gap, y: paddedArea.y, w, h });
  rects.push({ x: paddedArea.x, y: paddedArea.y + h + gap, w, h });
  rects.push({ x: paddedArea.x + w + gap, y: paddedArea.y + h + gap, w, h });
  return rects;
}

function getOddRegions(position) {
  const slide = { w: 11, h: 8.5 };
  const margin = 0.55;
  const gap = 0.22;
  const textRatio = 0.3;
  const textMin = 2.0;

  if (position === "left" || position === "right") {
    const textW = Math.max(textMin, (slide.w - margin * 2) * textRatio);
    if (position === "left") {
      return {
        textRect: { x: margin, y: margin, w: textW, h: slide.h - margin * 2 },
        imageRect: {
          x: margin + textW + gap,
          y: margin,
          w: slide.w - margin * 2 - textW - gap,
          h: slide.h - margin * 2
        }
      };
    }
    return {
      textRect: { x: slide.w - margin - textW, y: margin, w: textW, h: slide.h - margin * 2 },
      imageRect: {
        x: margin,
        y: margin,
        w: slide.w - margin * 2 - textW - gap,
        h: slide.h - margin * 2
      }
    };
  }

  const textH = Math.max(textMin, (slide.h - margin * 2) * textRatio);
  if (position === "top") {
    return {
      textRect: { x: margin, y: margin, w: slide.w - margin * 2, h: textH },
      imageRect: {
        x: margin,
        y: margin + textH + gap,
        w: slide.w - margin * 2,
        h: slide.h - margin * 2 - textH - gap
      }
    };
  }

  return {
    textRect: { x: margin, y: slide.h - margin - textH, w: slide.w - margin * 2, h: textH },
    imageRect: {
      x: margin,
      y: margin,
      w: slide.w - margin * 2,
      h: slide.h - margin * 2 - textH - gap
    }
  };
}

function toPercentRect(rect) {
  const slide = { w: 11, h: 8.5 };
  return {
    left: `${(rect.x / slide.w) * 100}%`,
    top: `${(rect.y / slide.h) * 100}%`,
    width: `${(rect.w / slide.w) * 100}%`,
    height: `${(rect.h / slide.h) * 100}%`
  };
}

function clearPreviewUrls() {
  previewObjectUrls.forEach((url) => URL.revokeObjectURL(url));
  previewObjectUrls = [];
}

const EVEN_PAGE_MIN_FONT_PX = 12;
const EVEN_PAGE_INITIAL_FONT_PX = 34;

/**
 * Reduces even-page story font size until content fits within the page (no overflow).
 * @param {HTMLElement} storyWrapper - The .preview-story container for the even page
 */
function shrinkEvenPageStoryToFit(storyWrapper) {
  if (!storyWrapper || !storyWrapper.parentElement) return;
  const salient = storyWrapper.querySelector(".preview-salient-features");
  let fontSize = EVEN_PAGE_INITIAL_FONT_PX;
  storyWrapper.style.fontSize = `${fontSize}px`;
  if (salient) salient.style.fontSize = `${fontSize}px`;

  while (fontSize > EVEN_PAGE_MIN_FONT_PX) {
    if (storyWrapper.scrollHeight <= storyWrapper.clientHeight) break;
    fontSize -= 2;
    storyWrapper.style.fontSize = `${fontSize}px`;
    if (salient) salient.style.fontSize = `${fontSize}px`;
  }
}

function renderPreview() {
  clearPreviewUrls();
  previewContainer.innerHTML = "";

  const spreads = collectSpreadsFromForm();
  if (!spreads.length) {
    return;
  }

  const title = (bookTitleInput.value || "CVI Book").trim();
  const oddTextPosition = oddTextPositionInput.value;
  const oddTextSize = Math.max(18, Number(oddTextSizeInput.value) || 72);
  const oddTextColor = oddTextColorInput.value;
  const oddBorderColor = oddBorderColorInput.value;
  const oddBgColor = oddBgColorInput.value;
  const storyTextColor = storyTextColorInput.value;
  const oddBorderSize = Math.max(0, Number(oddBorderSizeInput.value) || 0);
  const outline = getTextOutlineCss(oddBorderSize, oddBorderColor);

  const titlePage = document.createElement("div");
  titlePage.className = "preview-page odd";
  titlePage.style.background = oddBgColor;
  titlePage.innerHTML = `<div class="preview-story" style="transform:none;color:${oddTextColor};text-shadow:${outline};font-size:52px;">${title}</div>`;
  previewContainer.appendChild(titlePage);

  spreads.forEach((spread, spreadIndex) => {
    const evenPage = document.createElement("div");
    evenPage.className = "preview-page";
    evenPage.style.background = "#FFFFFF";
    const storyWrapper = document.createElement("div");
    storyWrapper.className = "preview-story";
    storyWrapper.style.flexDirection = "column";
    storyWrapper.style.gap = "0.5em";
    storyWrapper.style.fontSize = "34px";
    const storyText = document.createElement("span");
    storyText.style.color = storyTextColor;
    storyText.textContent = spread.storyText || "";
    storyWrapper.appendChild(storyText);
    if (spread.salientFeatures) {
      const salient = document.createElement("div");
      salient.className = "preview-salient-features";
      salient.style.color = "#CC0000";
      salient.style.fontSize = "34px";
      salient.textContent = spread.salientFeatures;
      storyWrapper.appendChild(salient);
    }
    evenPage.appendChild(storyWrapper);
    previewContainer.appendChild(evenPage);

    requestAnimationFrame(() => shrinkEvenPageStoryToFit(storyWrapper));

    const evenLabel = document.createElement("p");
    evenLabel.className = "preview-label";
    evenLabel.textContent = `Spread ${spreadIndex + 1} - Even page (rotated)`;
    previewContainer.appendChild(evenLabel);

    const oddPage = document.createElement("div");
    oddPage.className = "preview-page odd";
    oddPage.style.background = oddBgColor;

    const regions = getOddRegions(oddTextPosition);
    const textRect = toPercentRect(regions.textRect);
    const imageRect = toPercentRect(regions.imageRect);

    const oddTextZone = document.createElement("div");
    oddTextZone.className = "preview-zone preview-odd-text";
    oddTextZone.style.left = textRect.left;
    oddTextZone.style.top = textRect.top;
    oddTextZone.style.width = textRect.width;
    oddTextZone.style.height = textRect.height;
    oddTextZone.style.fontSize = `${oddTextSize}px`;
    oddTextZone.style.color = oddTextColor;
    oddTextZone.style.textShadow = outline;
    oddTextZone.textContent = spread.oddText || "";

    const oddImagesZone = document.createElement("div");
    oddImagesZone.className = "preview-zone preview-odd-images";
    oddImagesZone.style.left = imageRect.left;
    oddImagesZone.style.top = imageRect.top;
    oddImagesZone.style.width = imageRect.width;
    oddImagesZone.style.height = imageRect.height;

    const imageFiles = spread.imageFiles.slice(0, 4);
    oddImagesZone.classList.add(`count-${Math.max(1, imageFiles.length)}`);
    if (visualComplexityInput.value === "highSupport") {
      oddImagesZone.classList.add("high-support");
    }
    imageFiles.forEach((file) => {
      const wrap = document.createElement("div");
      wrap.className = "preview-img-wrap";
      const img = document.createElement("img");
      const url = URL.createObjectURL(file);
      previewObjectUrls.push(url);
      img.src = url;
      img.alt = file.name;
      wrap.appendChild(img);
      oddImagesZone.appendChild(wrap);
    });

    oddPage.appendChild(oddTextZone);
    oddPage.appendChild(oddImagesZone);
    previewContainer.appendChild(oddPage);

    const oddLabel = document.createElement("p");
    oddLabel.className = "preview-label";
    oddLabel.textContent = `Spread ${spreadIndex + 1} - Odd page`;
    previewContainer.appendChild(oddLabel);
  });
}

/**
 * Appends a final slide with standard CVI teaching activity extensions.
 * @param {PptxGenJS} pptx - The PptxGenJS instance
 * @param {Object} options - Styling options
 * @param {string} options.oddBgColor - Background color (hex without #)
 * @param {string} options.oddTextColor - Text color (hex without #)
 * @param {string} options.oddBorderColor - Text outline color (hex without #)
 * @param {number} options.oddBorderSize - Text outline size in pt
 * @param {string} [options.activityPrompt] - Optional story-specific activity from AI
 */
function addFinalActivitySlide(pptx, options) {
  const { oddBgColor, oddTextColor, oddBorderColor, oddBorderSize, activityPrompt } = options;
  const TITLE_FONT_SIZE = 44;
  const HEADING_FONT_SIZE = 28;
  const BODY_FONT_SIZE = 22;

  const slide = pptx.addSlide();
  slide.background = { color: oddBgColor };

  slide.addText("Activity Extensions", {
    x: 0.5,
    y: 0.5,
    w: 10,
    h: 1,
    align: "center",
    valign: "mid",
    bold: true,
    fontSize: TITLE_FONT_SIZE,
    color: oddTextColor,
    outline: { color: oddBorderColor, pt: oddBorderSize }
  });

  const activities = [
    {
      title: "Gather, Discuss & Explore",
      body: "Collect real objects related to the story. Discuss textures, shapes, and features. Let the child explore items in a clutter-free space."
    },
    {
      title: "Create a Feely Box",
      body: "Place story-related objects in a box with a hand-sized opening. The child reaches in to feel and identify items by touch, reinforcing concepts."
    }
  ];

  if (activityPrompt && activityPrompt.trim()) {
    activities.push({
      title: "Story-Specific Activity",
      body: activityPrompt.trim()
    });
  }

  const startY = 2;
  const availableHeight = 6;
  const gap = 0.25;
  const activityHeight = (availableHeight - gap * (activities.length - 1)) / activities.length;

  activities.forEach((activity, i) => {
    const y = startY + i * (activityHeight + gap);
    slide.addText(activity.title, {
      x: 0.6,
      y,
      w: 9.8,
      h: 0.45,
      bold: true,
      fontSize: HEADING_FONT_SIZE,
      color: oddTextColor,
      outline: { color: oddBorderColor, pt: Math.max(0, oddBorderSize - 1) }
    });
    slide.addText(activity.body, {
      x: 0.6,
      y: y + 0.5,
      w: 9.8,
      h: Math.max(0.8, activityHeight - 0.5),
      fontSize: BODY_FONT_SIZE,
      color: oddTextColor,
      valign: "top",
      fit: "shrink"
    });
  });
}

async function exportPptx() {
  if (typeof PptxGenJS === "undefined") {
    setStatus("PowerPoint library did not load. Refresh and try again.", true);
    return;
  }

  const spreads = collectSpreadsFromForm();
  if (!spreads.length) {
    setStatus("Add at least one spread before exporting.", true);
    return;
  }

  const position = oddTextPositionInput.value;
  const oddTextColor = safePptColor(oddTextColorInput.value);
  const oddBorderColor = safePptColor(oddBorderColorInput.value);
  const oddBgColor = safePptColor(oddBgColorInput.value);
  const storyTextColor = safePptColor(storyTextColorInput.value);
  const oddTextSize = Number(oddTextSizeInput.value) || 72;
  const oddBorderSize = Number(oddBorderSizeInput.value) || 0;
  const bookTitle = (bookTitleInput.value || "CVI Book").trim();

  exportPptxButton.disabled = true;
  setStatus("Preparing images...");

  try {
    const spreadsWithData = [];
    for (const spread of spreads) {
      const imageData = [];
      for (const file of spread.imageFiles) {
        const data = await readFileAsDataUrl(file);
        imageData.push(data);
      }
      spreadsWithData.push({ ...spread, imageData });
    }

    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: "LETTER_LAND", width: 11, height: 8.5 });
    pptx.layout = "LETTER_LAND";
    pptx.author = "CVI Book Builder";
    pptx.subject = "CVI Story Book";
    pptx.title = bookTitle;

    setStatus("Building PowerPoint...");

    const TITLE_FONT_SIZE = 52;
    const STORY_FONT_SIZE = 34;
    const SALIENT_FEATURES_COLOR = "CC0000"; // High-contrast red on white
    const EVEN_PAGE_BG = "FFFFFF";

    const titleSlide = pptx.addSlide();
    titleSlide.background = { color: oddBgColor };
    titleSlide.addText(bookTitle, {
      x: 0.5,
      y: 2.8,
      w: 10,
      h: 2,
      align: "center",
      valign: "mid",
      bold: true,
      fontSize: TITLE_FONT_SIZE,
      color: oddTextColor,
      outline: { color: oddBorderColor, pt: oddBorderSize }
    });

    for (const spread of spreadsWithData) {
      const evenSlide = pptx.addSlide();
      evenSlide.background = { color: EVEN_PAGE_BG };

      const evenTextRuns = [];
      if (spread.storyText) {
        evenTextRuns.push({
          text: spread.storyText + (spread.salientFeatures ? "\n\n" : ""),
          options: { color: storyTextColor, fontSize: STORY_FONT_SIZE }
        });
      }
      if (spread.salientFeatures) {
        evenTextRuns.push({
          text: spread.salientFeatures,
          options: { color: SALIENT_FEATURES_COLOR, fontSize: STORY_FONT_SIZE }
        });
      }
      if (evenTextRuns.length === 0) {
        evenTextRuns.push({ text: " ", options: { color: storyTextColor, fontSize: STORY_FONT_SIZE } });
      }

      evenSlide.addText(evenTextRuns, {
        x: 0.55,
        y: 0.55,
        w: 9.9,
        h: 7.4,
        align: "center",
        valign: "mid",
        margin: 0.08,
        rotate: 180,
        fit: "shrink"
      });

      const oddSlide = pptx.addSlide();
      oddSlide.background = { color: oddBgColor };
      const { textRect, imageRect } = getOddRegions(position);

      if (spread.oddText) {
        oddSlide.addText(spread.oddText, {
          x: textRect.x,
          y: textRect.y,
          w: textRect.w,
          h: textRect.h,
          align: "center",
          valign: "mid",
          bold: true,
          fontSize: oddTextSize,
          color: oddTextColor,
          fit: "shrink",
          outline: { color: oddBorderColor, pt: Math.max(3, oddBorderSize) }
        });
      }

      const useImages = spread.imageData.slice(0, 4);
      const highSupport = visualComplexityInput.value === "highSupport";
      const imgRects = distributeImagesInRect(useImages.length, imageRect, { highSupport });
      useImages.forEach((img, i) => {
        const r = imgRects[i];
        oddSlide.addImage({
          data: img,
          sizing: { type: "contain", x: r.x, y: r.y, w: r.w, h: r.h }
        });
      });
    }

    if (includeTeachingActivitiesInput.checked) {
      addFinalActivitySlide(pptx, {
        oddBgColor,
        oddTextColor,
        oddBorderColor,
        oddBorderSize,
        activityPrompt: (activityPromptInput.value || "").trim()
      });
    }

    const fileName = `${bookTitle.replace(/[^a-z0-9-_ ]/gi, "").trim() || "cvi-book"}.pptx`;
    await pptx.writeFile({ fileName });
    setStatus(`Done. Downloaded ${fileName}`);
  } catch (err) {
    console.error(err);
    setStatus(`Export failed: ${err.message || "Unknown error"}`, true);
  } finally {
    exportPptxButton.disabled = false;
  }
}

const AI_FORMATTING_PROMPT_TEMPLATE = `You are helping create CVI-appropriate story books for children with cortical visual impairment. Output your response using EXACTLY these tags. Do not add any other formatting or commentary.

TITLE: [Book title]

ECC_AREA: [ECC area focus, e.g., Compensatory Access, Self-Determination]

ACTIVITY_PROMPT: [One-sentence sensory activity suggestion, e.g., Create a feely box with a ball and a dog toy.]

SPREAD:
STORY: [Story text for the even page. Simple, concrete language. One or two sentences per spread.]
ODD_TEXT: [Single keyword or short phrase for the odd slide—the main concept or object]
SALIENT_FEATURES: [2–4 salient features: shape, color, texture, movement, or other visually distinctive qualities. Comma-separated.]
IMAGE_PROMPT: [Short prompt for an AI image generator: describe one clear, simple image for this spread. CVI-friendly: high contrast, uncluttered, single main subject. Example: A single red ball on plain white background, soft lighting.]

SPREAD:
STORY: [Next spread's story text.]
ODD_TEXT: [Keyword for this spread.]
SALIENT_FEATURES: [Salient features for this spread.]
IMAGE_PROMPT: [AI image generator prompt for this spread.]

[Repeat SPREAD / STORY / ODD_TEXT / SALIENT_FEATURES / IMAGE_PROMPT for each spread.]`;

function getSelectedOptionLabel(selectElement) {
  if (!selectElement) return "";
  const selected = selectElement.options[selectElement.selectedIndex];
  return selected ? selected.textContent : "";
}

function buildAiFormattingPromptWithSetup() {
  const bookTitle = (bookTitleInput.value || "").trim();
  const eccArea = (eccAreaInput.value || "").trim();
  const activityPrompt = (activityPromptInput.value || "").trim();
  const oddTextPosition = getSelectedOptionLabel(oddTextPositionInput) || oddTextPositionInput.value;
  const visualComplexity = getSelectedOptionLabel(visualComplexityInput) || visualComplexityInput.value;
  const includeActivities = includeTeachingActivitiesInput.checked ? "Yes" : "No";

  return `Use the current Book Setup values below when you generate content.

Current Book Setup
- TITLE preference: ${bookTitle || "[Book title]"}
- ECC_AREA preference: ${eccArea || "[ECC area]"}
- ACTIVITY_PROMPT preference: ${activityPrompt || "[One-sentence sensory activity]"}
- Odd page text position: ${oddTextPosition}
- Visual complexity: ${visualComplexity}
- Include Teaching Activities: ${includeActivities}

${AI_FORMATTING_PROMPT_TEMPLATE}`.trim();
}

function copyToClipboardFallback(text, successMsg, failMsg) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(
      () => setStatus(successMsg),
      () => setStatus(failMsg, true)
    );
  } else {
    try {
      const textArea = document.createElement("textarea");
      textArea.value = text;
      textArea.style.top = "0";
      textArea.style.left = "0";
      textArea.style.position = "fixed";
      document.body.appendChild(textArea);
      textArea.focus();
      textArea.select();
      const successful = document.execCommand('copy');
      document.body.removeChild(textArea);
      if (successful) {
        setStatus(successMsg);
      } else {
        setStatus(failMsg, true);
      }
    } catch (err) {
      console.error('Fallback: Oops, unable to copy', err);
      setStatus(failMsg, true);
    }
  }
}

function copyAiFormattingPrompt() {
  const prompt = buildAiFormattingPromptWithSetup();
  copyToClipboardFallback(
    prompt,
    "AI prompt + current Book Setup copied. Paste into any AI, then paste the AI output back here.",
    "Failed to copy. Check clipboard permissions."
  );
}

function copyFvaLmaSummary() {
  const bookTitle = (bookTitleInput.value || "CVI Book").trim();
  const eccArea = (eccAreaInput.value || "").trim();
  const activityPrompt = (activityPromptInput.value || "").trim();
  const studentTips = (studentCviTipsInput.value || "").trim();
  const spreadCount = spreadsContainer.querySelectorAll(".spread-card").length;
  const date = new Date().toLocaleDateString();

  const textColor = oddTextColorInput.value || "#FFFFFF";
  const bgColor = oddBgColorInput.value || "#000000";
  const textSize = oddTextSizeInput.value || "72";
  const complexityValue = visualComplexityInput.value || "normal";
  const complexityLabel = complexityValue === "highSupport" ? "High Support" : "Normal";

  const parts = [
    `CVI Book: ${bookTitle}`,
    eccArea ? `ECC Focus: ${eccArea}` : "",
    `Spreads: ${spreadCount}`,
    `Date: ${date}`
  ].filter(Boolean);

  parts.push(
    "",
    "Visual Presentation",
    `Materials were presented using ${textColor} text at ${textSize}pt font on a ${bgColor} background. Images were spaced using ${complexityLabel} complexity settings.`
  );

  if (activityPrompt) {
    parts.push("", `Activity: ${activityPrompt}`);
  }

  if (studentTips) {
    parts.push("", "Student-specific CVI tips:", studentTips);
  }

  parts.push(
    "",
    "Latency to visual attention: ___ seconds",
    "Preferred viewing distance: ___ inches"
  );

  const summary = parts.join("\n");

  copyToClipboardFallback(
    summary,
    "FVA/LMA summary copied to clipboard",
    "Failed to copy. Check clipboard permissions."
  );
}

addSpreadButton.addEventListener("click", () => addSpread());

parseAiButton.addEventListener("click", () => {
  const raw = aiInput.value.trim();
  if (!raw) {
    setStatus("Paste formatted text first, then parse.", true);
    return;
  }

  const parsed = parseAiFormattedText(raw);
  if (parsed.title) {
    bookTitleInput.value = parsed.title;
  }
  if (parsed.eccArea) {
    eccAreaInput.value = parsed.eccArea;
  }
  if (parsed.activityPrompt) {
    activityPromptInput.value = parsed.activityPrompt;
  }

  spreadsContainer.innerHTML = "";
  parsed.spreads.forEach((s) => addSpread(s));

  if (!parsed.spreads.length) {
    addSpread();
    setStatus("No spreads found. Check your formatting and edit manually.", true);
    renderPreview();
    return;
  }
  setStatus(`Parsed ${parsed.spreads.length} spread(s). Add images as needed.`);
  renderPreview();
});

exportPptxButton.addEventListener("click", exportPptx);
copyFvaLmaButton.addEventListener("click", copyFvaLmaSummary);
copyAiPromptButton.addEventListener("click", copyAiFormattingPrompt);

presetMaxContrastButton.addEventListener("click", () => {
  oddTextColorInput.value = "#FFFF00";
  oddBgColorInput.value = "#000000";
  oddBorderColorInput.value = "#FFFF00";
  storyTextColorInput.value = "#111111";
  ["oddTextColor", "oddBorderColor", "oddBgColor", "storyTextColor"].forEach(syncColorDisplay);
  renderPreview();
});

presetHighContrastButton.addEventListener("click", () => {
  oddTextColorInput.value = "#FF0000";
  oddBgColorInput.value = "#000000";
  oddBorderColorInput.value = "#FF0000";
  storyTextColorInput.value = "#111111";
  ["oddTextColor", "oddBorderColor", "oddBgColor", "storyTextColor"].forEach(syncColorDisplay);
  renderPreview();
});

presetStandardPrintButton.addEventListener("click", () => {
  oddTextColorInput.value = "#000000";
  oddBgColorInput.value = "#FFFFFF";
  oddBorderColorInput.value = "#000000";
  storyTextColorInput.value = "#000000";
  ["oddTextColor", "oddBorderColor", "oddBgColor", "storyTextColor"].forEach(syncColorDisplay);
  renderPreview();
});

refreshPreviewButton.addEventListener("click", () => {
  renderPreview();
  setStatus("Preview refreshed.");
});
printPdfButton.addEventListener("click", () => {
  renderPreview();
  printablePreviewSection.classList.add("preview-printable");
  window.print();
  printablePreviewSection.classList.remove("preview-printable");
});

if (closeWelcomeButton) {
  closeWelcomeButton.addEventListener("click", hideWelcomeModal);
}

if (welcomeModal) {
  welcomeModal.addEventListener("click", (e) => {
    const target = e.target;
    if (!(target instanceof Element)) return;

    if (target === welcomeModal || target.closest("#closeWelcomeButton")) {
      hideWelcomeModal();
    }
  });
}

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && welcomeModal && !welcomeModal.hidden) {
    hideWelcomeModal();
  }
});

initColorPickers();
initImageIsolator();

function initLivePreview() {
  if (bookTitleInput) bookTitleInput.addEventListener("input", () => scheduleLivePreview());
  if (oddTextPositionInput) oddTextPositionInput.addEventListener("change", () => scheduleLivePreview(true));
  if (oddTextSizeInput) oddTextSizeInput.addEventListener("input", () => scheduleLivePreview(true));
  if (oddBorderSizeInput) oddBorderSizeInput.addEventListener("input", () => scheduleLivePreview(true));
  if (visualComplexityInput) visualComplexityInput.addEventListener("change", () => scheduleLivePreview(true));
  spreadsContainer.addEventListener("input", (e) => {
    if (e.target.matches(".story-text, .salient-features, .odd-text, .image-prompt")) scheduleLivePreview();
  });
}

initLivePreview();

document.body.addEventListener("click", (e) => {
  if (!(e.target instanceof Element)) return;
  const btn = e.target.closest("button");
  if (btn && !btn.disabled && !window.matchMedia("(prefers-reduced-motion: reduce)").matches) {
    btn.classList.add("animate-pop");
    btn.addEventListener("animationend", () => btn.classList.remove("animate-pop"), { once: true });
  }
});

addSpread();
renderPreview();
showWelcomeModal();
setStatus("Ready. Add spreads and export PowerPoint.");

window.addEventListener("beforeunload", () => {
  clearPreviewUrls();
  revokeCurrentSourceObjectUrl();
  revokeProcessedResultUrl();
});
