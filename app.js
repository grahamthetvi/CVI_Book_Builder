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
const aiInput = document.getElementById("aiInput");
const printablePreviewSection = previewContainer.closest(".panel");

let previewObjectUrls = [];

function setStatus(text, isError = false) {
  statusMessage.textContent = text;
  statusMessage.style.color = isError ? "#ff9e9e" : "#c6d0dd";
}

function safePptColor(hex) {
  if (!hex) return "000000";
  return hex.replace("#", "").toUpperCase();
}

function addSpread(initial = {}) {
  const fragment = spreadTemplate.content.cloneNode(true);
  const card = fragment.querySelector(".spread-card");
  const storyText = fragment.querySelector(".story-text");
  const oddText = fragment.querySelector(".odd-text");
  const imageInput = fragment.querySelector(".odd-images");
  const selectedImagesText = fragment.querySelector(".selected-images");
  const removeButton = fragment.querySelector(".remove-spread");

  storyText.value = initial.storyText || "";
  oddText.value = initial.oddText || "";

  imageInput.addEventListener("change", () => {
    const files = Array.from(imageInput.files || []);
    if (files.length > 4) {
      setStatus("Only 1 to 4 images are allowed on odd pages. Kept first 4.", true);
      const dt = new DataTransfer();
      files.slice(0, 4).forEach((f) => dt.items.add(f));
      imageInput.files = dt.files;
    }
    const finalFiles = Array.from(imageInput.files || []);
    selectedImagesText.textContent = finalFiles.length
      ? `${finalFiles.length} image(s): ${finalFiles.map((f) => f.name).join(", ")}`
      : "No images selected.";
  });

  removeButton.addEventListener("click", () => {
    card.remove();
    renumberSpreads();
  });

  spreadsContainer.appendChild(fragment);
  renumberSpreads();
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

  const chunks = raw.split(/^\s*SPREAD:\s*$/gim).map((c) => c.trim()).filter(Boolean);
  const spreads = [];

  for (const chunk of chunks) {
    const storyMatch = chunk.match(/^\s*STORY:\s*([\s\S]*?)(?=^\s*ODD_TEXT:|\s*$)/im);
    const oddMatch = chunk.match(/^\s*ODD_TEXT:\s*(.+)\s*$/im);
    const storyText = storyMatch ? storyMatch[1].trim() : "";
    const oddText = oddMatch ? oddMatch[1].trim() : "";
    if (storyText || oddText) {
      spreads.push({ storyText, oddText });
    }
  }

  return { title, spreads };
}

function collectSpreadsFromForm() {
  const cards = Array.from(spreadsContainer.querySelectorAll(".spread-card"));
  return cards.map((card, index) => {
    const storyText = card.querySelector(".story-text").value.trim();
    const oddText = card.querySelector(".odd-text").value.trim();
    const imageFiles = Array.from(card.querySelector(".odd-images").files || []).slice(0, 4);
    return {
      index,
      storyText,
      oddText,
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

function distributeImagesInRect(count, area) {
  const gap = 0.12;
  const rects = [];

  if (count <= 0) return rects;
  if (count === 1) return [{ ...area }];

  if (count === 2) {
    const w = (area.w - gap) / 2;
    rects.push({ x: area.x, y: area.y, w, h: area.h });
    rects.push({ x: area.x + w + gap, y: area.y, w, h: area.h });
    return rects;
  }

  if (count === 3) {
    const topH = area.h * 0.54;
    const bottomH = area.h - topH - gap;
    rects.push({ x: area.x, y: area.y, w: area.w, h: topH });
    const bottomW = (area.w - gap) / 2;
    rects.push({ x: area.x, y: area.y + topH + gap, w: bottomW, h: bottomH });
    rects.push({ x: area.x + bottomW + gap, y: area.y + topH + gap, w: bottomW, h: bottomH });
    return rects;
  }

  const w = (area.w - gap) / 2;
  const h = (area.h - gap) / 2;
  rects.push({ x: area.x, y: area.y, w, h });
  rects.push({ x: area.x + w + gap, y: area.y, w, h });
  rects.push({ x: area.x, y: area.y + h + gap, w, h });
  rects.push({ x: area.x + w + gap, y: area.y + h + gap, w, h });
  return rects;
}

function getOddRegions(position) {
  const slide = { w: 11, h: 8.5 };
  const margin = 0.35;
  const gap = 0.14;
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

    const evenPage = document.createElement("div");
    evenPage.className = "preview-page";
    const story = document.createElement("div");
    story.className = "preview-story";
    story.style.color = storyTextColor;
    story.textContent = spread.storyText || "";
    evenPage.appendChild(story);
    previewContainer.appendChild(evenPage);

    const evenLabel = document.createElement("p");
    evenLabel.className = "preview-label";
    evenLabel.textContent = `Spread ${spreadIndex + 1} - Even page (rotated)`;
    previewContainer.appendChild(evenLabel);
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
      fontSize: 52,
      color: oddTextColor,
      outline: { color: oddBorderColor, pt: oddBorderSize }
    });

    for (const spread of spreadsWithData) {
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
          outline: { color: oddBorderColor, pt: oddBorderSize }
        });
      }

      const useImages = spread.imageData.slice(0, 4);
      const imgRects = distributeImagesInRect(useImages.length, imageRect);
      useImages.forEach((img, i) => {
        const r = imgRects[i];
        oddSlide.addImage({
          data: img,
          sizing: { type: "contain", x: r.x, y: r.y, w: r.w, h: r.h }
        });
      });

      const evenSlide = pptx.addSlide();
      evenSlide.background = { color: "FFFFFF" };
      evenSlide.addText(spread.storyText || " ", {
        x: 0.55,
        y: 0.55,
        w: 9.9,
        h: 7.4,
        align: "center",
        valign: "mid",
        margin: 0.08,
        fontSize: 34,
        color: storyTextColor,
        rotate: 180,
        fit: "shrink"
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

addSpread();
renderPreview();
setStatus("Ready. Add spreads and export PowerPoint.");
