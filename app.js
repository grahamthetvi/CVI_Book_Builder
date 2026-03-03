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
  const salientFeatures = fragment.querySelector(".salient-features");
  const oddText = fragment.querySelector(".odd-text");
  const imageInput = fragment.querySelector(".odd-images");
  const selectedImagesText = fragment.querySelector(".selected-images");
  const removeButton = fragment.querySelector(".remove-spread");

  storyText.value = initial.storyText || "";
  salientFeatures.value = initial.salientFeatures || "";
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

  const eccMatch = raw.match(/^\s*ECC_AREA:\s*(.+)\s*$/im);
  const eccArea = eccMatch ? eccMatch[1].trim() : "";

  const activityMatch = raw.match(/^\s*ACTIVITY_PROMPT:\s*(.+)\s*$/im);
  const activityPrompt = activityMatch ? activityMatch[1].trim() : "";

  const chunks = raw.split(/^\s*SPREAD:\s*$/gim).map((c) => c.trim()).filter(Boolean);
  const spreads = [];

  for (const chunk of chunks) {
    const storyMatch = chunk.match(/^\s*STORY:\s*([\s\S]*?)(?=^\s*SALIENT_FEATURES:|^\s*ODD_TEXT:|\s*$)/im);
    const salientMatch = chunk.match(/^\s*SALIENT_FEATURES:\s*([\s\S]*?)(?=^\s*ODD_TEXT:|\s*$)/im);
    const oddMatch = chunk.match(/^\s*ODD_TEXT:\s*(.+)\s*$/im);
    const storyText = storyMatch ? storyMatch[1].trim() : "";
    const salientFeatures = salientMatch ? salientMatch[1].trim() : "";
    const oddText = oddMatch ? oddMatch[1].trim() : "";
    if (storyText || oddText || salientFeatures) {
      spreads.push({ storyText, salientFeatures, oddText });
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
    const imageFiles = Array.from(card.querySelector(".odd-images").files || []).slice(0, 4);
    return {
      index,
      storyText,
      salientFeatures,
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

    const evenLabel = document.createElement("p");
    evenLabel.className = "preview-label";
    evenLabel.textContent = `Spread ${spreadIndex + 1} - Even page (rotated)`;
    previewContainer.appendChild(evenLabel);
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
          fontSize: TITLE_FONT_SIZE,
          color: oddTextColor,
          fit: "shrink",
          outline: { color: oddBorderColor, pt: oddBorderSize }
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

const AI_FORMATTING_PROMPT = `You are helping create CVI-appropriate story books for children with cortical visual impairment. Output your response using EXACTLY these tags. Do not add any other formatting or commentary.

TITLE: [Book title]

ECC_AREA: [ECC area focus, e.g., Compensatory Access, Self-Determination]

ACTIVITY_PROMPT: [One-sentence sensory activity suggestion, e.g., Create a feely box with a ball and a dog toy.]

SPREAD:
STORY: [Story text for the even page. Simple, concrete language. One or two sentences per spread.]
ODD_TEXT: [Single keyword or short phrase for the odd slide—the main concept or object]
SALIENT_FEATURES: [2–4 salient features: shape, color, texture, movement, or other visually distinctive qualities. Comma-separated.]

SPREAD:
STORY: [Next spread's story text.]
ODD_TEXT: [Keyword for this spread.]
SALIENT_FEATURES: [Salient features for this spread.]

[Repeat SPREAD / STORY / ODD_TEXT / SALIENT_FEATURES for each spread.]`;

function copyAiFormattingPrompt() {
  navigator.clipboard.writeText(AI_FORMATTING_PROMPT).then(
    () => setStatus("AI formatting prompt copied. Paste into any AI, then paste the AI output back here."),
    () => setStatus("Failed to copy. Check clipboard permissions.", true)
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

  navigator.clipboard.writeText(summary).then(
    () => setStatus("FVA/LMA summary copied to clipboard"),
    () => setStatus("Failed to copy. Check clipboard permissions.", true)
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
  renderPreview();
});

presetHighContrastButton.addEventListener("click", () => {
  oddTextColorInput.value = "#FF0000";
  oddBgColorInput.value = "#000000";
  oddBorderColorInput.value = "#FF0000";
  storyTextColorInput.value = "#111111";
  renderPreview();
});

presetStandardPrintButton.addEventListener("click", () => {
  oddTextColorInput.value = "#000000";
  oddBgColorInput.value = "#FFFFFF";
  oddBorderColorInput.value = "#000000";
  storyTextColorInput.value = "#000000";
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

addSpread();
renderPreview();
setStatus("Ready. Add spreads and export PowerPoint.");
