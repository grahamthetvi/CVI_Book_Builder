/** CVI Book Nook Ready-to-Print PDF text helpers (no DOM). */

const SALIENT_RE = /salient\s*features\s*:/i;

/**
 * Join pdf.js getTextContent items into readable page text.
 * @param {{ str?: string, hasEOL?: boolean }[]} items
 */
export function joinPdfTextItems(items) {
  let out = "";
  for (const item of items || []) {
    const str = item?.str;
    if (!str) {
      if (item?.hasEOL) out += "\n";
      continue;
    }
    out += str;
    if (item.hasEOL) out += "\n";
    else out += " ";
  }
  return out.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}

/**
 * @param {string} text
 * @param {number} pageNumber 1-based
 * @returns {"cover"|"story"|"photo"}
 */
export function classifyNookPage(text, pageNumber) {
  const raw = text || "";
  if (SALIENT_RE.test(raw)) return "story";
  if (/book\s*cover/i.test(raw) || /created\s+by\b/i.test(raw)) return "cover";
  if (pageNumber === 1) return "cover";
  return "photo";
}

/**
 * @param {string} text
 */
export function parseNookCoverTitle(text) {
  let raw = (text || "").replace(/\s+/g, " ").trim();
  raw = raw.replace(/^book\s*cover\s*/i, "");
  raw = raw.replace(/\s*created\s+by\b.*$/i, "");
  raw = raw.replace(/^\d+\s*/, "").trim();
  return raw;
}

function stripLeadingPageNumber(text) {
  return (text || "")
    .replace(/^\s*\d+\s*/, "")
    .replace(/[\r\n]+\s*\d+\s*$/, "")
    .trim();
}

function normalizeNookProse(text) {
  return (text || "")
    .replace(/\s+/g, " ")
    .replace(/\s+([.,!?;:])/g, "$1")
    .replace(/([.!?])([A-Za-zÀ-ÿ])/g, "$1 $2")
    .trim();
}

/**
 * @param {string} text
 * @returns {{ storyText: string, salientFeatures: string }}
 */
export function parseNookStoryPage(text) {
  const cleaned = stripLeadingPageNumber(text);
  const parts = cleaned.split(/salient\s*features\s*:\s*/i);
  const storyText = normalizeNookProse(parts[0] || "");
  const salientFeatures = normalizeNookProse(parts.slice(1).join(" "));
  return { storyText, salientFeatures };
}

/**
 * Last word of the first sentence (e.g. "red", "plate").
 * @param {string} storyText
 */
export function deriveNookOddText(storyText) {
  const cleaned = normalizeNookProse(storyText);
  if (!cleaned) return "";
  const firstSentence = (cleaned.match(/^.+?[.!?]+/) || [cleaned])[0].trim();
  const words = firstSentence
    .replace(/[""''«»]+/g, "")
    .split(" ")
    .map((w) => w.replace(/^[.,!?;:]+|[.,!?;:]+$/g, ""))
    .filter(Boolean);
  const last = words[words.length - 1] || "";
  if (last) return last;
  const phrase = words.slice(0, 3).join(" ");
  return phrase.length > 28 ? phrase.slice(0, 28).trim() : phrase;
}

/**
 * @param {{ pageNumber: number, text: string }[]} pages
 * @returns {{ title: string, spreads: { storyText: string, salientFeatures: string, oddText: string, photoPageNumber: number|null }[] }}
 */
export function pairNookSpreads(pages) {
  const classified = (pages || []).map((p) => ({
    ...p,
    kind: classifyNookPage(p.text, p.pageNumber)
  }));

  let title = "";
  const spreads = [];

  for (let i = 0; i < classified.length; i += 1) {
    const cur = classified[i];
    if (cur.kind === "cover" && !title) {
      title = parseNookCoverTitle(cur.text);
    }
    if (cur.kind !== "story") continue;

    const parsed = parseNookStoryPage(cur.text);
    let photoPageNumber = null;
    const next = classified[i + 1];
    if (next && next.kind === "photo") {
      photoPageNumber = next.pageNumber;
      i += 1;
    }

    spreads.push({
      storyText: parsed.storyText,
      salientFeatures: parsed.salientFeatures,
      oddText: deriveNookOddText(parsed.storyText),
      photoPageNumber
    });
  }

  return { title, spreads };
}
