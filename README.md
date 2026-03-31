# CVI Book Builder (V2)

Static GitHub Pages app that builds a downloadable `.pptx` CVI story book.

## V2 features

- Letter landscape PowerPoint output (`11 x 8.5`)
- Print-friendly preview with browser Save as PDF path
- **Export progress** — full-screen message and spinner while images are encoded and the PowerPoint file is built (large books can take a while)
- **Drafts on this device** — auto-saves when you export PowerPoint or open print/Save as PDF, plus optional named snapshots; load or delete any draft from the list (text and images are restored when storage allows)
- Title slide
- Alternating spread export:
  - odd slide: accessible high-contrast text + 1 to 4 images
  - even slide: story text rotated 180 degrees
- **Salient Features** — per-spread CVI salient feature descriptions on even pages
- **TVI Session Notes export** — Copy FVA/LMA Summary to clipboard with book metadata, visual presentation details, and blank tracking fields for latency and viewing distance
- **High Support visual complexity** — increased spacing between images on odd slides for reduced visual clutter
- **Automated Activity Extensions slide** — optional final slide with standard CVI activities plus a story-specific activity from the AI prompt
- User controls for:
  - odd-page text color
  - odd-page text size
  - odd-page border color
  - odd-page border size
  - odd-page background color (default black)
  - odd-page text position (top, bottom, left, right)
- AI formatted text parsing + manual spread editing
- Live preview refresh button before export/print

## Drafts and autosave

- The app saves an **auto-save** entry when you **Download PowerPoint (.pptx)** (after the file is written) or when you click **Print / Save as PDF** (before the print dialog opens). Recent auto-saves within about 45 seconds are merged into one entry so the list does not flood.
- Use **Save snapshot now** to pin the current book with an optional name.
- Up to **25** drafts are kept (oldest dropped). Drafts live in **this browser on this device only** (`localStorage`).
- Books with **many large images** may hit browser storage limits; the app shows an error if a save fails. Export a `.pptx` as a backup for important work.
- **Load** replaces the current book (you are asked to confirm). **Delete** removes only that draft.

## Use

1. Open `index.html` (or deploy root to GitHub Pages).
2. Set title and style options.
3. Add spreads manually, or paste formatted text and click **Parse Into Spreads**.
4. For each spread, optionally upload 1 to 4 odd-page images.
5. Click **Refresh Preview** to check layout.
6. Click **Print / Save as PDF** for PDF output, or **Download PowerPoint (.pptx)**.
7. Use **Copy FVA/LMA Summary** to copy session notes for the TVI.

### Print / Save as PDF

The app marks only the live preview for printing. The print stylesheet uses letter landscape. **Different browsers** may position margins slightly differently; if one browser’s PDF looks off, try another.

**Why “after print” matters (technical):** The code must know when printing has **finished** before removing the print-only styling. Some browsers run the actual print layout **after** `window.print()` has already returned. The app listens for the browser’s **`afterprint`** event (with a long fallback timer) so the preview panel stays in “print mode” until the dialog completes. That helps the saved PDF match what you saw in print preview.

## AI Formatting Prompt

Use this prompt with any AI assistant (Gemini, ChatGPT, Claude, etc.) to generate formatted text for the CVI Book Builder. Paste the AI output into the **AI formatted text** field and click **Parse Into Spreads**.

```
You are helping create CVI-appropriate story books for children with cortical visual impairment. Output your response using EXACTLY these tags. Do not add any other formatting or commentary.

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

[Repeat SPREAD / STORY / ODD_TEXT / SALIENT_FEATURES for each spread.]
```

### Tag reference

| Tag | Description |
|-----|-------------|
| `TITLE:` | Book title (one line) |
| `ECC_AREA:` | ECC area focus (one line) |
| `ACTIVITY_PROMPT:` | One-sentence sensory activity for the Activity Extensions slide |
| `SPREAD:` | Marks the start of a new spread |
| `STORY:` | Story text for the even page |
| `ODD_TEXT:` | Keyword or short phrase for the odd slide |
| `SALIENT_FEATURES:` | Comma-separated salient features for CVI |

After parsing, add images in each spread before export.

## License

**This repository (CVI Book Builder code written here)** is under the [MIT License](LICENSE).

### Third-party software

The static web app loads additional libraries from CDNs at runtime. Notable terms:

| Component | Use | License |
|-----------|-----|---------|
| [heic-to](https://github.com/hoppergee/heic-to) | HEIC/HEIF → JPEG in the browser (primary converter) | [LGPL-3.0](https://www.gnu.org/licenses/lgpl-3.0.html) |
| [heic2any](https://github.com/alexcorvi/heic2any) | HEIC fallback conversion | MIT |
| [PptxGenJS](https://github.com/gitbrent/PptxGenJS) | PowerPoint (`.pptx`) generation | MIT |

Copies of **LGPL-3.0** and **GPL-3.0** (the LGPL refers to the GPL) are in the [`licenses/`](licenses/) folder for anyone who receives this project. Source for LGPL components: [heic-to on GitHub](https://github.com/hoppergee/heic-to).

*This is not legal advice.* If you redistribute the app or change how libraries are bundled, double-check LGPL-3.0 obligations for your situation (for example, how the library is linked or loaded).
