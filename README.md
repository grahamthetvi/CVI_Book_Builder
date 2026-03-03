# CVI Book Builder (V2)

Static GitHub Pages app that builds a downloadable `.pptx` CVI story book.

## V2 features

- Letter landscape PowerPoint output (`11 x 8.5`)
- Print-friendly preview with browser Save as PDF path
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

## Use

1. Open `index.html` (or deploy root to GitHub Pages).
2. Set title and style options.
3. Add spreads manually, or paste formatted text and click **Parse Into Spreads**.
4. For each spread, optionally upload 1 to 4 odd-page images.
5. Click **Refresh Preview** to check layout.
6. Click **Print / Save as PDF** for PDF output, or **Download PowerPoint (.pptx)**.
7. Use **Copy FVA/LMA Summary** to copy session notes for the TVI.

## Gemini System Prompt

Use this prompt with Gemini (or similar AI) to generate formatted text for the CVI Book Builder. Paste the AI output into the **AI formatted text** field and click **Parse Into Spreads**.

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
