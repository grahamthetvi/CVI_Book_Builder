# CVI Book Builder (V1)

Static GitHub Pages app that builds a downloadable `.pptx` CVI story book.

## V1 features

- Letter landscape PowerPoint output (`11 x 8.5`)
- Print-friendly preview with browser Save as PDF path
- Title slide
- Alternating spread export:
  - odd slide: accessible high-contrast text + 1 to 4 images
  - even slide: story text rotated 180 degrees
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

## AI formatting template

```txt
TITLE: My Book Title
SPREAD:
STORY: Story text for the even page.
ODD_TEXT: KEYWORD

SPREAD:
STORY: Next story text.
ODD_TEXT: NEXT KEYWORD
```

After parsing, add images in each spread before export.
