# KMG Eurobond Investor Presentation Site

Static website for the JSC NC KazMunayGas Eurobond Investor Presentation, hosted on GitHub Pages.

## Structure

```
/
├── index.html          # Main page
├── styles.css          # Styles (petrol/gold design system)
├── app.js              # Gallery, lightbox, mobile nav
├── data.json           # Key metrics + slide list (edit to update numbers)
├── assets/
│   ├── slides/         # Slide previews: slide-01.jpg … slide-21.jpg
│   ├── og-image.jpg    # Social/messenger preview image
│   └── favicon.png
├── downloads/
│   └── KMG_Eurobond_Investor_Presentation.pptx
└── build/              # pptxgenjs source scripts
```

## Updating content

### Update key metrics
Edit `data.json` → `metrics` array. No HTML changes needed.

### Replace the presentation file
Drop the new file into `downloads/` with the same filename, then commit and push.

### Regenerate slide previews
1. Export `.pptx` to PDF via LibreOffice:
   ```
   soffice --headless --convert-to pdf downloads/KMG_Eurobond_Investor_Presentation.pptx
   ```
2. Convert PDF pages to JPEG:
   ```
   pdftoppm -jpeg -r 150 KMG_Eurobond_Investor_Presentation.pdf assets/slides/slide
   ```
3. Rename output to `slide-01.jpg`, `slide-02.jpg`, … `slide-21.jpg`

### Deploy
```
git add .
git commit -m "Update presentation content"
git push
```
GitHub Pages redeploys automatically within ~1 minute.
