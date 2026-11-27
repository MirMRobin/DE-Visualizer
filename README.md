# Differential Evolution Visualizer

A browser-based research workbench for Differential Evolution and general research figures. The Evolution Lab explores named DE method profiles; Data Studio imports common research files, reports data-quality signals, and builds configurable descriptive figures. Research Workflow turns the current profile and figure settings into a reviewable analysis plan and provenance record. The main parsers and plotting engine are bundled in `vendor/`. Uploaded source files are processed in the browser.

## Open the application

For a full local preview, including PDF workers and OCR, serve the folder over HTTP:

```sh
python3 -m http.server 8000
```

Then open `http://localhost:8000`. GitHub Pages is also supported. Browser security restrictions may limit web workers on `file://` URLs.

Use **Guide** for the file workflow and limits, and **Methods** for algorithm and figure details. The page includes a project card crediting Taharimul Islam.

## Data Studio file support

- CSV, TSV, JSON, TXT, and Markdown.
- Excel XLSX, XLS, XLSM, and ODS, including worksheet selection. Formula cells use their cached workbook values; macros are not executed and formulas are not recalculated.
- Word DOCX, including text, headings, lists, and selectable tables. Complex layouts and merged cells need source review.
- Text-based PDF extraction up to 50 pages. Scanned PDFs can use optional OCR for up to 10 pages.
- PNG, JPEG, WebP, BMP, and GIF image profiling. Pixel data is downsampled to at most 10,000 samples and can be viewed as a spatial brightness map.

Limits: 25 MB per file, 50,000 rows, 1,000 fields, and 1,000,000 cells in the figure workspace. Oversized sources are capped and reported. See [guide.html](guide.html) for limitations, workflow, and privacy details. The Research Workflow can export a JSON provenance record or Markdown plan, including source checksum when available, profile signals, chart settings, user notes, and an explicitly self-reported review checklist.

## Data profiling and figures

Profiling runs in a Web Worker where supported. The report includes numeric and categorical fields, mixed-type fields, category levels, missing cells, duplicate rows, numeric ranges, and exploratory 1.5×IQR flags. Missing cells are not imputed; duplicate rows are not removed. The recommendations are a transparent data-shape heuristic, not a scientific interpretation or statistical test.

Available figures: line, area, scatter, bubble, category bar, grouped bar, histogram, empirical CDF, box plot, Pearson correlation heatmap, Pareto chart, image intensity map, data table, concept map, and process sequence. Choose axes, optional groups, aggregation, palette, labels, title, type size, and export dimensions.

Export vector SVG, PNG at 2× or 4×, CSV, TSV, JSON, Excel, figure JSON, standalone static HTML reports, Markdown, Mermaid, and OPML (the latter two are enabled for concept and process views). Figure JSON contains imported rows and the selected figure settings. For large data, plotted series may be deterministically sampled; the caption says when this applies, and data exports preserve every row included in the workspace.

OCR is the only feature that fetches runtime assets from a CDN. It downloads the OCR engine and English language model on demand. The selected image/PDF is not sent to that CDN; recognition is performed in the browser.

## Evolution Lab

Choose among 13 named DE profiles plus a configurable baseline: DE/rand/1/bin, DE/best/1/bin, DE/current-to-best/1/bin, DE/rand/2/bin, DE/current-to-rand/1, jDE, JADE, SHADE, L-SHADE, jSO, SaDE, CoDE, and EPSDE. The lab supports dimensions 2–100, population sizes 8–250, F and CR values from 0 to 1, and objective-evaluation budgets through 10,000.

These are compact educational profiles, not certified reference implementations or exact competition configurations. No CEC suite is run. See [methods.html](methods.html) and [docs-method.md](docs-method.md) for mechanisms, caveats, and primary references.

## Third-party browser libraries

Pinned local bundles: Apache ECharts 6.1.0, SheetJS 0.20.3, Mammoth 1.13.0, and PDF.js 3.11.174. Their license texts are in `vendor/licenses/`. Tesseract.js 7.0.0 and English language data are loaded on demand from pinned CDN URLs; see [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

## License

The application is MIT licensed. See [LICENSE](LICENSE). Third-party dependencies retain their own licenses.
