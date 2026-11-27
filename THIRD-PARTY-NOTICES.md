# Third-party notices

The browser bundles listed below are included under their upstream licenses. Full license texts are stored in `vendor/licenses/`.

| Component | Version | License | Upstream |
|---|---:|---|---|
| Apache ECharts | 6.1.0 | Apache-2.0 | https://github.com/apache/echarts |
| SheetJS Community Edition | 0.20.3 | Apache-2.0 | https://github.com/SheetJS/sheetjs |
| Mammoth.js | 1.13.0 | BSD-2-Clause | https://github.com/mwilliamson/mammoth.js |
| PDF.js | 3.11.174 | Apache-2.0 | https://github.com/mozilla/pdf.js |

## Optional OCR assets

When a user chooses OCR, the page loads Tesseract.js 7.0.0 from `cdn.jsdelivr.net` and the English `@tesseract.js-data/eng` language model on demand. These assets are not bundled in this repository. They are executed by the user's browser; uploaded files are not sent to the CDN. See the [Tesseract.js project](https://github.com/naptha/tesseract.js) and its license for the applicable terms.
