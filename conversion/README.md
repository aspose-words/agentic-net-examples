# Conversion Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **Conversion** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Conversion**
- Slug: **conversion**
- Total examples: **108**
- Publish-ready successful examples: **108 / 108**
- Native conversion API examples: **22**
- Input-bootstrap examples: **30**
- Save-options examples: **24**
- Stream / response simulation examples: **8**
- Batch / folder bootstrap examples: **24**

## Category rules that shaped these examples

- Use native Aspose.Words conversion APIs directly.
- Create realistic local sample inputs whenever the task mentions an existing file, stream, DOC, DOCX, HTML, TXT, XML, folder, or batch workflow.
- Use `Aspose.Drawing` instead of `System.Drawing` when drawing-related types are needed.
- Simulate network or browser responses with a `MemoryStream` instead of starting a real server.
- Validate that output files exist or output streams contain data before finishing.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`
- Aspose.Drawing.Common

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.Drawing.Common

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\conversion\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `conversion/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.Drawing.Common

# PowerShell example
Copy-Item ..\conversion\load-a-docx-file-from-disk-and-save-it-as-pdf-preserving-layout.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-from-disk-and-save-it-as-pdf-preserving-layout.cs` | Load a DOCX file from disk and save it as PDF preserving layout. | input-bootstrap | pdf, docx | verified |
| 2 | `convert-a-doc-file-to-pdf-a-1b-by-setting-pdfsaveoptions-compliance-before-saving.cs` | Convert a DOC file to PDF/A‑1b by setting PdfSaveOptions.Compliance before saving. | input-bootstrap | pdf, doc | verified |
| 3 | `stream-a-docx-document-into-memory-and-write-the-pdf-output-directly-to-response-stream.cs` | Stream a DOCX document into memory and write the PDF output directly to response stream. | stream-response | pdf, docx | verified |
| 4 | `batch-process-all-rtf-files-in-a-folder-converting-each-to-pdf-with-default-layout.cs` | Batch process all RTF files in a folder, converting each to PDF with default layout. | batch-bootstrap | pdf, rtf | verified |
| 5 | `load-an-html-page-from-a-url-and-convert-it-to-pdf-using-document-load-and-saveformat-pdf.cs` | Load an HTML page from a URL and convert it to PDF using Document.Load and SaveFormat.Pdf. | input-bootstrap | pdf, html | verified |
| 6 | `convert-a-png-image-to-pdf-by-loading-the-image-file-and-saving-with-saveformat-pdf.cs` | Convert a PNG image to PDF by loading the image file and saving with SaveFormat.Pdf. | native-conversion-api | pdf, png | verified |
| 7 | `transform-a-jpeg-picture-into-a-pdf-document-while-preserving-original-resolution-and-colo.cs` | Transform a JPEG picture into a PDF document while preserving original resolution and color depth. | native-conversion-api | pdf, jpeg | verified |
| 8 | `load-a-pdf-file-and-export-its-pages-as-separate-png-images-using-saveformat-png.cs` | Load a PDF file and export its pages as separate PNG images using SaveFormat.Png. | save-options | pdf, png | verified |
| 9 | `convert-a-pdf-document-to-high-quality-jpg-images-by-setting-jpegsaveoptions-quality-befor.cs` | Convert a PDF document to high‑quality JPG images by setting JpegSaveOptions.Quality before saving. | save-options | pdf, jpeg | verified |
| 10 | `save-a-pdf-as-pdf-ua-compliant-file-by-configuring-pdfsaveoptions-compliance-to-pdfcomplia.cs` | Save a PDF as PDF/UA compliant file by configuring PdfSaveOptions.Compliance to PdfCompliance.Ua. | save-options | pdf | verified |
| 11 | `load-a-pdf-and-export-it-to-docx-format-for-further-editing-using-saveformat-docx.cs` | Load a PDF and export it to DOCX format for further editing using SaveFormat.Docx. | save-options | pdf, docx, doc | verified |
| 12 | `transform-a-pdf-into-an-epub-e-book-by-loading-the-pdf-and-saving-with-saveformat-epub.cs` | Transform a PDF into an EPUB e‑book by loading the PDF and saving with SaveFormat.Epub. | native-conversion-api | pdf, epub | verified |
| 13 | `convert-an-html-file-to-pdf-while-preserving-css-styles-by-loading-html-and-using-saveform.cs` | Convert an HTML file to PDF while preserving CSS styles by loading HTML and using SaveFormat.Pdf. | input-bootstrap | pdf, html | verified |
| 14 | `batch-convert-all-docx-files-in-a-directory-to-html-with-round-trip-information-enabled.cs` | Batch convert all DOCX files in a directory to HTML with round‑trip information enabled. | batch-bootstrap | docx, html | verified |
| 15 | `save-a-word-document-as-html-with-embedded-fonts-exported-as-base64-using-htmlsaveoptions.cs` | Save a Word document as HTML with embedded fonts exported as Base64 using HtmlSaveOptions. | input-bootstrap | docx, html | verified |
| 16 | `enable-exportroundtripinfo-when-saving-docx-to-html-to-retain-editing-capabilities-after-r.cs` | Enable ExportRoundTripInfo when saving DOCX to HTML to retain editing capabilities after round‑trip. | input-bootstrap | docx, html | verified |
| 17 | `export-fonts-as-separate-files-while-saving-a-document-to-html-by-setting-exportfontresour.cs` | Export fonts as separate files while saving a document to HTML by setting ExportFontResources true. | input-bootstrap | html | verified |
| 18 | `convert-a-docx-file-to-mhtml-format-automatically-embedding-images-and-fonts-within-the-ou.cs` | Convert a DOCX file to MHTML format, automatically embedding images and fonts within the output. | input-bootstrap | docx, mhtml | verified |
| 19 | `create-an-email-message-attach-an-mhtml-file-and-send-it-via-smtp-server.cs` | Create an email message, attach an MHTML file, and send it via SMTP server. | input-bootstrap | mhtml | verified |
| 20 | `load-a-pdf-and-convert-each-page-to-png-images-storing-them-in-a-specified-output-folder.cs` | Load a PDF and convert each page to PNG images, storing them in a specified output folder. | batch-bootstrap | pdf, png | verified |
| 21 | `convert-a-doc-file-to-xlsx-workbook-with-default-compression-using-xlsxsaveoptions-compres.cs` | Convert a DOC file to XLSX workbook with default compression using XlsxSaveOptions.CompressionLevel. | input-bootstrap | doc | verified |
| 22 | `apply-a-find-and-replace-operation-on-a-word-document-before-exporting-it-to-excel-format.cs` | Apply a find‑and‑replace operation on a Word document before exporting it to Excel format. | save-options | docx | verified |
| 23 | `set-xlsxsaveoptions-compressionlevel-to-maximum-before-saving-a-docx-as-xlsx-to-reduce-fil.cs` | Set XlsxSaveOptions.CompressionLevel to Maximum before saving a DOCX as XLSX to reduce file size. | save-options | docx | verified |
| 24 | `stream-a-large-pdf-directly-to-a-memorystream-as-pdf-a-2b-without-loading-the-entire-file.cs` | Stream a large PDF directly to a MemoryStream as PDF/A‑2b without loading the entire file into memory. | stream-response | pdf | verified |
| 25 | `convert-multiple-image-files-png-jpeg-into-a-single-pdf-document-using-documentbuilder-and.cs` | Convert multiple image files (PNG, JPEG) into a single PDF document using DocumentBuilder and InsertImage. | batch-bootstrap | pdf, jpeg | verified |
| 26 | `load-an-epub-e-book-and-convert-it-to-pdf-while-preserving-chapter-headings-and-page-break.cs` | Load an EPUB e‑book and convert it to PDF while preserving chapter headings and page breaks. | native-conversion-api | pdf, epub | verified |
| 27 | `save-a-pdf-as-xps-format-for-compatibility-with-windows-xps-viewer-using-saveformat-xps.cs` | Save a PDF as XPS format for compatibility with Windows XPS viewer using SaveFormat.Xps. | native-conversion-api | pdf, xps | verified |
| 28 | `convert-an-xps-document-to-pdf-by-loading-xps-and-saving-with-saveformat-pdf.cs` | Convert an XPS document to PDF by loading XPS and saving with SaveFormat.Pdf. | native-conversion-api | pdf, xps | verified |
| 29 | `batch-convert-all-html-files-in-a-directory-to-mhtml-embedding-resources-automatically-for.cs` | Batch convert all HTML files in a directory to MHTML, embedding resources automatically for each file. | batch-bootstrap | html, mhtml | verified |
| 30 | `load-a-docx-from-a-network-stream-and-convert-it-to-pdf-a-2u-compliance-using-pdfsaveoptio.cs` | Load a DOCX from a network stream and convert it to PDF/A‑2u compliance using PdfSaveOptions. | stream-response | pdf, docx | verified |
| 31 | `convert-a-pdf-to-html-with-external-css-files-by-disabling-exportembeddedcss-in-htmlsaveop.cs` | Convert a PDF to HTML with external CSS files by disabling ExportEmbeddedCss in HtmlSaveOptions. | input-bootstrap | pdf, html | verified |
| 32 | `enable-exportembeddedimages-when-saving-pdf-to-html-to-embed-images-as-base64-data-uris.cs` | Enable ExportEmbeddedImages when saving PDF to HTML to embed images as Base64 data URIs. | input-bootstrap | pdf, html | verified |
| 33 | `save-a-word-document-as-html-with-css-class-prefixes-to-avoid-style-conflicts-using-htmlsa.cs` | Save a Word document as HTML with CSS class prefixes to avoid style conflicts using HtmlSaveOptions. | input-bootstrap | docx, html | verified |
| 34 | `convert-a-pdf-to-mhtml-and-then-send-it-via-email-with-custom-mime-type-using-aspose-email.cs` | Convert a PDF to MHTML and then send it via email with custom MIME type using Aspose.Email. | input-bootstrap | pdf, mhtml | verified |
| 35 | `load-a-doc-file-from-a-byte-array-and-convert-it-to-pdf-without-writing-intermediate-files.cs` | Load a DOC file from a byte array and convert it to PDF without writing intermediate files. | input-bootstrap | pdf, doc | verified |
| 36 | `batch-process-pdfs-to-generate-jpeg-thumbnails-of-the-first-page-using-jpegsaveoptions-wit.cs` | Batch process PDFs to generate JPEG thumbnails of the first page using JpegSaveOptions with low quality. | batch-bootstrap | pdf, jpeg | verified |
| 37 | `convert-a-docx-to-pdf-and-embed-a-custom-font-by-setting-fontembeddingmode-to-embedallfont.cs` | Convert a DOCX to PDF and embed a custom font by setting FontEmbeddingMode to EmbedAllFonts. | native-conversion-api | pdf, docx | verified |
| 38 | `save-a-pdf-as-pdf-a-3b-with-embedded-xml-metadata-using-pdfsaveoptions-custommetadataexpor.cs` | Save a PDF as PDF/A‑3b with embedded XML metadata using PdfSaveOptions.CustomMetadataExport for archival. | input-bootstrap | pdf, xml | verified |
| 39 | `convert-a-large-docx-to-pdf-using-streaming-to-minimize-memory-consumption-during-conversi.cs` | Convert a large DOCX to PDF using streaming to minimize memory consumption during conversion. | stream-response | pdf, docx | verified |
| 40 | `load-a-pdf-from-a-url-stream-and-convert-it-to-docx-without-saving-the-intermediate-file.cs` | Load a PDF from a URL stream and convert it to DOCX without saving the intermediate file. | stream-response | pdf, docx, doc | verified |
| 41 | `batch-convert-a-set-of-rtf-files-to-pdf-a-1a-compliance-for-legal-document-archiving.cs` | Batch convert a set of RTF files to PDF/A‑1a compliance for legal document archiving. | batch-bootstrap | pdf, rtf | verified |
| 42 | `export-a-pdf-to-png-images-with-300-dpi-resolution-by-configuring-pngsaveoptions-resolutio.cs` | Export a PDF to PNG images with 300 DPI resolution by configuring PngSaveOptions.Resolution. | save-options | pdf, png | verified |
| 43 | `convert-a-pdf-to-multiple-image-formats-jpg-and-png-in-a-single-operation-using-separate-s.cs` | Convert a PDF to multiple image formats (JPG and PNG) in a single operation using separate save calls. | batch-bootstrap | pdf, png | verified |
| 44 | `export-a-word-document-as-html-with-external-javascript-files-disabled-by-setting-exportja.cs` | Export a Word document as HTML with external JavaScript files disabled by setting ExportJavaScript false. | input-bootstrap | docx, html | verified |
| 45 | `load-an-mhtml-file-and-convert-it-back-to-pdf-while-preserving-embedded-images-and-styles.cs` | Load an MHTML file and convert it back to PDF while preserving embedded images and styles. | input-bootstrap | pdf, mhtml | verified |
| 46 | `convert-a-pdf-containing-form-fields-to-docx-while-preserving-form-data-for-further-editin.cs` | Convert a PDF containing form fields to DOCX while preserving form data for further editing. | native-conversion-api | pdf, docx, doc | verified |
| 47 | `batch-convert-html-files-to-epub-format-for-creating-a-collection-of-e-books-from-web-cont.cs` | Batch convert HTML files to EPUB format for creating a collection of e‑books from web content. | batch-bootstrap | html, epub | verified |
| 48 | `load-a-docx-replace-placeholder-text-with-actual-values-then-export-to-pdf-for-report-gene.cs` | Load a DOCX, replace placeholder text with actual values, then export to PDF for report generation. | save-options | pdf, docx | verified |
| 49 | `convert-a-pdf-to-html-with-embedded-fonts-as-base64-by-enabling-exportembeddedfonts-option.cs` | Convert a PDF to HTML with embedded fonts as Base64 by enabling ExportEmbeddedFonts option. | input-bootstrap | pdf, html | verified |
| 50 | `stream-a-doc-file-directly-to-a-network-response-as-pdf-without-intermediate-storage-on-se.cs` | Stream a DOC file directly to a network response as PDF without intermediate storage on server. | stream-response | pdf, doc | verified |
| 51 | `batch-convert-a-collection-of-png-images-into-a-single-pdf-document-with-each-image-on-sep.cs` | Batch convert a collection of PNG images into a single PDF document with each image on separate page. | batch-bootstrap | pdf, png | verified |
| 52 | `load-a-pdf-and-extract-its-text-content-to-a-plain-txt-file-using-document-gettext.cs` | Load a PDF and extract its text content to a plain TXT file using Document.GetText. | input-bootstrap | pdf, txt | verified |
| 53 | `convert-a-docx-to-mhtml-and-automatically-embed-all-linked-css-files-within-the-output.cs` | Convert a DOCX to MHTML and automatically embed all linked CSS files within the output. | input-bootstrap | docx, mhtml | verified |
| 54 | `save-a-pdf-as-an-image-sequence-in-png-format-with-transparent-background-using-pngsaveopt.cs` | Save a PDF as an image sequence in PNG format with transparent background using PngSaveOptions. | save-options | pdf, png | verified |
| 55 | `apply-a-custom-page-size-when-converting-a-doc-to-pdf-by-setting-pdfsaveoptions-pagesize.cs` | Apply a custom page size when converting a DOC to PDF by setting PdfSaveOptions.PageSize. | save-options | pdf | verified |
| 56 | `convert-a-pdf-to-an-xps-file-for-printing-on-windows-printers-that-require-xps-input.cs` | Convert a PDF to an XPS file for printing on Windows printers that require XPS input. | native-conversion-api | pdf, xps | verified |
| 57 | `batch-convert-multiple-pdfs-to-html-files-preserving-original-layout-and-fonts-using-htmls.cs` | Batch convert multiple PDFs to HTML files, preserving original layout and fonts using HtmlSaveOptions. | batch-bootstrap | pdf, html | verified |
| 58 | `load-an-epub-and-convert-it-to-mhtml-embedding-all-resources-for-offline-viewing.cs` | Load an EPUB and convert it to MHTML, embedding all resources for offline viewing. | input-bootstrap | epub, mhtml | verified |
| 59 | `save-a-docx-as-pdf-with-image-compression-level-set-to-high-to-reduce-file-size.cs` | Save a DOCX as PDF with image compression level set to High to reduce file size. | native-conversion-api | pdf, docx | verified |
| 60 | `convert-a-pdf-to-a-searchable-pdf-a-2b-by-enabling-ocr-during-the-save-operation.cs` | Convert a PDF to a searchable PDF/A‑2b by enabling OCR during the save operation. | save-options | pdf | verified |
| 61 | `batch-process-a-folder-of-doc-files-converting-each-to-pdf-and-logging-conversion-status.cs` | Batch process a folder of DOC files, converting each to PDF and logging conversion status. | batch-bootstrap | pdf, doc | verified |
| 62 | `load-a-pdf-and-convert-it-to-a-high-resolution-png-image-for-detailed-graphical-analysis.cs` | Load a PDF and convert it to a high‑resolution PNG image for detailed graphical analysis. | native-conversion-api | pdf, png | verified |
| 63 | `export-a-word-document-to-html-with-external-css-files-by-disabling-exportembeddedcss-opti.cs` | Export a Word document to HTML with external CSS files by disabling ExportEmbeddedCss option. | input-bootstrap | docx, html | verified |
| 64 | `convert-a-pdf-to-an-epub-while-preserving-chapter-hierarchy-and-navigation-metadata.cs` | Convert a PDF to an EPUB while preserving chapter hierarchy and navigation metadata. | native-conversion-api | pdf, epub | verified |
| 65 | `save-a-pdf-as-a-pdf-a-1b-compliant-file-with-embedded-icc-color-profile-using-pdfsaveoptio.cs` | Save a PDF as a PDF/A‑1b compliant file with embedded ICC color profile using PdfSaveOptions. | save-options | pdf | verified |
| 66 | `load-a-docx-from-a-cloud-storage-stream-and-convert-it-to-pdf-for-client-download.cs` | Load a DOCX from a cloud storage stream and convert it to PDF for client download. | stream-response | pdf, docx | verified |
| 67 | `batch-convert-html-files-to-pdf-with-custom-page-margins-defined-in-pdfsaveoptions.cs` | Batch convert HTML files to PDF with custom page margins defined in PdfSaveOptions. | batch-bootstrap | pdf, html | verified |
| 68 | `convert-a-pdf-containing-vector-graphics-to-high-resolution-png-to-retain-visual-fidelity.cs` | Convert a PDF containing vector graphics to high‑resolution PNG to retain visual fidelity. | native-conversion-api | pdf, png | verified |
| 69 | `add-a-header-footer-to-a-docx-before-converting-to-pdf-using-documentbuilder.cs` | Add a header footer to a DOCX before converting to PDF using DocumentBuilder. | native-conversion-api | pdf, docx | verified |
| 70 | `convert-a-pdf-to-a-docx-and-then-to-xlsx-to-extract-tables-into-spreadsheet-format.cs` | Convert a PDF to a DOCX and then to XLSX to extract tables into spreadsheet format. | native-conversion-api | pdf, docx | verified |
| 71 | `apply-compression-to-an-xlsx-file-generated-from-a-docx-by-setting-xlsxsaveoptions-compres.cs` | Apply compression to an XLSX file generated from a DOCX by setting XlsxSaveOptions.CompressionLevel to Fast. | save-options | docx | verified |
| 72 | `load-a-docx-replace-all-occurrences-of-a-placeholder-with-actual-data-then-save-as-pdf.cs` | Load a DOCX, replace all occurrences of a placeholder with actual data, then save as PDF. | native-conversion-api | pdf, docx | verified |
| 73 | `convert-a-pdf-to-html-with-embedded-fonts-as-base64-and-images-as-external-files-for-web-u.cs` | Convert a PDF to HTML with embedded fonts as Base64 and images as external files for web use. | input-bootstrap | pdf, html | verified |
| 74 | `save-a-pdf-as-a-pdf-a-3u-document-and-embed-an-associated-xml-file-for-enhanced-metadata.cs` | Save a PDF as a PDF/A‑3u document and embed an associated XML file for enhanced metadata. | input-bootstrap | pdf, xml | verified |
| 75 | `batch-convert-a-collection-of-html-files-to-mhtml-ensuring-all-linked-resources-are-embedd.cs` | Batch convert a collection of HTML files to MHTML, ensuring all linked resources are embedded automatically. | batch-bootstrap | html, mhtml | verified |
| 76 | `load-a-pdf-and-convert-it-to-a-docx-while-preserving-original-layout-using-layout-preserva.cs` | Load a PDF and convert it to a DOCX while preserving original layout using layout preservation options. | save-options | pdf, docx | verified |
| 77 | `convert-a-docx-to-pdf-and-embed-a-custom-cover-page-image-using-documentbuilder-insertion.cs` | Convert a DOCX to PDF and embed a custom cover page image using DocumentBuilder insertion. | native-conversion-api | pdf, docx | verified |
| 78 | `batch-convert-multiple-pdfs-to-high-resolution-png-images-with-600-dpi-for-print-ready-out.cs` | Batch convert multiple PDFs to high‑resolution PNG images with 600 DPI for print‑ready output. | batch-bootstrap | pdf, png | verified |
| 79 | `save-a-pdf-as-a-compressed-pdf-a-2u-file-by-enabling-image-compression-and-removing-unused.cs` | Save a PDF as a compressed PDF/A‑2u file by enabling image compression and removing unused objects. | save-options | pdf | verified |
| 80 | `load-an-html-file-with-external-javascript-and-convert-it-to-pdf-ignoring-scripts-for-stat.cs` | Load an HTML file with external JavaScript and convert it to PDF, ignoring scripts for static rendering. | input-bootstrap | pdf, html | verified |
| 81 | `convert-a-pdf-to-an-xlsx-spreadsheet-extracting-any-tabular-data-into-separate-worksheets.cs` | Convert a PDF to an XLSX spreadsheet, extracting any tabular data into separate worksheets. | native-conversion-api | pdf | verified |
| 82 | `batch-process-docx-files-applying-a-company-wide-header-before-converting-each-to-pdf-for.cs` | Batch process DOCX files, applying a company‑wide header before converting each to PDF for distribution. | batch-bootstrap | pdf, docx | verified |
| 83 | `save-a-pdf-as-an-image-sequence-in-png-format-with-lossless-compression-for-archival-purpo.cs` | Save a PDF as an image sequence in PNG format with lossless compression for archival purposes. | native-conversion-api | pdf, png | verified |
| 84 | `convert-a-pdf-to-html-with-external-css-files-and-separate-image-folder-for-optimized-web.cs` | Convert a PDF to HTML with external CSS files and separate image folder for optimized web delivery. | batch-bootstrap | pdf, html | verified |
| 85 | `load-a-docx-from-a-sharepoint-stream-and-convert-it-to-pdf-for-on-demand-download.cs` | Load a DOCX from a SharePoint stream and convert it to PDF for on‑demand download. | stream-response | pdf, docx | verified |
| 86 | `batch-convert-a-set-of-pdf-files-to-epub-preserving-original-chapter-structure-for-e-readi.cs` | Batch convert a set of PDF files to EPUB, preserving original chapter structure for e‑reading devices. | batch-bootstrap | pdf, epub | verified |
| 87 | `apply-a-custom-pdf-a-2b-compliance-level-when-converting-doc-to-pdf-using-pdfsaveoptions.cs` | Apply a custom PDF/A‑2b compliance level when converting DOC to PDF using PdfSaveOptions. | save-options | pdf | verified |
| 88 | `convert-a-pdf-to-a-searchable-pdf-a-1a-by-enabling-ocr-and-setting-appropriate-compliance.cs` | Convert a PDF to a searchable PDF/A‑1a by enabling OCR and setting appropriate compliance. | save-options | pdf | verified |
| 89 | `save-a-docx-as-html-with-external-javascript-disabled-to-improve-security-of-the-generated.cs` | Save a DOCX as HTML with external JavaScript disabled to improve security of the generated page. | input-bootstrap | docx, html | verified |
| 90 | `batch-convert-multiple-png-images-to-a-single-pdf-arranging-each-image-on-a-separate-page.cs` | Batch convert multiple PNG images to a single PDF, arranging each image on a separate page. | batch-bootstrap | pdf, png | verified |
| 91 | `load-a-pdf-and-convert-it-to-an-xps-document-while-preserving-annotations-for-review.cs` | Load a PDF and convert it to an XPS document while preserving annotations for review. | native-conversion-api | pdf, xps | verified |
| 92 | `convert-a-docx-to-mhtml-and-then-embed-it-in-an-email-body-using-aspose-email-html-body-su.cs` | Convert a DOCX to MHTML and then embed it in an email body using Aspose.Email HTML body support. | input-bootstrap | docx, html, mhtml | verified |
| 93 | `save-a-pdf-as-a-pdf-a-3u-file-and-attach-an-associated-xml-metadata-file-for-compliance.cs` | Save a PDF as a PDF/A‑3u file and attach an associated XML metadata file for compliance. | input-bootstrap | pdf, xml | verified |
| 94 | `batch-process-html-files-converting-each-to-pdf-with-custom-page-size-defined-in-pdfsaveop.cs` | Batch process HTML files, converting each to PDF with custom page size defined in PdfSaveOptions. | batch-bootstrap | pdf, html | verified |
| 95 | `convert-a-pdf-to-a-high-resolution-jpeg-image-setting-jpegsaveoptions-compressionlevel-to.cs` | Convert a PDF to a high‑resolution JPEG image, setting JpegSaveOptions.CompressionLevel to Low for quality. | save-options | pdf, jpeg | verified |
| 96 | `load-a-docx-replace-all-placeholder-dates-with-current-date-and-export-to-pdf-for-reportin.cs` | Load a DOCX, replace all placeholder dates with current date, and export to PDF for reporting. | save-options | pdf, docx | verified |
| 97 | `load-a-docx-file-and-save-it-as-markdown-extracting-images-to-a-custom-folder.cs` | Load a DOCX file and save it as Markdown, extracting images to a custom folder. | batch-bootstrap | docx, md | verified |
| 98 | `convert-a-pdf-document-to-docx-format-while-preserving-text-formatting-and-hyperlinks.cs` | Convert a PDF document to DOCX format while preserving text formatting and hyperlinks. | native-conversion-api | pdf, docx, doc | verified |
| 99 | `load-a-pdf-with-pdfloadoptions-to-ignore-password-protection-and-save-it-as-a-docx-file.cs` | Load a PDF with PdfLoadOptions to ignore password protection and save it as a DOCX file. | input-bootstrap | pdf, docx | verified |
| 100 | `load-a-pdf-with-password-protection-provide-password-via-pdfloadoptions-and-convert-to-png.cs` | Load a PDF with password protection, provide password via PdfLoadOptions, and convert to PNG without errors. | save-options | pdf, png | verified |
| 101 | `load-multiple-pdf-files-from-a-list-convert-each-to-docx-and-write-results-to-an-output-fo.cs` | Load multiple PDF files from a list, convert each to DOCX, and write results to an output folder. | batch-bootstrap | pdf, docx, doc | verified |
| 102 | `load-a-pdf-set-pdfloadoptions-to-ignore-errors-and-save-as-docx-while-preserving-page-brea.cs` | Load a PDF, set PdfLoadOptions to ignore errors, and save as DOCX while preserving page breaks. | save-options | pdf, docx, doc | verified |
| 103 | `export-a-pdf-to-a-single-high-quality-jpeg-image-using-multipagelayout-horizontal-and-cust.cs` | Export a PDF to a single high‑quality JPEG image using MultiPageLayout.Horizontal and custom JpegQuality. | save-options | pdf, jpeg | verified |
| 104 | `export-a-pdf-as-a-series-of-png-files-using-pageset-to-export-only-even-numbered-pages.cs` | Export a PDF as a series of PNG files, using PageSet to export only even‑numbered pages. | save-options | pdf, png | verified |
| 105 | `export-selected-pages-1-4-7-of-a-pdf-as-separate-png-images-with-custom-resolution-setting.cs` | Export selected pages 1,4,7 of a PDF as separate PNG images with custom resolution settings. | save-options | pdf, png | verified |
| 106 | `save-a-pdf-as-markdown-extracting-images-to-a-temporary-folder-and-embedding-relative-path.cs` | Save a PDF as Markdown, extracting images to a temporary folder and embedding relative paths in the output. | batch-bootstrap | pdf, md | verified |
| 107 | `convert-a-pdf-to-markdown-ensuring-that-all-extracted-images-are-saved-in-a-subfolder-name.cs` | Convert a PDF to Markdown, ensuring that all extracted images are saved in a subfolder named "assets". | batch-bootstrap | pdf, md | verified |
| 108 | `convert-a-pdf-to-tiff-using-lzw-compression-and-setting-imagecontrast-to-10-for-improved-c.cs` | Convert a PDF to TIFF, using LZW compression and setting ImageContrast to 10 for improved clarity. | native-conversion-api | pdf, tiff | verified |

## Common failure patterns seen during generation and how they were corrected

### System.Drawing or Font ambiguity

- Symptom: Compile errors caused by System.Drawing usage or ambiguous Font references between Aspose.Drawing.Font and Aspose.Words.Font.
- Fix: Use Aspose.Drawing only, and declare drawing fonts explicitly as Aspose.Drawing.Font.

### Missing input bootstrap

- Symptom: Conversion task assumes a DOC, DOCX, HTML, TXT, XML, folder, or other source already exists.
- Fix: Create the local source file, source document, or input folder inside the example before loading it.

### Response-stream task hangs

- Symptom: Example starts a server or waits for a request instead of finishing.
- Fix: Simulate the response with a MemoryStream, write conversion output to it, validate length, and exit.

### Incorrect SaveFormat or extension pairing

- Symptom: Output format does not match the requested target type or file extension.
- Fix: Use the correct SaveFormat or specialized SaveOptions for the requested output.

### Stream position or disposal issue

- Symptom: Reloading or reading a stream fails because Position was not reset or the stream was disposed too early.
- Fix: Reset MemoryStream.Position before reading and keep the stream alive until validation is complete.

## Notes for maintainers

- This category is now **publish-ready** for the current run.
- Preserve file-to-task traceability when updating this folder.
- For future updates, keep the examples standalone and continue bootstrapping local inputs inside the example whenever external sources are mentioned.
