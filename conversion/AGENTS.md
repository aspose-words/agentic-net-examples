---
name: conversion
description: Verified C# examples for conversion scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Conversion

## Purpose

This folder is a **live, curated example set** for conversion scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free document format conversion workflows using direct Aspose.Words APIs.

## Non-negotiable conventions

- Use native Aspose.Words conversion APIs directly.
- Bootstrap local input files or documents whenever the task implies an existing source file, folder, stream, or source format.
- Use `Aspose.Drawing`, not `System.Drawing`, when drawing-related types are needed.
- For drawing fonts, prefer explicit declarations such as `Aspose.Drawing.Font font = new Aspose.Drawing.Font(...)`.
- For response-stream tasks, simulate the response with a `MemoryStream`; do not start a real server or listener.
- Guard maybe-null values to avoid nullable-reference warnings such as `CS8600`, `CS8602`, and `CS8604`.

## Recommended workflow selection

- **Native conversion API workflow**: 22 examples
- **Input-bootstrap workflow**: 30 examples
- **Save-options workflow**: 24 examples
- **Stream / response simulation workflow**: 8 examples
- **Batch / folder bootstrap workflow**: 24 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. Source documents or source files must be bootstrapped locally whenever the task implies an existing input.
3. Output files or output streams must actually be produced and validated.
4. Response-stream tasks must finish without waiting for external requests.
5. Drawing-related types must use `Aspose.Drawing` and not `System.Drawing`.

## File-to-task reference

- `load-a-docx-file-from-disk-and-save-it-as-pdf-preserving-layout.cs`
  - Task: Load a DOCX file from disk and save it as PDF preserving layout.
  - Workflow: input-bootstrap
  - Outputs: pdf, docx
  - Selected engine: verified
- `convert-a-doc-file-to-pdf-a-1b-by-setting-pdfsaveoptions-compliance-before-saving.cs`
  - Task: Convert a DOC file to PDF/A‑1b by setting PdfSaveOptions.Compliance before saving.
  - Workflow: input-bootstrap
  - Outputs: pdf, doc
  - Selected engine: verified
- `stream-a-docx-document-into-memory-and-write-the-pdf-output-directly-to-response-stream.cs`
  - Task: Stream a DOCX document into memory and write the PDF output directly to response stream.
  - Workflow: stream-response
  - Outputs: pdf, docx
  - Selected engine: verified
- `batch-process-all-rtf-files-in-a-folder-converting-each-to-pdf-with-default-layout.cs`
  - Task: Batch process all RTF files in a folder, converting each to PDF with default layout.
  - Workflow: batch-bootstrap
  - Outputs: pdf, rtf
  - Selected engine: verified
- `load-an-html-page-from-a-url-and-convert-it-to-pdf-using-document-load-and-saveformat-pdf.cs`
  - Task: Load an HTML page from a URL and convert it to PDF using Document.Load and SaveFormat.Pdf.
  - Workflow: input-bootstrap
  - Outputs: pdf, html
  - Selected engine: verified
- `convert-a-png-image-to-pdf-by-loading-the-image-file-and-saving-with-saveformat-pdf.cs`
  - Task: Convert a PNG image to PDF by loading the image file and saving with SaveFormat.Pdf.
  - Workflow: native-conversion-api
  - Outputs: pdf, png
  - Selected engine: verified
- `transform-a-jpeg-picture-into-a-pdf-document-while-preserving-original-resolution-and-colo.cs`
  - Task: Transform a JPEG picture into a PDF document while preserving original resolution and color depth.
  - Workflow: native-conversion-api
  - Outputs: pdf, jpeg
  - Selected engine: verified
- `load-a-pdf-file-and-export-its-pages-as-separate-png-images-using-saveformat-png.cs`
  - Task: Load a PDF file and export its pages as separate PNG images using SaveFormat.Png.
  - Workflow: save-options
  - Outputs: pdf, png
  - Selected engine: verified
- `convert-a-pdf-document-to-high-quality-jpg-images-by-setting-jpegsaveoptions-quality-befor.cs`
  - Task: Convert a PDF document to high‑quality JPG images by setting JpegSaveOptions.Quality before saving.
  - Workflow: save-options
  - Outputs: pdf, jpeg
  - Selected engine: verified
- `save-a-pdf-as-pdf-ua-compliant-file-by-configuring-pdfsaveoptions-compliance-to-pdfcomplia.cs`
  - Task: Save a PDF as PDF/UA compliant file by configuring PdfSaveOptions.Compliance to PdfCompliance.Ua.
  - Workflow: save-options
  - Outputs: pdf
  - Selected engine: verified
- `load-a-pdf-and-export-it-to-docx-format-for-further-editing-using-saveformat-docx.cs`
  - Task: Load a PDF and export it to DOCX format for further editing using SaveFormat.Docx.
  - Workflow: save-options
  - Outputs: pdf, docx, doc
  - Selected engine: verified
- `transform-a-pdf-into-an-epub-e-book-by-loading-the-pdf-and-saving-with-saveformat-epub.cs`
  - Task: Transform a PDF into an EPUB e‑book by loading the PDF and saving with SaveFormat.Epub.
  - Workflow: native-conversion-api
  - Outputs: pdf, epub
  - Selected engine: verified
- `convert-an-html-file-to-pdf-while-preserving-css-styles-by-loading-html-and-using-saveform.cs`
  - Task: Convert an HTML file to PDF while preserving CSS styles by loading HTML and using SaveFormat.Pdf.
  - Workflow: input-bootstrap
  - Outputs: pdf, html
  - Selected engine: verified
- `batch-convert-all-docx-files-in-a-directory-to-html-with-round-trip-information-enabled.cs`
  - Task: Batch convert all DOCX files in a directory to HTML with round‑trip information enabled.
  - Workflow: batch-bootstrap
  - Outputs: docx, html
  - Selected engine: verified
- `save-a-word-document-as-html-with-embedded-fonts-exported-as-base64-using-htmlsaveoptions.cs`
  - Task: Save a Word document as HTML with embedded fonts exported as Base64 using HtmlSaveOptions.
  - Workflow: input-bootstrap
  - Outputs: docx, html
  - Selected engine: verified
- `enable-exportroundtripinfo-when-saving-docx-to-html-to-retain-editing-capabilities-after-r.cs`
  - Task: Enable ExportRoundTripInfo when saving DOCX to HTML to retain editing capabilities after round‑trip.
  - Workflow: input-bootstrap
  - Outputs: docx, html
  - Selected engine: verified
- `export-fonts-as-separate-files-while-saving-a-document-to-html-by-setting-exportfontresour.cs`
  - Task: Export fonts as separate files while saving a document to HTML by setting ExportFontResources true.
  - Workflow: input-bootstrap
  - Outputs: html
  - Selected engine: verified
- `convert-a-docx-file-to-mhtml-format-automatically-embedding-images-and-fonts-within-the-ou.cs`
  - Task: Convert a DOCX file to MHTML format, automatically embedding images and fonts within the output.
  - Workflow: input-bootstrap
  - Outputs: docx, mhtml
  - Selected engine: verified
- `create-an-email-message-attach-an-mhtml-file-and-send-it-via-smtp-server.cs`
  - Task: Create an email message, attach an MHTML file, and send it via SMTP server.
  - Workflow: input-bootstrap
  - Outputs: mhtml
  - Selected engine: verified
- `load-a-pdf-and-convert-each-page-to-png-images-storing-them-in-a-specified-output-folder.cs`
  - Task: Load a PDF and convert each page to PNG images, storing them in a specified output folder.
  - Workflow: batch-bootstrap
  - Outputs: pdf, png
  - Selected engine: verified
- `convert-a-doc-file-to-xlsx-workbook-with-default-compression-using-xlsxsaveoptions-compres.cs`
  - Task: Convert a DOC file to XLSX workbook with default compression using XlsxSaveOptions.CompressionLevel.
  - Workflow: input-bootstrap
  - Outputs: doc
  - Selected engine: verified
- `apply-a-find-and-replace-operation-on-a-word-document-before-exporting-it-to-excel-format.cs`
  - Task: Apply a find‑and‑replace operation on a Word document before exporting it to Excel format.
  - Workflow: save-options
  - Outputs: docx
  - Selected engine: verified
- `set-xlsxsaveoptions-compressionlevel-to-maximum-before-saving-a-docx-as-xlsx-to-reduce-fil.cs`
  - Task: Set XlsxSaveOptions.CompressionLevel to Maximum before saving a DOCX as XLSX to reduce file size.
  - Workflow: save-options
  - Outputs: docx
  - Selected engine: verified
- `stream-a-large-pdf-directly-to-a-memorystream-as-pdf-a-2b-without-loading-the-entire-file.cs`
  - Task: Stream a large PDF directly to a MemoryStream as PDF/A‑2b without loading the entire file into memory.
  - Workflow: stream-response
  - Outputs: pdf
  - Selected engine: verified
- `convert-multiple-image-files-png-jpeg-into-a-single-pdf-document-using-documentbuilder-and.cs`
  - Task: Convert multiple image files (PNG, JPEG) into a single PDF document using DocumentBuilder and InsertImage.
  - Workflow: batch-bootstrap
  - Outputs: pdf, jpeg
  - Selected engine: verified
- `load-an-epub-e-book-and-convert-it-to-pdf-while-preserving-chapter-headings-and-page-break.cs`
  - Task: Load an EPUB e‑book and convert it to PDF while preserving chapter headings and page breaks.
  - Workflow: native-conversion-api
  - Outputs: pdf, epub
  - Selected engine: verified
- `save-a-pdf-as-xps-format-for-compatibility-with-windows-xps-viewer-using-saveformat-xps.cs`
  - Task: Save a PDF as XPS format for compatibility with Windows XPS viewer using SaveFormat.Xps.
  - Workflow: native-conversion-api
  - Outputs: pdf, xps
  - Selected engine: verified
- `convert-an-xps-document-to-pdf-by-loading-xps-and-saving-with-saveformat-pdf.cs`
  - Task: Convert an XPS document to PDF by loading XPS and saving with SaveFormat.Pdf.
  - Workflow: native-conversion-api
  - Outputs: pdf, xps
  - Selected engine: verified
- `batch-convert-all-html-files-in-a-directory-to-mhtml-embedding-resources-automatically-for.cs`
  - Task: Batch convert all HTML files in a directory to MHTML, embedding resources automatically for each file.
  - Workflow: batch-bootstrap
  - Outputs: html, mhtml
  - Selected engine: verified
- `load-a-docx-from-a-network-stream-and-convert-it-to-pdf-a-2u-compliance-using-pdfsaveoptio.cs`
  - Task: Load a DOCX from a network stream and convert it to PDF/A‑2u compliance using PdfSaveOptions.
  - Workflow: stream-response
  - Outputs: pdf, docx
  - Selected engine: verified
- `convert-a-pdf-to-html-with-external-css-files-by-disabling-exportembeddedcss-in-htmlsaveop.cs`
  - Task: Convert a PDF to HTML with external CSS files by disabling ExportEmbeddedCss in HtmlSaveOptions.
  - Workflow: input-bootstrap
  - Outputs: pdf, html
  - Selected engine: verified
- `enable-exportembeddedimages-when-saving-pdf-to-html-to-embed-images-as-base64-data-uris.cs`
  - Task: Enable ExportEmbeddedImages when saving PDF to HTML to embed images as Base64 data URIs.
  - Workflow: input-bootstrap
  - Outputs: pdf, html
  - Selected engine: verified
- `save-a-word-document-as-html-with-css-class-prefixes-to-avoid-style-conflicts-using-htmlsa.cs`
  - Task: Save a Word document as HTML with CSS class prefixes to avoid style conflicts using HtmlSaveOptions.
  - Workflow: input-bootstrap
  - Outputs: docx, html
  - Selected engine: verified
- `convert-a-pdf-to-mhtml-and-then-send-it-via-email-with-custom-mime-type-using-aspose-email.cs`
  - Task: Convert a PDF to MHTML and then send it via email with custom MIME type using Aspose.Email.
  - Workflow: input-bootstrap
  - Outputs: pdf, mhtml
  - Selected engine: verified
- `load-a-doc-file-from-a-byte-array-and-convert-it-to-pdf-without-writing-intermediate-files.cs`
  - Task: Load a DOC file from a byte array and convert it to PDF without writing intermediate files.
  - Workflow: input-bootstrap
  - Outputs: pdf, doc
  - Selected engine: verified
- `batch-process-pdfs-to-generate-jpeg-thumbnails-of-the-first-page-using-jpegsaveoptions-wit.cs`
  - Task: Batch process PDFs to generate JPEG thumbnails of the first page using JpegSaveOptions with low quality.
  - Workflow: batch-bootstrap
  - Outputs: pdf, jpeg
  - Selected engine: verified
- `convert-a-docx-to-pdf-and-embed-a-custom-font-by-setting-fontembeddingmode-to-embedallfont.cs`
  - Task: Convert a DOCX to PDF and embed a custom font by setting FontEmbeddingMode to EmbedAllFonts.
  - Workflow: native-conversion-api
  - Outputs: pdf, docx
  - Selected engine: verified
- `save-a-pdf-as-pdf-a-3b-with-embedded-xml-metadata-using-pdfsaveoptions-custommetadataexpor.cs`
  - Task: Save a PDF as PDF/A‑3b with embedded XML metadata using PdfSaveOptions.CustomMetadataExport for archival.
  - Workflow: input-bootstrap
  - Outputs: pdf, xml
  - Selected engine: verified
- `convert-a-large-docx-to-pdf-using-streaming-to-minimize-memory-consumption-during-conversi.cs`
  - Task: Convert a large DOCX to PDF using streaming to minimize memory consumption during conversion.
  - Workflow: stream-response
  - Outputs: pdf, docx
  - Selected engine: verified
- `load-a-pdf-from-a-url-stream-and-convert-it-to-docx-without-saving-the-intermediate-file.cs`
  - Task: Load a PDF from a URL stream and convert it to DOCX without saving the intermediate file.
  - Workflow: stream-response
  - Outputs: pdf, docx, doc
  - Selected engine: verified
- `batch-convert-a-set-of-rtf-files-to-pdf-a-1a-compliance-for-legal-document-archiving.cs`
  - Task: Batch convert a set of RTF files to PDF/A‑1a compliance for legal document archiving.
  - Workflow: batch-bootstrap
  - Outputs: pdf, rtf
  - Selected engine: verified
- `export-a-pdf-to-png-images-with-300-dpi-resolution-by-configuring-pngsaveoptions-resolutio.cs`
  - Task: Export a PDF to PNG images with 300 DPI resolution by configuring PngSaveOptions.Resolution.
  - Workflow: save-options
  - Outputs: pdf, png
  - Selected engine: verified
- `convert-a-pdf-to-multiple-image-formats-jpg-and-png-in-a-single-operation-using-separate-s.cs`
  - Task: Convert a PDF to multiple image formats (JPG and PNG) in a single operation using separate save calls.
  - Workflow: batch-bootstrap
  - Outputs: pdf, png
  - Selected engine: verified
- `export-a-word-document-as-html-with-external-javascript-files-disabled-by-setting-exportja.cs`
  - Task: Export a Word document as HTML with external JavaScript files disabled by setting ExportJavaScript false.
  - Workflow: input-bootstrap
  - Outputs: docx, html
  - Selected engine: verified
- `load-an-mhtml-file-and-convert-it-back-to-pdf-while-preserving-embedded-images-and-styles.cs`
  - Task: Load an MHTML file and convert it back to PDF while preserving embedded images and styles.
  - Workflow: input-bootstrap
  - Outputs: pdf, mhtml
  - Selected engine: verified
- `convert-a-pdf-containing-form-fields-to-docx-while-preserving-form-data-for-further-editin.cs`
  - Task: Convert a PDF containing form fields to DOCX while preserving form data for further editing.
  - Workflow: native-conversion-api
  - Outputs: pdf, docx, doc
  - Selected engine: verified
- `batch-convert-html-files-to-epub-format-for-creating-a-collection-of-e-books-from-web-cont.cs`
  - Task: Batch convert HTML files to EPUB format for creating a collection of e‑books from web content.
  - Workflow: batch-bootstrap
  - Outputs: html, epub
  - Selected engine: verified
- `load-a-docx-replace-placeholder-text-with-actual-values-then-export-to-pdf-for-report-gene.cs`
  - Task: Load a DOCX, replace placeholder text with actual values, then export to PDF for report generation.
  - Workflow: save-options
  - Outputs: pdf, docx
  - Selected engine: verified
- `convert-a-pdf-to-html-with-embedded-fonts-as-base64-by-enabling-exportembeddedfonts-option.cs`
  - Task: Convert a PDF to HTML with embedded fonts as Base64 by enabling ExportEmbeddedFonts option.
  - Workflow: input-bootstrap
  - Outputs: pdf, html
  - Selected engine: verified
- `stream-a-doc-file-directly-to-a-network-response-as-pdf-without-intermediate-storage-on-se.cs`
  - Task: Stream a DOC file directly to a network response as PDF without intermediate storage on server.
  - Workflow: stream-response
  - Outputs: pdf, doc
  - Selected engine: verified
- `batch-convert-a-collection-of-png-images-into-a-single-pdf-document-with-each-image-on-sep.cs`
  - Task: Batch convert a collection of PNG images into a single PDF document with each image on separate page.
  - Workflow: batch-bootstrap
  - Outputs: pdf, png
  - Selected engine: verified
- `load-a-pdf-and-extract-its-text-content-to-a-plain-txt-file-using-document-gettext.cs`
  - Task: Load a PDF and extract its text content to a plain TXT file using Document.GetText.
  - Workflow: input-bootstrap
  - Outputs: pdf, txt
  - Selected engine: verified
- `convert-a-docx-to-mhtml-and-automatically-embed-all-linked-css-files-within-the-output.cs`
  - Task: Convert a DOCX to MHTML and automatically embed all linked CSS files within the output.
  - Workflow: input-bootstrap
  - Outputs: docx, mhtml
  - Selected engine: verified
- `save-a-pdf-as-an-image-sequence-in-png-format-with-transparent-background-using-pngsaveopt.cs`
  - Task: Save a PDF as an image sequence in PNG format with transparent background using PngSaveOptions.
  - Workflow: save-options
  - Outputs: pdf, png
  - Selected engine: verified
- `apply-a-custom-page-size-when-converting-a-doc-to-pdf-by-setting-pdfsaveoptions-pagesize.cs`
  - Task: Apply a custom page size when converting a DOC to PDF by setting PdfSaveOptions.PageSize.
  - Workflow: save-options
  - Outputs: pdf
  - Selected engine: verified
- `convert-a-pdf-to-an-xps-file-for-printing-on-windows-printers-that-require-xps-input.cs`
  - Task: Convert a PDF to an XPS file for printing on Windows printers that require XPS input.
  - Workflow: native-conversion-api
  - Outputs: pdf, xps
  - Selected engine: verified
- `batch-convert-multiple-pdfs-to-html-files-preserving-original-layout-and-fonts-using-htmls.cs`
  - Task: Batch convert multiple PDFs to HTML files, preserving original layout and fonts using HtmlSaveOptions.
  - Workflow: batch-bootstrap
  - Outputs: pdf, html
  - Selected engine: verified
- `load-an-epub-and-convert-it-to-mhtml-embedding-all-resources-for-offline-viewing.cs`
  - Task: Load an EPUB and convert it to MHTML, embedding all resources for offline viewing.
  - Workflow: input-bootstrap
  - Outputs: epub, mhtml
  - Selected engine: verified
- `save-a-docx-as-pdf-with-image-compression-level-set-to-high-to-reduce-file-size.cs`
  - Task: Save a DOCX as PDF with image compression level set to High to reduce file size.
  - Workflow: native-conversion-api
  - Outputs: pdf, docx
  - Selected engine: verified
- `convert-a-pdf-to-a-searchable-pdf-a-2b-by-enabling-ocr-during-the-save-operation.cs`
  - Task: Convert a PDF to a searchable PDF/A‑2b by enabling OCR during the save operation.
  - Workflow: save-options
  - Outputs: pdf
  - Selected engine: verified
- `batch-process-a-folder-of-doc-files-converting-each-to-pdf-and-logging-conversion-status.cs`
  - Task: Batch process a folder of DOC files, converting each to PDF and logging conversion status.
  - Workflow: batch-bootstrap
  - Outputs: pdf, doc
  - Selected engine: verified
- `load-a-pdf-and-convert-it-to-a-high-resolution-png-image-for-detailed-graphical-analysis.cs`
  - Task: Load a PDF and convert it to a high‑resolution PNG image for detailed graphical analysis.
  - Workflow: native-conversion-api
  - Outputs: pdf, png
  - Selected engine: verified
- `export-a-word-document-to-html-with-external-css-files-by-disabling-exportembeddedcss-opti.cs`
  - Task: Export a Word document to HTML with external CSS files by disabling ExportEmbeddedCss option.
  - Workflow: input-bootstrap
  - Outputs: docx, html
  - Selected engine: verified
- `convert-a-pdf-to-an-epub-while-preserving-chapter-hierarchy-and-navigation-metadata.cs`
  - Task: Convert a PDF to an EPUB while preserving chapter hierarchy and navigation metadata.
  - Workflow: native-conversion-api
  - Outputs: pdf, epub
  - Selected engine: verified
- `save-a-pdf-as-a-pdf-a-1b-compliant-file-with-embedded-icc-color-profile-using-pdfsaveoptio.cs`
  - Task: Save a PDF as a PDF/A‑1b compliant file with embedded ICC color profile using PdfSaveOptions.
  - Workflow: save-options
  - Outputs: pdf
  - Selected engine: verified
- `load-a-docx-from-a-cloud-storage-stream-and-convert-it-to-pdf-for-client-download.cs`
  - Task: Load a DOCX from a cloud storage stream and convert it to PDF for client download.
  - Workflow: stream-response
  - Outputs: pdf, docx
  - Selected engine: verified
- `batch-convert-html-files-to-pdf-with-custom-page-margins-defined-in-pdfsaveoptions.cs`
  - Task: Batch convert HTML files to PDF with custom page margins defined in PdfSaveOptions.
  - Workflow: batch-bootstrap
  - Outputs: pdf, html
  - Selected engine: verified
- `convert-a-pdf-containing-vector-graphics-to-high-resolution-png-to-retain-visual-fidelity.cs`
  - Task: Convert a PDF containing vector graphics to high‑resolution PNG to retain visual fidelity.
  - Workflow: native-conversion-api
  - Outputs: pdf, png
  - Selected engine: verified
- `add-a-header-footer-to-a-docx-before-converting-to-pdf-using-documentbuilder.cs`
  - Task: Add a header footer to a DOCX before converting to PDF using DocumentBuilder.
  - Workflow: native-conversion-api
  - Outputs: pdf, docx
  - Selected engine: verified
- `convert-a-pdf-to-a-docx-and-then-to-xlsx-to-extract-tables-into-spreadsheet-format.cs`
  - Task: Convert a PDF to a DOCX and then to XLSX to extract tables into spreadsheet format.
  - Workflow: native-conversion-api
  - Outputs: pdf, docx
  - Selected engine: verified
- `apply-compression-to-an-xlsx-file-generated-from-a-docx-by-setting-xlsxsaveoptions-compres.cs`
  - Task: Apply compression to an XLSX file generated from a DOCX by setting XlsxSaveOptions.CompressionLevel to Fast.
  - Workflow: save-options
  - Outputs: docx
  - Selected engine: verified
- `load-a-docx-replace-all-occurrences-of-a-placeholder-with-actual-data-then-save-as-pdf.cs`
  - Task: Load a DOCX, replace all occurrences of a placeholder with actual data, then save as PDF.
  - Workflow: native-conversion-api
  - Outputs: pdf, docx
  - Selected engine: verified
- `convert-a-pdf-to-html-with-embedded-fonts-as-base64-and-images-as-external-files-for-web-u.cs`
  - Task: Convert a PDF to HTML with embedded fonts as Base64 and images as external files for web use.
  - Workflow: input-bootstrap
  - Outputs: pdf, html
  - Selected engine: verified
- `save-a-pdf-as-a-pdf-a-3u-document-and-embed-an-associated-xml-file-for-enhanced-metadata.cs`
  - Task: Save a PDF as a PDF/A‑3u document and embed an associated XML file for enhanced metadata.
  - Workflow: input-bootstrap
  - Outputs: pdf, xml
  - Selected engine: verified
- `batch-convert-a-collection-of-html-files-to-mhtml-ensuring-all-linked-resources-are-embedd.cs`
  - Task: Batch convert a collection of HTML files to MHTML, ensuring all linked resources are embedded automatically.
  - Workflow: batch-bootstrap
  - Outputs: html, mhtml
  - Selected engine: verified
- `load-a-pdf-and-convert-it-to-a-docx-while-preserving-original-layout-using-layout-preserva.cs`
  - Task: Load a PDF and convert it to a DOCX while preserving original layout using layout preservation options.
  - Workflow: save-options
  - Outputs: pdf, docx
  - Selected engine: verified
- `convert-a-docx-to-pdf-and-embed-a-custom-cover-page-image-using-documentbuilder-insertion.cs`
  - Task: Convert a DOCX to PDF and embed a custom cover page image using DocumentBuilder insertion.
  - Workflow: native-conversion-api
  - Outputs: pdf, docx
  - Selected engine: verified
- `batch-convert-multiple-pdfs-to-high-resolution-png-images-with-600-dpi-for-print-ready-out.cs`
  - Task: Batch convert multiple PDFs to high‑resolution PNG images with 600 DPI for print‑ready output.
  - Workflow: batch-bootstrap
  - Outputs: pdf, png
  - Selected engine: verified
- `save-a-pdf-as-a-compressed-pdf-a-2u-file-by-enabling-image-compression-and-removing-unused.cs`
  - Task: Save a PDF as a compressed PDF/A‑2u file by enabling image compression and removing unused objects.
  - Workflow: save-options
  - Outputs: pdf
  - Selected engine: verified
- `load-an-html-file-with-external-javascript-and-convert-it-to-pdf-ignoring-scripts-for-stat.cs`
  - Task: Load an HTML file with external JavaScript and convert it to PDF, ignoring scripts for static rendering.
  - Workflow: input-bootstrap
  - Outputs: pdf, html
  - Selected engine: verified
- `convert-a-pdf-to-an-xlsx-spreadsheet-extracting-any-tabular-data-into-separate-worksheets.cs`
  - Task: Convert a PDF to an XLSX spreadsheet, extracting any tabular data into separate worksheets.
  - Workflow: native-conversion-api
  - Outputs: pdf
  - Selected engine: verified
- `batch-process-docx-files-applying-a-company-wide-header-before-converting-each-to-pdf-for.cs`
  - Task: Batch process DOCX files, applying a company‑wide header before converting each to PDF for distribution.
  - Workflow: batch-bootstrap
  - Outputs: pdf, docx
  - Selected engine: verified
- `save-a-pdf-as-an-image-sequence-in-png-format-with-lossless-compression-for-archival-purpo.cs`
  - Task: Save a PDF as an image sequence in PNG format with lossless compression for archival purposes.
  - Workflow: native-conversion-api
  - Outputs: pdf, png
  - Selected engine: verified
- `convert-a-pdf-to-html-with-external-css-files-and-separate-image-folder-for-optimized-web.cs`
  - Task: Convert a PDF to HTML with external CSS files and separate image folder for optimized web delivery.
  - Workflow: batch-bootstrap
  - Outputs: pdf, html
  - Selected engine: verified
- `load-a-docx-from-a-sharepoint-stream-and-convert-it-to-pdf-for-on-demand-download.cs`
  - Task: Load a DOCX from a SharePoint stream and convert it to PDF for on‑demand download.
  - Workflow: stream-response
  - Outputs: pdf, docx
  - Selected engine: verified
- `batch-convert-a-set-of-pdf-files-to-epub-preserving-original-chapter-structure-for-e-readi.cs`
  - Task: Batch convert a set of PDF files to EPUB, preserving original chapter structure for e‑reading devices.
  - Workflow: batch-bootstrap
  - Outputs: pdf, epub
  - Selected engine: verified
- `apply-a-custom-pdf-a-2b-compliance-level-when-converting-doc-to-pdf-using-pdfsaveoptions.cs`
  - Task: Apply a custom PDF/A‑2b compliance level when converting DOC to PDF using PdfSaveOptions.
  - Workflow: save-options
  - Outputs: pdf
  - Selected engine: verified
- `convert-a-pdf-to-a-searchable-pdf-a-1a-by-enabling-ocr-and-setting-appropriate-compliance.cs`
  - Task: Convert a PDF to a searchable PDF/A‑1a by enabling OCR and setting appropriate compliance.
  - Workflow: save-options
  - Outputs: pdf
  - Selected engine: verified
- `save-a-docx-as-html-with-external-javascript-disabled-to-improve-security-of-the-generated.cs`
  - Task: Save a DOCX as HTML with external JavaScript disabled to improve security of the generated page.
  - Workflow: input-bootstrap
  - Outputs: docx, html
  - Selected engine: verified
- `batch-convert-multiple-png-images-to-a-single-pdf-arranging-each-image-on-a-separate-page.cs`
  - Task: Batch convert multiple PNG images to a single PDF, arranging each image on a separate page.
  - Workflow: batch-bootstrap
  - Outputs: pdf, png
  - Selected engine: verified
- `load-a-pdf-and-convert-it-to-an-xps-document-while-preserving-annotations-for-review.cs`
  - Task: Load a PDF and convert it to an XPS document while preserving annotations for review.
  - Workflow: native-conversion-api
  - Outputs: pdf, xps
  - Selected engine: verified
- `convert-a-docx-to-mhtml-and-then-embed-it-in-an-email-body-using-aspose-email-html-body-su.cs`
  - Task: Convert a DOCX to MHTML and then embed it in an email body using Aspose.Email HTML body support.
  - Workflow: input-bootstrap
  - Outputs: docx, html, mhtml
  - Selected engine: verified
- `save-a-pdf-as-a-pdf-a-3u-file-and-attach-an-associated-xml-metadata-file-for-compliance.cs`
  - Task: Save a PDF as a PDF/A‑3u file and attach an associated XML metadata file for compliance.
  - Workflow: input-bootstrap
  - Outputs: pdf, xml
  - Selected engine: verified
- `batch-process-html-files-converting-each-to-pdf-with-custom-page-size-defined-in-pdfsaveop.cs`
  - Task: Batch process HTML files, converting each to PDF with custom page size defined in PdfSaveOptions.
  - Workflow: batch-bootstrap
  - Outputs: pdf, html
  - Selected engine: verified
- `convert-a-pdf-to-a-high-resolution-jpeg-image-setting-jpegsaveoptions-compressionlevel-to.cs`
  - Task: Convert a PDF to a high‑resolution JPEG image, setting JpegSaveOptions.CompressionLevel to Low for quality.
  - Workflow: save-options
  - Outputs: pdf, jpeg
  - Selected engine: verified
- `load-a-docx-replace-all-placeholder-dates-with-current-date-and-export-to-pdf-for-reportin.cs`
  - Task: Load a DOCX, replace all placeholder dates with current date, and export to PDF for reporting.
  - Workflow: save-options
  - Outputs: pdf, docx
  - Selected engine: verified
- `load-a-docx-file-and-save-it-as-markdown-extracting-images-to-a-custom-folder.cs`
  - Task: Load a DOCX file and save it as Markdown, extracting images to a custom folder.
  - Workflow: batch-bootstrap
  - Outputs: docx, md
  - Selected engine: verified
- `convert-a-pdf-document-to-docx-format-while-preserving-text-formatting-and-hyperlinks.cs`
  - Task: Convert a PDF document to DOCX format while preserving text formatting and hyperlinks.
  - Workflow: native-conversion-api
  - Outputs: pdf, docx, doc
  - Selected engine: verified
- `load-a-pdf-with-pdfloadoptions-to-ignore-password-protection-and-save-it-as-a-docx-file.cs`
  - Task: Load a PDF with PdfLoadOptions to ignore password protection and save it as a DOCX file.
  - Workflow: input-bootstrap
  - Outputs: pdf, docx
  - Selected engine: verified
- `load-a-pdf-with-password-protection-provide-password-via-pdfloadoptions-and-convert-to-png.cs`
  - Task: Load a PDF with password protection, provide password via PdfLoadOptions, and convert to PNG without errors.
  - Workflow: save-options
  - Outputs: pdf, png
  - Selected engine: verified
- `load-multiple-pdf-files-from-a-list-convert-each-to-docx-and-write-results-to-an-output-fo.cs`
  - Task: Load multiple PDF files from a list, convert each to DOCX, and write results to an output folder.
  - Workflow: batch-bootstrap
  - Outputs: pdf, docx, doc
  - Selected engine: verified
- `load-a-pdf-set-pdfloadoptions-to-ignore-errors-and-save-as-docx-while-preserving-page-brea.cs`
  - Task: Load a PDF, set PdfLoadOptions to ignore errors, and save as DOCX while preserving page breaks.
  - Workflow: save-options
  - Outputs: pdf, docx, doc
  - Selected engine: verified
- `export-a-pdf-to-a-single-high-quality-jpeg-image-using-multipagelayout-horizontal-and-cust.cs`
  - Task: Export a PDF to a single high‑quality JPEG image using MultiPageLayout.Horizontal and custom JpegQuality.
  - Workflow: save-options
  - Outputs: pdf, jpeg
  - Selected engine: verified
- `export-a-pdf-as-a-series-of-png-files-using-pageset-to-export-only-even-numbered-pages.cs`
  - Task: Export a PDF as a series of PNG files, using PageSet to export only even‑numbered pages.
  - Workflow: save-options
  - Outputs: pdf, png
  - Selected engine: verified
- `export-selected-pages-1-4-7-of-a-pdf-as-separate-png-images-with-custom-resolution-setting.cs`
  - Task: Export selected pages 1,4,7 of a PDF as separate PNG images with custom resolution settings.
  - Workflow: save-options
  - Outputs: pdf, png
  - Selected engine: verified
- `save-a-pdf-as-markdown-extracting-images-to-a-temporary-folder-and-embedding-relative-path.cs`
  - Task: Save a PDF as Markdown, extracting images to a temporary folder and embedding relative paths in the output.
  - Workflow: batch-bootstrap
  - Outputs: pdf, md
  - Selected engine: verified
- `convert-a-pdf-to-markdown-ensuring-that-all-extracted-images-are-saved-in-a-subfolder-name.cs`
  - Task: Convert a PDF to Markdown, ensuring that all extracted images are saved in a subfolder named "assets".
  - Workflow: batch-bootstrap
  - Outputs: pdf, md
  - Selected engine: verified
- `convert-a-pdf-to-tiff-using-lzw-compression-and-setting-imagecontrast-to-10-for-improved-c.cs`
  - Task: Convert a PDF to TIFF, using LZW compression and setting ImageContrast to 10 for improved clarity.
  - Workflow: native-conversion-api
  - Outputs: pdf, tiff
  - Selected engine: verified

## Common failure patterns and preferred agent fixes

- **System.Drawing or Font ambiguity**
  - Symptom: Compile errors caused by System.Drawing usage or ambiguous Font references between Aspose.Drawing.Font and Aspose.Words.Font.
  - Preferred fix: Use Aspose.Drawing only, and declare drawing fonts explicitly as Aspose.Drawing.Font.

- **Missing input bootstrap**
  - Symptom: Conversion task assumes a DOC, DOCX, HTML, TXT, XML, folder, or other source already exists.
  - Preferred fix: Create the local source file, source document, or input folder inside the example before loading it.

- **Response-stream task hangs**
  - Symptom: Example starts a server or waits for a request instead of finishing.
  - Preferred fix: Simulate the response with a MemoryStream, write conversion output to it, validate length, and exit.

- **Incorrect SaveFormat or extension pairing**
  - Symptom: Output format does not match the requested target type or file extension.
  - Preferred fix: Use the correct SaveFormat or specialized SaveOptions for the requested output.

- **Stream position or disposal issue**
  - Symptom: Reloading or reading a stream fails because Position was not reset or the stream was disposed too early.
  - Preferred fix: Reset MemoryStream.Position before reading and keep the stream alive until validation is complete.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`
- Additional package: `Aspose.Drawing.Common`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.Drawing.Common
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\conversion\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer direct Aspose.Words conversion APIs over speculative shortcuts.
