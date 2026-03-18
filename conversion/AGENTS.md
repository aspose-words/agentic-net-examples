---
name: conversion
description: C# examples for conversion using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - conversion

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **conversion** category.
This folder contains standalone C# examples for conversion operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **conversion**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (100/100 files) ← category-specific
- `using System;` (95/100 files)
- `using Aspose.Words.Saving;` (88/100 files)
- `using System.IO;` (39/100 files)
- `using Aspose.Words.Loading;` (12/100 files)
- `using Aspose.Words.Drawing;` (4/100 files)
- `using System.Text;` (4/100 files)
- `using System.Net.Http;` (3/100 files)
- `using System.Collections.Generic;` (2/100 files)
- `using System.Drawing;` (2/100 files)
- `using Aspose.Words.Replacing;` (1/100 files)
- `using Aspose.Words.Rendering;` (1/100 files)

## Common Code Pattern

Most files follow this pattern:

```csharp
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
// ... operations ...
doc.Save("output.docx");
```

## Files in this folder

| File | Key APIs | Description |
|------|----------|-------------|
| [add-header-footer-docx-before-converting-pdf-documentbu...](./add-header-footer-docx-before-converting-pdf-documentbuilder.cs) | `Document`, `DocumentBuilder`, `Aspose` | Add header footer docx before converting pdf documentbuilder |
| [apply-compression-xlsx-file-docx-setting-xlsxsaveoption...](./apply-compression-xlsx-file-docx-setting-xlsxsaveoptions-compressionlevel-fast.cs) | `Document`, `XlsxSaveOptions`, `Aspose` | Apply compression xlsx file docx setting xlsxsaveoptions compressionlevel fast |
| [apply-custom-page-size-when-converting-doc-pdf-setting-...](./apply-custom-page-size-when-converting-doc-pdf-setting-pdfsaveoptions-pagesize.cs) | `PageSetup`, `Document`, `DocumentBuilder` | Apply custom page size when converting doc pdf setting pdfsaveoptions pagesize |
| [apply-custom-pdf-2b-compliance-level-when-converting-do...](./apply-custom-pdf-2b-compliance-level-when-converting-doc-pdf-pdfsaveoptions.cs) | `Document`, `Aspose`, `PdfSaveOptions` | Apply custom pdf 2b compliance level when converting doc pdf pdfsaveoptions |
| [apply-find-replace-operation-word-document-before-expor...](./apply-find-replace-operation-word-document-before-exporting-it-excel-format.cs) | `Document`, `Aspose`, `SaveFormat` | Apply find replace operation word document before exporting it excel format |
| [batch-convert-all-docx-files-directory-html-round-trip-...](./batch-convert-all-docx-files-directory-html-round-trip-information-enabled.cs) | `HtmlSaveOptions`, `Document`, `Aspose` | Batch convert all docx files directory html round trip information enabled |
| [batch-convert-all-html-files-directory-mhtml-embedding-...](./batch-convert-all-html-files-directory-mhtml-embedding-resources-automatically-each.cs) | `Document`, `Aspose`, `HtmlSaveOptions` | Batch convert all html files directory mhtml embedding resources automaticall... |
| [batch-convert-collection-html-files-mhtml-ensuring-all-...](./batch-convert-collection-html-files-mhtml-ensuring-all-linked-resources-are-embedded.cs) | `Document`, `HtmlSaveOptions`, `Aspose` | Batch convert collection html files mhtml ensuring all linked resources are e... |
| [batch-convert-collection-png-images-single-pdf-document...](./batch-convert-collection-png-images-single-pdf-document-each-image-separate-page.cs) | `Aspose`, `Document`, `DocumentBuilder` | Batch convert collection png images single pdf document each image separate page |
| [batch-convert-html-files-epub-format-creating-collectio...](./batch-convert-html-files-epub-format-creating-collection-e-books-web-content.cs) | `Document`, `HtmlSaveOptions`, `System` | Batch convert html files epub format creating collection e books web content |
| [batch-convert-html-files-pdf-custom-page-margins-define...](./batch-convert-html-files-pdf-custom-page-margins-defined-pdfsaveoptions.cs) | `PageSetup`, `Document`, `PdfSaveOptions` | Batch convert html files pdf custom page margins defined pdfsaveoptions |
| [batch-convert-multiple-pdfs-high-resolution-png-images-...](./batch-convert-multiple-pdfs-high-resolution-png-images-600-dpi-print-ready-output.cs) | `ImageSaveOptions`, `Document`, `Aspose` | Batch convert multiple pdfs high resolution png images 600 dpi print ready ou... |
| [batch-convert-multiple-pdfs-html-files-preserving-origi...](./batch-convert-multiple-pdfs-html-files-preserving-original-layout-fonts-htmlsaveoptions.cs) | `ArgumentException`, `HtmlSaveOptions`, `Document` | Batch convert multiple pdfs html files preserving original layout fonts htmls... |
| [batch-convert-multiple-png-images-single-pdf-arranging-...](./batch-convert-multiple-png-images-single-pdf-arranging-each-image-separate-page.cs) | `Aspose`, `Document`, `DocumentBuilder` | Batch convert multiple png images single pdf arranging each image separate page |
| [batch-convert-set-pdf-files-epub-preserving-original-ch...](./batch-convert-set-pdf-files-epub-preserving-original-chapter-structure-e-reading.cs) | `Document`, `HtmlSaveOptions`, `System` | Batch convert set pdf files epub preserving original chapter structure e reading |
| [batch-convert-set-rtf-files-pdf-1a-compliance-legal-doc...](./batch-convert-set-rtf-files-pdf-1a-compliance-legal-document-archiving.cs) | `Aspose`, `RtfLoadOptions`, `Document` | Batch convert set rtf files pdf 1a compliance legal document archiving |
| [batch-process-all-rtf-files-folder-converting-each-pdf-...](./batch-process-all-rtf-files-folder-converting-each-pdf-default-layout.cs) | `Document`, `Aspose`, `System` | Batch process all rtf files folder converting each pdf default layout |
| [batch-process-docx-files-applying-company-wide-header-b...](./batch-process-docx-files-applying-company-wide-header-before-converting-each-pdf.cs) | `HeaderFooterType`, `Document`, `DocumentBuilder` | Batch process docx files applying company wide header before converting each pdf |
| [batch-process-folder-doc-files-converting-each-pdf-logg...](./batch-process-folder-doc-files-converting-each-pdf-logging-conversion-status.cs) | `Aspose`, `Document`, `System` | Batch process folder doc files converting each pdf logging conversion status |
| [batch-process-html-files-converting-each-pdf-custom-pag...](./batch-process-html-files-converting-each-pdf-custom-page-size-defined-pdfsaveoptions.cs) | `Aspose`, `Words`, `PageSetup` | Batch process html files converting each pdf custom page size defined pdfsave... |
| [batch-process-pdfs-jpeg-thumbnails-first-page-jpegsaveo...](./batch-process-pdfs-jpeg-thumbnails-first-page-jpegsaveoptions-low-quality.cs) | `Document`, `ImageSaveOptions`, `Aspose` | Batch process pdfs jpeg thumbnails first page jpegsaveoptions low quality |
| [convert-doc-file-pdf-1b-setting-pdfsaveoptions-complian...](./convert-doc-file-pdf-1b-setting-pdfsaveoptions-compliance-before.cs) | `Document`, `Aspose`, `PdfSaveOptions` | Convert doc file pdf 1b setting pdfsaveoptions compliance before |
| [convert-doc-file-xlsx-workbook-default-compression-xlsx...](./convert-doc-file-xlsx-workbook-default-compression-xlsxsaveoptions-compressionlevel.cs) | `Document`, `XlsxSaveOptions`, `Aspose` | Convert doc file xlsx workbook default compression xlsxsaveoptions compressio... |
| [convert-docx-file-mhtml-format-automatically-embedding-...](./convert-docx-file-mhtml-format-automatically-embedding-images-fonts-within-output.cs) | `Aspose`, `Document`, `HtmlSaveOptions` | Convert docx file mhtml format automatically embedding images fonts within ou... |
| [convert-docx-mhtml-automatically-embed-all-linked-css-f...](./convert-docx-mhtml-automatically-embed-all-linked-css-files-within-output.cs) | `Aspose`, `Document`, `HtmlSaveOptions` | Convert docx mhtml automatically embed all linked css files within output |
| [convert-docx-pdf-embed-custom-cover-page-image-document...](./convert-docx-pdf-embed-custom-cover-page-image-documentbuilder-insertion.cs) | `Document`, `DocumentBuilder`, `Aspose` | Convert docx pdf embed custom cover page image documentbuilder insertion |
| [convert-docx-pdf-embed-custom-font-setting-fontembeddin...](./convert-docx-pdf-embed-custom-font-setting-fontembeddingmode-embedallfonts.cs) | `Document`, `Aspose`, `PdfSaveOptions` | Convert docx pdf embed custom font setting fontembeddingmode embedallfonts |
| [convert-html-file-pdf-while-preserving-css-styles-html-...](./convert-html-file-pdf-while-preserving-css-styles-html-saveformat-pdf.cs) | `Document`, `Aspose`, `SaveFormat` | Convert html file pdf while preserving css styles html saveformat pdf |
| [convert-large-docx-pdf-streaming-minimize-memory-consum...](./convert-large-docx-pdf-streaming-minimize-memory-consumption-during-conversion.cs) | `Document`, `Aspose`, `LargeDocument` | Convert large docx pdf streaming minimize memory consumption during conversion |
| [convert-multiple-image-files-png-jpeg-single-pdf-docume...](./convert-multiple-image-files-png-jpeg-single-pdf-document-documentbuilder-insertimage.cs) | `Aspose`, `StringComparison`, `Document` | Convert multiple image files png jpeg single pdf document documentbuilder ins... |
| ... | | *and 70 more files* |

## Category Statistics
- Total examples: 100

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for conversion patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082331`
<!-- AUTOGENERATED:END -->
