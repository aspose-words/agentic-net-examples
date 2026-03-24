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

- `using Aspose.Words;` (95/95 files) ← category-specific
- `using System;` (92/95 files)
- `using Aspose.Words.Saving;` (86/95 files)
- `using System.IO;` (69/95 files)
- `using Aspose.Words.Loading;` (11/95 files)
- `using System.Text;` (4/95 files)
- `using Aspose.Words.Drawing;` (3/95 files)
- `using System.Collections.Generic;` (2/95 files)
- `using System.Linq;` (1/95 files)
- `using Aspose.Words.Replacing;` (1/95 files)
- `using System.Net;` (1/95 files)
- `using System.Net.Mail;` (1/95 files)

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
| [add-header-footer-docx-before-converting-pdf-documentbu...](./add-header-footer-docx-before-converting-pdf-documentbuilder.cs) | `Document`, `DocumentBuilder`, `HeaderFooterType` | Add header footer docx before converting pdf documentbuilder |
| [apply-compression-xlsx-file-docx-setting-xlsxsaveoption...](./apply-compression-xlsx-file-docx-setting-xlsxsaveoptions-compressionlevel-fast.cs) | `Document`, `DocumentBuilder`, `XlsxSaveOptions` | Apply compression xlsx file docx setting xlsxsaveoptions compressionlevel fast |
| [apply-custom-page-size-when-converting-doc-pdf-setting-...](./apply-custom-page-size-when-converting-doc-pdf-setting-pdfsaveoptions-pagesize.cs) | `PageSetup`, `Document`, `DocumentBuilder` | Apply custom page size when converting doc pdf setting pdfsaveoptions pagesize |
| [apply-custom-pdf-2b-compliance-level-when-converting-do...](./apply-custom-pdf-2b-compliance-level-when-converting-doc-pdf-pdfsaveoptions.cs) | `Document`, `DocumentBuilder`, `PdfCompliance` | Apply custom pdf 2b compliance level when converting doc pdf pdfsaveoptions |
| [batch-convert-all-docx-files-directory-html-round-trip-...](./batch-convert-all-docx-files-directory-html-round-trip-information-enabled.cs) | `HtmlSaveOptions`, `Document`, `AppDomain` | Batch convert all docx files directory html round trip information enabled |
| [batch-convert-all-html-files-directory-mhtml-embedding-...](./batch-convert-all-html-files-directory-mhtml-embedding-resources-automatically-each.cs) | `Document`, `HtmlSaveOptions`, `MHTML` | Batch convert all html files directory mhtml embedding resources automaticall... |
| [batch-convert-collection-html-files-mhtml-ensuring-all-...](./batch-convert-collection-html-files-mhtml-ensuring-all-linked-resources-are-embedded.cs) | `Document`, `HtmlSaveOptions`, `MHTML` | Batch convert collection html files mhtml ensuring all linked resources are e... |
| [batch-convert-collection-png-images-single-pdf-document...](./batch-convert-collection-png-images-single-pdf-document-each-image-separate-page.cs) | `Document`, `DocumentBuilder`, `PdfSaveOptions` | Batch convert collection png images single pdf document each image separate page |
| [batch-convert-html-files-epub-format-creating-collectio...](./batch-convert-html-files-epub-format-creating-collection-e-books-web-content.cs) | `Document`, `DirectoryNotFoundException`, `EPUB` | Batch convert html files epub format creating collection e books web content |
| [batch-convert-html-files-pdf-custom-page-margins-define...](./batch-convert-html-files-pdf-custom-page-margins-defined-pdfsaveoptions.cs) | `PageSetup`, `Document`, `PdfSaveOptions` | Batch convert html files pdf custom page margins defined pdfsaveoptions |
| [batch-convert-multiple-pdfs-high-resolution-png-images-...](./batch-convert-multiple-pdfs-high-resolution-png-images-600-dpi-print-ready-output.cs) | `Document`, `ImageSaveOptions`, `SaveFormat` | Batch convert multiple pdfs high resolution png images 600 dpi print ready ou... |
| [batch-convert-multiple-pdfs-html-files-preserving-origi...](./batch-convert-multiple-pdfs-html-files-preserving-original-layout-fonts-htmlsaveoptions.cs) | `ArgumentException`, `HtmlSaveOptions`, `Document` | Batch convert multiple pdfs html files preserving original layout fonts htmls... |
| [batch-convert-multiple-png-images-single-pdf-arranging-...](./batch-convert-multiple-png-images-single-pdf-arranging-each-image-separate-page.cs) | `Document`, `DocumentBuilder`, `PdfSaveOptions` | Batch convert multiple png images single pdf arranging each image separate page |
| [batch-convert-set-pdf-files-epub-preserving-original-ch...](./batch-convert-set-pdf-files-epub-preserving-original-chapter-structure-e-reading.cs) | `Document`, `HtmlSaveOptions`, `AppContext` | Batch convert set pdf files epub preserving original chapter structure e reading |
| [batch-convert-set-rtf-files-pdf-1a-compliance-legal-doc...](./batch-convert-set-rtf-files-pdf-1a-compliance-legal-document-archiving.cs) | `RtfLoadOptions`, `Document`, `AppContext` | Batch convert set rtf files pdf 1a compliance legal document archiving |
| [batch-process-all-rtf-files-folder-converting-each-pdf-...](./batch-process-all-rtf-files-folder-converting-each-pdf-default-layout.cs) | `Document`, `SaveFormat`, `DocumentBuilder` | Batch process all rtf files folder converting each pdf default layout |
| [batch-process-docx-files-applying-company-wide-header-b...](./batch-process-docx-files-applying-company-wide-header-before-converting-each-pdf.cs) | `HeaderFooterType`, `Document`, `DocumentBuilder` | Batch process docx files applying company wide header before converting each pdf |
| [batch-process-folder-doc-files-converting-each-pdf-logg...](./batch-process-folder-doc-files-converting-each-pdf-logging-conversion-status.cs) | `Document`, `SearchOption`, `SaveFormat` | Batch process folder doc files converting each pdf logging conversion status |
| [batch-process-html-files-converting-each-pdf-custom-pag...](./batch-process-html-files-converting-each-pdf-custom-page-size-defined-pdfsaveoptions.cs) | `PageSetup`, `Document`, `PdfSaveOptions` | Batch process html files converting each pdf custom page size defined pdfsave... |
| [batch-process-pdfs-jpeg-thumbnails-first-page-jpegsaveo...](./batch-process-pdfs-jpeg-thumbnails-first-page-jpegsaveoptions-low-quality.cs) | `Document`, `ImageSaveOptions`, `AppContext` | Batch process pdfs jpeg thumbnails first page jpegsaveoptions low quality |
| [convert-doc-file-pdf-1b-setting-pdfsaveoptions-complian...](./convert-doc-file-pdf-1b-setting-pdfsaveoptions-compliance-before.cs) | `Document`, `PdfSaveOptions`, `PdfCompliance` | Convert doc file pdf 1b setting pdfsaveoptions compliance before |
| [convert-doc-file-xlsx-workbook-default-compression-xlsx...](./convert-doc-file-xlsx-workbook-default-compression-xlsxsaveoptions-compressionlevel.cs) | `Document`, `XlsxSaveOptions`, `Input` | Convert doc file xlsx workbook default compression xlsxsaveoptions compressio... |
| [convert-docx-file-mhtml-format-automatically-embedding-...](./convert-docx-file-mhtml-format-automatically-embedding-images-fonts-within-output.cs) | `Document`, `HtmlSaveOptions`, `MHTML` | Convert docx file mhtml format automatically embedding images fonts within ou... |
| [convert-docx-mhtml-automatically-embed-all-linked-css-f...](./convert-docx-mhtml-automatically-embed-all-linked-css-files-within-output.cs) | `Document`, `Sample`, `DocumentBuilder` | Convert docx mhtml automatically embed all linked css files within output |
| [convert-docx-pdf-embed-custom-cover-page-image-document...](./convert-docx-pdf-embed-custom-cover-page-image-documentbuilder-insertion.cs) | `Document`, `DocumentBuilder`, `ResultDocument` | Convert docx pdf embed custom cover page image documentbuilder insertion |
| [convert-docx-pdf-embed-custom-font-setting-fontembeddin...](./convert-docx-pdf-embed-custom-font-setting-fontembeddingmode-embedallfonts.cs) | `Document`, `Input`, `DocumentBuilder` | Convert docx pdf embed custom font setting fontembeddingmode embedallfonts |
| [convert-html-file-pdf-while-preserving-css-styles-html-...](./convert-html-file-pdf-while-preserving-css-styles-html-saveformat-pdf.cs) | `Document`, `CSS`, `Text` | Convert html file pdf while preserving css styles html saveformat pdf |
| [convert-large-docx-pdf-streaming-minimize-memory-consum...](./convert-large-docx-pdf-streaming-minimize-memory-consumption-during-conversion.cs) | `Document`, `DocumentBuilder`, `AppContext` | Convert large docx pdf streaming minimize memory consumption during conversion |
| [convert-multiple-image-files-png-jpeg-single-pdf-docume...](./convert-multiple-image-files-png-jpeg-single-pdf-document-documentbuilder-insertimage.cs) | `StringComparison`, `Document`, `DocumentBuilder` | Convert multiple image files png jpeg single pdf document documentbuilder ins... |
| [convert-pdf-containing-form-fields-docx-while-preservin...](./convert-pdf-containing-form-fields-docx-while-preserving-form-data-further-editing.cs) | `Document`, `AppDomain`, `CurrentDomain` | Convert pdf containing form fields docx while preserving form data further ed... |
| ... | | *and 65 more files* |

## Category Statistics
- Total examples: 95

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for conversion patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
