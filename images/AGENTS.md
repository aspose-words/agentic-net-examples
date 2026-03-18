---
name: images
description: C# examples for images using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - images

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **images** category.
This folder contains standalone C# examples for images operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **images**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (88/88 files) ← category-specific
- `using System;` (87/88 files)
- `using System.IO;` (77/88 files)
- `using Aspose.Words.Drawing;` (63/88 files)
- `using Aspose.Words.Saving;` (52/88 files)
- `using System.Linq;` (29/88 files)
- `using System.Collections.Generic;` (14/88 files)
- `using System.Text;` (8/88 files)
- `using System.IO.Compression;` (5/88 files)
- `using Aspose.Words.Rendering;` (5/88 files)
- `using System.Drawing;` (3/88 files)
- `using Aspose.Words.Loading;` (3/88 files)

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
| [apply-border-5-pixels-red-color-all-extracted-png-image...](./apply-border-5-pixels-red-color-all-extracted-png-images-before.cs) | `ImageData`, `System`, `Aspose` | Apply border 5 pixels red color all extracted png images before |
| [apply-color-balance-adjustment-all-extracted-png-images...](./apply-color-balance-adjustment-all-extracted-png-images-before-them-output-folder.cs) | `Document`, `ImageSaveOptions`, `Aspose` | Apply color balance adjustment all extracted png images before them output fo... |
| [apply-contrast-enhancement-filter-all-extracted-png-ima...](./apply-contrast-enhancement-filter-all-extracted-png-images-before-them-disk.cs) | `ImageData`, `Aspose`, `Document` | Apply contrast enhancement filter all extracted png images before them disk |
| [apply-grayscale-filter-all-jpeg-images-extracted-doc-fi...](./apply-grayscale-filter-all-jpeg-images-extracted-doc-files-before-them.cs) | `Aspose`, `Document`, `Words` | Apply grayscale filter all jpeg images extracted doc files before them |
| [apply-lossless-compression-tiff-images-extracted-rtf-fi...](./apply-lossless-compression-tiff-images-extracted-rtf-files-store-them-archive.cs) | `Document`, `ImageSaveOptions`, `ZipArchive` | Apply lossless compression tiff images extracted rtf files store them archive |
| [batch-convert-all-extracted-images-collection-word-file...](./batch-convert-all-extracted-images-collection-word-files-webp-format-web.cs) | `Document`, `System`, `Aspose` | Batch convert all extracted images collection word files webp format web |
| [batch-convert-extracted-bmp-images-jpeg-80-quality-log-...](./batch-convert-extracted-bmp-images-jpeg-80-quality-log-conversion-results.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Batch convert extracted bmp images jpeg 80 quality log conversion results |
| [batch-convert-extracted-bmp-images-webp-lossless-compre...](./batch-convert-extracted-bmp-images-webp-lossless-compression-log-conversion-details.cs) | `Aspose`, `StreamWriter`, `FileInfo` | Batch convert extracted bmp images webp lossless compression log conversion d... |
| [batch-convert-extracted-gif-images-animated-webp-files-...](./batch-convert-extracted-gif-images-animated-webp-files-while-preserving-original.cs) | `Aspose`, `Document`, `ImageSaveOptions` | Batch convert extracted gif images animated webp files while preserving original |
| [batch-convert-extracted-tiff-images-jpeg-90-quality-sto...](./batch-convert-extracted-tiff-images-jpeg-90-quality-store-them-output-directory.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Batch convert extracted tiff images jpeg 90 quality store them output directory |
| [batch-extract-images-collection-docx-files-html-index-page](./batch-extract-images-collection-docx-files-html-index-page.cs) | `System`, `Aspose`, `Document` | Batch extract images collection docx files html index page |
| [batch-extract-images-doc-files-organize-them-subfolders...](./batch-extract-images-doc-files-organize-them-subfolders-based-image-format-type.cs) | `Aspose`, `Document`, `System` | Batch extract images doc files organize them subfolders based image format type |
| [batch-extract-images-set-docx-files-pdf-catalog-thumbnails](./batch-extract-images-set-docx-files-pdf-catalog-thumbnails.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Batch extract images set docx files pdf catalog thumbnails |
| [batch-extract-images-set-odt-files-markdown-gallery-thu...](./batch-extract-images-set-odt-files-markdown-gallery-thumbnails.cs) | `Aspose`, `Document`, `ImageRenamer` | Batch extract images set odt files markdown gallery thumbnails |
| [batch-extract-images-set-odt-files-organize-them-origin...](./batch-extract-images-set-odt-files-organize-them-original-document-name.cs) | `Document`, `System`, `Aspose` | Batch extract images set odt files organize them original document name |
| [batch-extract-images-set-odt-files-searchable-pdf-catalog](./batch-extract-images-set-odt-files-searchable-pdf-catalog.cs) | `Document`, `Aspose`, `ImageData` | Batch extract images set odt files searchable pdf catalog |
| [batch-extract-images-set-pdf-files-rename-them-source-d...](./batch-extract-images-set-pdf-files-rename-them-source-document-title.cs) | `Aspose`, `Words`, `PdfLoadOptions` | Batch extract images set pdf files rename them source document title |
| [batch-process-collection-docx-files-extracting-images-c...](./batch-process-collection-docx-files-extracting-images-creating-summary-pdf-catalog.cs) | `Font`, `Document`, `System` | Batch process collection docx files extracting images creating summary pdf ca... |
| [batch-process-doc-files-extracting-images-creating-comp...](./batch-process-doc-files-extracting-images-creating-compressed-zip-archive-password.cs) | `ImageType`, `Aspose`, `ZipArchive` | Batch process doc files extracting images creating compressed zip archive pas... |
| [batch-process-doc-files-extracting-images-creating-zip-...](./batch-process-doc-files-extracting-images-creating-zip-archive-containing-all.cs) | `System`, `Document`, `Aspose` | Batch process doc files extracting images creating zip archive containing all |
| [batch-process-docx-files-extracting-images-creating-sum...](./batch-process-docx-files-extracting-images-creating-summary-csv-containing-image-sizes.cs) | `System`, `Document`, `Aspose` | Batch process docx files extracting images creating summary csv containing im... |
| [batch-process-docx-files-extracting-images-generating-j...](./batch-process-docx-files-extracting-images-generating-json-manifest-containing-image.cs) | `System`, `Document`, `Aspose` | Batch process docx files extracting images generating json manifest containin... |
| [batch-process-multiple-doc-files-extracting-images-gene...](./batch-process-multiple-doc-files-extracting-images-generating-consolidated-pdf-report.cs) | `Document`, `Aspose`, `DocumentBuilder` | Batch process multiple doc files extracting images generating consolidated pd... |
| [batch-process-multiple-docx-files-extracting-images-gen...](./batch-process-multiple-docx-files-extracting-images-generating-csv-report-image.cs) | `ImageData`, `System`, `Aspose` | Batch process multiple docx files extracting images generating csv report image |
| [batch-process-multiple-odt-files-extracting-images-gene...](./batch-process-multiple-odt-files-extracting-images-generating-consolidated-json.cs) | `System`, `ImageData`, `Document` | Batch process multiple odt files extracting images generating consolidated json |
| [convert-each-extracted-png-image-jpeg-format-while-pres...](./convert-each-extracted-png-image-jpeg-format-while-preserving-its-original-dimensions.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Convert each extracted png image jpeg format while preserving its original di... |
| [convert-extracted-bmp-images-png-format-while-reducing-...](./convert-extracted-bmp-images-png-format-while-reducing-color-depth-256-colors.cs) | `Aspose`, `Document`, `DocumentBuilder` | Convert extracted bmp images png format while reducing color depth 256 colors |
| [convert-extracted-jpeg-images-grayscale-bmp-files-store...](./convert-extracted-jpeg-images-grayscale-bmp-files-store-them-secure-archive.cs) | `Document`, `System`, `Aspose` | Convert extracted jpeg images grayscale bmp files store them secure archive |
| [convert-extracted-jpeg-images-high-quality-webp-optimiz...](./convert-extracted-jpeg-images-high-quality-webp-optimized-web-delivery.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Convert extracted jpeg images high quality webp optimized web delivery |
| [convert-extracted-jpeg-images-high-resolution-tiff-arch...](./convert-extracted-jpeg-images-high-resolution-tiff-archival-storage-lzw-compression.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Convert extracted jpeg images high resolution tiff archival storage lzw compr... |
| ... | | *and 58 more files* |

## Category Statistics
- Total examples: 88

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for images patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082423`
<!-- AUTOGENERATED:END -->
