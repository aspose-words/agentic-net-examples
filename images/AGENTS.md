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

- `using System;` (74/75 files) ← category-specific
- `using Aspose.Words;` (73/75 files)
- `using System.IO;` (70/75 files)
- `using Aspose.Words.Drawing;` (55/75 files)
- `using Aspose.Words.Saving;` (38/75 files)
- `using System.Linq;` (34/75 files)
- `using System.Collections.Generic;` (10/75 files)
- `using System.Text;` (8/75 files)
- `using System.IO.Compression;` (5/75 files)
- `using Aspose.Words.Loading;` (3/75 files)
- `using Aspose.Words.Tables;` (3/75 files)
- `using Aspose.Words.Markup;` (3/75 files)

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
| [apply-color-balance-adjustment-all-extracted-png-images...](./apply-color-balance-adjustment-all-extracted-png-images-before-them-output-folder.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Apply color balance adjustment all extracted png images before them output fo... |
| [apply-contrast-enhancement-filter-all-extracted-png-ima...](./apply-contrast-enhancement-filter-all-extracted-png-images-before-them-disk.cs) | `ImageData`, `Document`, `Input` | Apply contrast enhancement filter all extracted png images before them disk |
| [apply-lossless-compression-tiff-images-extracted-rtf-fi...](./apply-lossless-compression-tiff-images-extracted-rtf-files-store-them-archive.cs) | `Document`, `ImageSaveOptions`, `ZipArchive` | Apply lossless compression tiff images extracted rtf files store them archive |
| [batch-convert-all-extracted-images-collection-word-file...](./batch-convert-all-extracted-images-collection-word-files-webp-format-web.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Batch convert all extracted images collection word files webp format web |
| [batch-convert-extracted-bmp-images-webp-lossless-compre...](./batch-convert-extracted-bmp-images-webp-lossless-compression-log-conversion-details.cs) | `StreamWriter`, `FileInfo`, `Document` | Batch convert extracted bmp images webp lossless compression log conversion d... |
| [batch-convert-extracted-gif-images-animated-webp-files-...](./batch-convert-extracted-gif-images-animated-webp-files-while-preserving-original.cs) | `Document`, `ImageSaveOptions`, `AppContext` | Batch convert extracted gif images animated webp files while preserving original |
| [batch-convert-extracted-tiff-images-jpeg-90-quality-sto...](./batch-convert-extracted-tiff-images-jpeg-90-quality-store-them-output-directory.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Batch convert extracted tiff images jpeg 90 quality store them output directory |
| [batch-extract-images-collection-docx-files-html-index-page](./batch-extract-images-collection-docx-files-html-index-page.cs) | `Document`, `HtmlSaveOptions`, `StringBuilder` | Batch extract images collection docx files html index page |
| [batch-extract-images-doc-files-organize-them-subfolders...](./batch-extract-images-doc-files-organize-them-subfolders-based-image-format-type.cs) | `Document`, `AppContext`, `StringComparison` | Batch extract images doc files organize them subfolders based image format type |
| [batch-extract-images-set-docx-files-pdf-catalog-thumbnails](./batch-extract-images-set-docx-files-pdf-catalog-thumbnails.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Batch extract images set docx files pdf catalog thumbnails |
| [batch-extract-images-set-odt-files-markdown-gallery-thu...](./batch-extract-images-set-odt-files-markdown-gallery-thumbnails.cs) | `Document`, `AppContext`, `ImageRenamer` | Batch extract images set odt files markdown gallery thumbnails |
| [batch-extract-images-set-odt-files-organize-them-origin...](./batch-extract-images-set-odt-files-organize-them-original-document-name.cs) | `Document`, `AppContext`, `ImageData` | Batch extract images set odt files organize them original document name |
| [batch-extract-images-set-odt-files-searchable-pdf-catalog](./batch-extract-images-set-odt-files-searchable-pdf-catalog.cs) | `Document`, `DocumentBuilder`, `ImageData` | Batch extract images set odt files searchable pdf catalog |
| [batch-extract-images-set-pdf-files-rename-them-source-d...](./batch-extract-images-set-pdf-files-rename-them-source-document-title.cs) | `Document`, `AppContext`, `ImageData` | Batch extract images set pdf files rename them source document title |
| [batch-process-collection-docx-files-extracting-images-c...](./batch-process-collection-docx-files-extracting-images-creating-summary-pdf-catalog.cs) | `Font`, `Document`, `DocumentBuilder` | Batch process collection docx files extracting images creating summary pdf ca... |
| [batch-process-doc-files-extracting-images-creating-comp...](./batch-process-doc-files-extracting-images-creating-compressed-zip-archive-password.cs) | `ImageType`, `ZipArchive`, `Document` | Batch process doc files extracting images creating compressed zip archive pas... |
| [batch-process-doc-files-extracting-images-creating-zip-...](./batch-process-doc-files-extracting-images-creating-zip-archive-containing-all.cs) | `Document`, `ImageData`, `Collections` | Batch process doc files extracting images creating zip archive containing all |
| [batch-process-docx-files-extracting-images-creating-sum...](./batch-process-docx-files-extracting-images-creating-summary-csv-containing-image-sizes.cs) | `Document`, `ImageData`, `StreamWriter` | Batch process docx files extracting images creating summary csv containing im... |
| [batch-process-multiple-doc-files-extracting-images-gene...](./batch-process-multiple-doc-files-extracting-images-generating-consolidated-pdf-report.cs) | `Document`, `DocumentBuilder`, `ImageData` | Batch process multiple doc files extracting images generating consolidated pd... |
| [batch-process-multiple-docx-files-extracting-images-gen...](./batch-process-multiple-docx-files-extracting-images-generating-csv-report-image.cs) | `ImageData`, `Document`, `StringBuilder` | Batch process multiple docx files extracting images generating csv report image |
| [convert-extracted-bmp-images-png-format-while-reducing-...](./convert-extracted-bmp-images-png-format-while-reducing-color-depth-256-colors.cs) | `AppDomain`, `CurrentDomain`, `Drawing` | Convert extracted bmp images png format while reducing color depth 256 colors |
| [convert-extracted-jpeg-images-grayscale-bmp-files-store...](./convert-extracted-jpeg-images-grayscale-bmp-files-store-them-secure-archive.cs) | `Document`, `ImageData`, `Shape` | Convert extracted jpeg images grayscale bmp files store them secure archive |
| [convert-extracted-jpeg-images-high-quality-webp-optimiz...](./convert-extracted-jpeg-images-high-quality-webp-optimized-web-delivery.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Convert extracted jpeg images high quality webp optimized web delivery |
| [convert-extracted-jpeg-images-high-resolution-tiff-arch...](./convert-extracted-jpeg-images-high-resolution-tiff-archival-storage-lzw-compression.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Convert extracted jpeg images high resolution tiff archival storage lzw compr... |
| [convert-extracted-tiff-images-grayscale-jpeg-low-bandwi...](./convert-extracted-tiff-images-grayscale-jpeg-low-bandwidth-environments.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Convert extracted tiff images grayscale jpeg low bandwidth environments |
| [convert-extracted-tiff-images-pdf-each-image-separate-p...](./convert-extracted-tiff-images-pdf-each-image-separate-page-embed-metadata.cs) | `AppContext`, `ArgumentException`, `StreamWriter` | Convert extracted tiff images pdf each image separate page embed metadata |
| [convert-extracted-tiff-images-pdf-pages-each-image-occu...](./convert-extracted-tiff-images-pdf-pages-each-image-occupying-full-page-output.cs) | `Document`, `DocumentBuilder`, `PDF` | Convert extracted tiff images pdf pages each image occupying full page output |
| [doc-file-extract-all-embedded-vector-images-convert-the...](./doc-file-extract-all-embedded-vector-images-convert-them-emf-format.cs) | `Document`, `DocumentBuilder`, `ImageData` | Doc file extract all embedded vector images convert them emf format |
| [docx-document-replace-all-images-placeholders-export-mo...](./docx-document-replace-all-images-placeholders-export-modified-document.cs) | `Document`, `DocumentBuilder`, `Run` | Docx document replace all images placeholders export modified document |
| [docx-file-extract-all-embedded-images-specified-output-...](./docx-file-extract-all-embedded-images-specified-output-folder.cs) | `Document`, `ImageData`, `AppContext` | Docx file extract all embedded images specified output folder |
| ... | | *and 45 more files* |

## Category Statistics
- Total examples: 75

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for images patterns.


## Command Reference

### Build and Run

Files in this folder are standalone `.cs` examples. Run one example at a time by copying it into a temporary console project as `Program.cs`.

```bash
# Create a temporary console project from the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\images\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

### Notes for Agents and Developers

- Treat every `.cs` file in `images/` as a full console program, not a snippet.
- Run one file at a time by copying it to `Program.cs`.
- If a sample needs input documents, images, fonts, or data files, place them in the temporary project directory before running.
- See the root [AGENTS.md](../AGENTS.md) for repository-wide prerequisites, project file template, and testing guidance.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
