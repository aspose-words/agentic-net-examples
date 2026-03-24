---
name: watermark
description: C# examples for watermark using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - watermark

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **watermark** category.
This folder contains standalone C# examples for watermark operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **watermark**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (28/28 files) ← category-specific
- `using System;` (26/28 files)
- `using Aspose.Words.Drawing;` (21/28 files)
- `using System.IO;` (11/28 files)
- `using System.Drawing;` (10/28 files)
- `using Aspose.Words.Saving;` (4/28 files)
- `using Aspose.Words.Tables;` (3/28 files)
- `using Aspose.Words.Settings;` (2/28 files)

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
| [add-confidential-text-watermark-all-new-documents-creat...](./add-confidential-text-watermark-all-new-documents-created-automated-report-generator.cs) | `Document`, `ReportGenerator`, `TextWatermarkOptions` | Add confidential text watermark all new documents created automated report ge... |
| [add-text-watermark-docx-document-watermark-settext-cust...](./add-text-watermark-docx-document-watermark-settext-custom-font-settings.cs) | `Document`, `TextWatermarkOptions`, `Watermarked` | Add text watermark docx document watermark settext custom font settings |
| [add-watermark-table-cell-that-spans-multiple-rows-colum...](./add-watermark-table-cell-that-spans-multiple-rows-columns-complex-word-table.cs) | `CellFormat`, `CellMerge`, `Document` | Add watermark table cell that spans multiple rows columns complex word table |
| [apply-image-watermark-word-document-then-document-as-docx](./apply-image-watermark-word-document-then-document-as-docx.cs) | `Document`, `ImageWatermarkOptions`, `Logo` | Apply image watermark word document then document as docx |
| [apply-text-watermark-word-document-then-document-as-docx](./apply-text-watermark-word-document-then-document-as-docx.cs) | `Document`, `Watermark`, `WatermarkedDocument` | Apply text watermark word document then document as docx |
| [batch-convert-docx-files-pdf-while-adding-corporate-log...](./batch-convert-docx-files-pdf-while-adding-corporate-logo-image-watermark-each-pdf.cs) | `Document`, `DocumentBuilder`, `AppContext` | Batch convert docx files pdf while adding corporate logo image watermark each... |
| [batch-process-folder-doc-files-add-same-image-watermark...](./batch-process-folder-doc-files-add-same-image-watermark-each-document.cs) | `Document`, `Watermark`, `ImageWatermarkOptions` | Batch process folder doc files add same image watermark each document |
| [batch-process-multiple-word-documents-directory-add-tex...](./batch-process-multiple-word-documents-directory-add-text-watermark-each-file.cs) | `Document`, `TextWatermarkOptions`, `AppContext` | Batch process multiple word documents directory add text watermark each file |
| [batch-process-multiple-word-documents-directory-remove-...](./batch-process-multiple-word-documents-directory-remove-existing-watermarks-each-file.cs) | `Document`, `Watermark`, `AppContext` | Batch process multiple word documents directory remove existing watermarks ea... |
| [combine-text-image-watermarks-first-setting-text-waterm...](./combine-text-image-watermarks-first-setting-text-watermark-then-overlaying-image.cs) | `Document`, `Watermark`, `TextWatermarkOptions` | Combine text image watermarks first setting text watermark then overlaying image |
| [command-line-tool-that-accepts-directory-path-adds-spec...](./command-line-tool-that-accepts-directory-path-adds-specified-watermark-each-file.cs) | `Document`, `WatermarkBatchTool`, `Watermark` | Command line tool that accepts directory path adds specified watermark each file |
| [customize-image-watermark-opacity-scaling-configuring-i...](./customize-image-watermark-opacity-scaling-configuring-imagewatermarkoptions-before.cs) | `Document`, `ImageWatermarkOptions`, `Logo` | Customize image watermark opacity scaling configuring imagewatermarkoptions b... |
| [implement-unit-test-that-verifies-watermark-remove-succ...](./implement-unit-test-that-verifies-watermark-remove-successfully-deletes-previously.cs) | `Watermark`, `Document`, `InvalidOperationException` | Implement unit test that verifies watermark remove successfully deletes previ... |
| [insert-image-watermark-file-path-word-document-after-op...](./insert-image-watermark-file-path-word-document-after-optimizing-document.cs) | `Document`, `ImageWatermarkOptions`, `Logo` | Insert image watermark file path word document after optimizing document |
| [insert-text-watermark-specific-table-cell-within-word-d...](./insert-text-watermark-specific-table-cell-within-word-document-watermark-class.cs) | `Document`, `DocumentBuilder`, `TextWatermarkOptions` | Insert text watermark specific table cell within word document watermark class |
| [insert-watermark-each-cell-first-row-table-watermark-class](./insert-watermark-each-cell-first-row-table-watermark-class.cs) | `Document`, `DocumentBuilder`, `TextWatermarkOptions` | Insert watermark each cell first row table watermark class |
| [optimize-large-docx-file-before-applying-image-watermar...](./optimize-large-docx-file-before-applying-image-watermark-improve-performance-memory.cs) | `Document`, `OoxmlSaveOptions`, `DocumentBuilder` | Optimize large docx file before applying image watermark improve performance... |
| [remove-all-existing-watermarks-loaded-word-document-wat...](./remove-all-existing-watermarks-loaded-word-document-watermark-remove-method.cs) | `Watermark`, `Document`, `WatermarkType` | Remove all existing watermarks loaded word document watermark remove method |
| [reusable-method-that-adds-configurable-text-watermark-a...](./reusable-method-that-adds-configurable-text-watermark-any-document-object.cs) | `Watermark`, `ArgumentNullException`, `Document` | Reusable method that adds configurable text watermark any document object |
| [utility-method-that-removes-all-watermarks-document-wat...](./utility-method-that-removes-all-watermarks-document-watermark-remove.cs) | `Document`, `Watermark`, `WatermarkType` | Utility method that removes all watermarks document watermark remove |
| [validate-that-document-contains-no-watermarks-before-pu...](./validate-that-document-contains-no-watermarks-before-publishing-watermarktype-none.cs) | `Document`, `WatermarkType`, `InvalidOperationException` | Validate that document contains no watermarks before publishing watermarktype... |
| [watermark-settext-textwatermarkoptions-set-watermark-fo...](./watermark-settext-textwatermarkoptions-set-watermark-font-size-color-spacing.cs) | `Document`, `TextWatermarkOptions`, `Color` | Watermark settext textwatermarkoptions set watermark font size color spacing |
| [watermarked-word-document-directly-pdf-format-while-pre...](./watermarked-word-document-directly-pdf-format-while-preserving-watermark-appearance.cs) | `Document`, `SaveFormat`, `TextWatermarkOptions` | Watermarked word document directly pdf format while preserving watermark appe... |
| [watermarktype-enumeration-switch-between-text-image-wat...](./watermarktype-enumeration-switch-between-text-image-watermarks-based-user-selection.cs) | `WatermarkType`, `Watermark`, `Document` | Watermarktype enumeration switch between text image watermarks based user sel... |
| [watermarktype-enumeration-verify-document-has-no-waterm...](./watermarktype-enumeration-verify-document-has-no-watermark-before-adding-new-one.cs) | `Watermark`, `Document`, `WatermarkType` | Watermarktype enumeration verify document has no watermark before adding new one |
| [word-document-file-path-add-image-watermark-watermark-s...](./word-document-file-path-add-image-watermark-watermark-setimage.cs) | `Document`, `DocumentBuilder`, `Watermark` | Word document file path add image watermark watermark setimage |
| [word-document-file-path-add-text-watermark-watermark-se...](./word-document-file-path-add-text-watermark-watermark-settext.cs) | `Document`, `DocumentBuilder`, `Watermark` | Word document file path add text watermark watermark settext |
| [word-document-memory-stream-apply-text-watermark-withou...](./word-document-memory-stream-apply-text-watermark-without-writing-disk.cs) | `Document`, `InputDocument`, `Watermark` | Word document memory stream apply text watermark without writing disk |

## Category Statistics
- Total examples: 28

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for watermark patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
