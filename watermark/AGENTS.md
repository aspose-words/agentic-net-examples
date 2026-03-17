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

- `using Aspose.Words;` (32/32 files) ← category-specific
- `using System;` (29/32 files)
- `using Aspose.Words.Drawing;` (25/32 files)
- `using System.Drawing;` (12/32 files)
- `using System.IO;` (10/32 files)
- `using Aspose.Words.Tables;` (4/32 files)
- `using Aspose.Words.Saving;` (4/32 files)
- `using Aspose.Words.Settings;` (2/32 files)
- `using System.Collections.Generic;` (1/32 files)
- `using System.Text.Json;` (1/32 files)
- `using Aspose.Words.Fields;` (1/32 files)

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
| [add-confidential-text-watermark-all-new-documents-creat...](./add-confidential-text-watermark-all-new-documents-created-automated-report-generator.cs) | `Document`, `Aspose`, `ReportGenerator` | Add confidential text watermark all new documents created automated report ge... |
| [add-text-watermark-docx-document-watermark-settext-cust...](./add-text-watermark-docx-document-watermark-settext-custom-font-settings.cs) | `Document`, `Aspose`, `TextWatermarkOptions` | Add text watermark docx document watermark settext custom font settings |
| [add-watermark-document-opened-network-share-ensuring-pr...](./add-watermark-document-opened-network-share-ensuring-proper-disposal-file-handles.cs) | `Document`, `System`, `Aspose` | Add watermark document opened network share ensuring proper disposal file han... |
| [add-watermark-table-cell-that-spans-multiple-rows-colum...](./add-watermark-table-cell-that-spans-multiple-rows-columns-complex-word-table.cs) | `Aspose`, `CellFormat`, `Document` | Add watermark table cell that spans multiple rows columns complex word table |
| [apply-image-watermark-word-document-then-document-as-docx](./apply-image-watermark-word-document-then-document-as-docx.cs) | `Document`, `Aspose`, `ImageWatermarkOptions` | Apply image watermark word document then document as docx |
| [apply-text-watermark-word-document-then-document-as-docx](./apply-text-watermark-word-document-then-document-as-docx.cs) | `Document`, `Aspose`, `Watermark` | Apply text watermark word document then document as docx |
| [batch-convert-docx-files-pdf-while-adding-corporate-log...](./batch-convert-docx-files-pdf-while-adding-corporate-logo-image-watermark-each-pdf.cs) | `Document`, `Aspose`, `ImageWatermarkOptions` | Batch convert docx files pdf while adding corporate logo image watermark each... |
| [batch-process-folder-doc-files-add-same-image-watermark...](./batch-process-folder-doc-files-add-same-image-watermark-each-document.cs) | `Document`, `Aspose`, `Watermark` | Batch process folder doc files add same image watermark each document |
| [batch-process-multiple-word-documents-directory-add-tex...](./batch-process-multiple-word-documents-directory-add-text-watermark-each-file.cs) | `Document`, `System`, `Aspose` | Batch process multiple word documents directory add text watermark each file |
| [batch-process-multiple-word-documents-directory-remove-...](./batch-process-multiple-word-documents-directory-remove-existing-watermarks-each-file.cs) | `Document`, `Watermark`, `System` | Batch process multiple word documents directory remove existing watermarks ea... |
| [combine-text-image-watermarks-first-setting-text-waterm...](./combine-text-image-watermarks-first-setting-text-watermark-then-overlaying-image.cs) | `Aspose`, `Document`, `Words` | Combine text image watermarks first setting text watermark then overlaying image |
| [command-line-tool-that-accepts-directory-path-adds-spec...](./command-line-tool-that-accepts-directory-path-adds-specified-watermark-each-file.cs) | `Document`, `Aspose`, `WatermarkBatchTool` | Command line tool that accepts directory path adds specified watermark each file |
| [configuration-file-define-watermark-text-font-opacity-t...](./configuration-file-define-watermark-text-font-opacity-then-apply-it-multiple-documents.cs) | `System`, `Document`, `TextWatermarkOptions` | Configuration file define watermark text font opacity then apply it multiple... |
| [customize-image-watermark-opacity-scaling-configuring-i...](./customize-image-watermark-opacity-scaling-configuring-imagewatermarkoptions-before.cs) | `Aspose`, `Document`, `ImageWatermarkOptions` | Customize image watermark opacity scaling configuring imagewatermarkoptions b... |
| [implement-unit-test-that-verifies-watermark-remove-succ...](./implement-unit-test-that-verifies-watermark-remove-successfully-deletes-previously.cs) | `Aspose`, `Watermark`, `Words` | Implement unit test that verifies watermark remove successfully deletes previ... |
| [insert-image-watermark-file-path-word-document-after-op...](./insert-image-watermark-file-path-word-document-after-optimizing-document.cs) | `Aspose`, `Document`, `Words` | Insert image watermark file path word document after optimizing document |
| [insert-text-watermark-specific-table-cell-within-word-d...](./insert-text-watermark-specific-table-cell-within-word-document-watermark-class.cs) | `Aspose`, `Document`, `DocumentBuilder` | Insert text watermark specific table cell within word document watermark class |
| [insert-watermark-each-cell-first-row-table-watermark-class](./insert-watermark-each-cell-first-row-table-watermark-class.cs) | `Aspose`, `Document`, `DocumentBuilder` | Insert watermark each cell first row table watermark class |
| [insert-watermark-table-cell-that-contains-merged-cells-...](./insert-watermark-table-cell-that-contains-merged-cells-without-disrupting-table-layout.cs) | `CellFormat`, `CellMerge`, `Aspose` | Insert watermark table cell that contains merged cells without disrupting tab... |
| [optimize-large-docx-file-before-applying-image-watermar...](./optimize-large-docx-file-before-applying-image-watermark-improve-performance-memory.cs) | `Aspose`, `Document`, `Words` | Optimize large docx file before applying image watermark improve performance... |
| [remove-all-existing-watermarks-loaded-word-document-wat...](./remove-all-existing-watermarks-loaded-word-document-watermark-remove-method.cs) | `Document`, `Watermark`, `Aspose` | Remove all existing watermarks loaded word document watermark remove method |
| [reusable-method-that-adds-configurable-text-watermark-a...](./reusable-method-that-adds-configurable-text-watermark-any-document-object.cs) | `Aspose`, `Watermark`, `ArgumentNullException` | Reusable method that adds configurable text watermark any document object |
| [utility-method-that-removes-all-watermarks-document-wat...](./utility-method-that-removes-all-watermarks-document-watermark-remove.cs) | `Document`, `Watermark`, `Aspose` | Utility method that removes all watermarks document watermark remove |
| [validate-that-document-contains-no-watermarks-before-pu...](./validate-that-document-contains-no-watermarks-before-publishing-watermarktype-none.cs) | `Document`, `Aspose`, `WatermarkType` | Validate that document contains no watermarks before publishing watermarktype... |
| [watermark-setimage-byte-array-stream-embed-dynamically-...](./watermark-setimage-byte-array-stream-embed-dynamically-barcode-watermark.cs) | `Aspose`, `Document`, `Words` | Watermark setimage byte array stream embed dynamically barcode watermark |
| [watermark-settext-textwatermarkoptions-set-watermark-fo...](./watermark-settext-textwatermarkoptions-set-watermark-font-size-color-spacing.cs) | `Document`, `TextWatermarkOptions`, `Aspose` | Watermark settext textwatermarkoptions set watermark font size color spacing |
| [watermarked-word-document-directly-pdf-format-while-pre...](./watermarked-word-document-directly-pdf-format-while-preserving-watermark-appearance.cs) | `Document`, `Aspose`, `SaveFormat` | Watermarked word document directly pdf format while preserving watermark appe... |
| [watermarktype-enumeration-switch-between-text-image-wat...](./watermarktype-enumeration-switch-between-text-image-watermarks-based-user-selection.cs) | `WatermarkType`, `Watermark`, `Document` | Watermarktype enumeration switch between text image watermarks based user sel... |
| [watermarktype-enumeration-verify-document-has-no-waterm...](./watermarktype-enumeration-verify-document-has-no-watermark-before-adding-new-one.cs) | `Watermark`, `Document`, `Aspose` | Watermarktype enumeration verify document has no watermark before adding new one |
| [word-document-file-path-add-image-watermark-watermark-s...](./word-document-file-path-add-image-watermark-watermark-setimage.cs) | `Aspose`, `Document`, `Watermark` | Word document file path add image watermark watermark setimage |
| ... | | *and 2 more files* |

## Category Statistics
- Total examples: 32

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for watermark patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082644`
<!-- AUTOGENERATED:END -->
