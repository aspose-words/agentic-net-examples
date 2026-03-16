---
name: split-document
description: C# examples for split-document using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - split-document

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **split-document** category.
This folder contains standalone C# examples for split-document operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **split-document**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using System;` (29/29 files) ← category-specific
- `using Aspose.Words;` (29/29 files)
- `using Aspose.Words.Saving;` (25/29 files)
- `using System.IO;` (18/29 files)
- `using System.Collections.Generic;` (3/29 files)
- `using Aspose.Words.Tables;` (2/29 files)
- `using System.Diagnostics;` (1/29 files)
- `using System.Linq;` (1/29 files)
- `using Aspose.Words.Rendering;` (1/29 files)
- `using Aspose.Words.Drawing;` (1/29 files)

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
| [after-splitting-open-each-output-document-programmatica...](./after-splitting-open-each-output-document-programmatically-verify-headers-footers.cs) | `Document`, `System`, `Aspose` | After splitting open each output document programmatically verify headers foo... |
| [apply-documentpartsavingcallback-customize-file-naming-...](./apply-documentpartsavingcallback-customize-file-naming-convention-each-split-output.cs) | `DocumentSplitCriteria`, `Aspose`, `Document` | Apply documentpartsavingcallback customize file naming convention each split... |
| [combine-heading-section-flags-documentsplitcriteria-spl...](./combine-heading-section-flags-documentsplitcriteria-split-both-structures.cs) | `ParagraphFormat`, `StyleIdentifier`, `Aspose` | Combine heading section flags documentsplitcriteria split both structures |
| [combine-page-heading-flags-documentsplitcriteria-start-...](./combine-page-heading-flags-documentsplitcriteria-start-each-part-new-page.cs) | `Document`, `HtmlSaveOptions`, `Aspose` | Combine page heading flags documentsplitcriteria start each part new page |
| [combine-page-range-heading-criteria-produce-parts-that-...](./combine-page-range-heading-criteria-produce-parts-that-start-at-each-heading-new-page.cs) | `StyleIdentifier`, `Aspose`, `ParagraphFormat` | Combine page range heading criteria produce parts that start at each heading... |
| [documentsplitcriteria-enumeration-split-sections-then-m...](./documentsplitcriteria-enumeration-split-sections-then-merge-selected-parts.cs) | `Document`, `DocumentSplitCriteria`, `Section` | Documentsplitcriteria enumeration split sections then merge selected parts |
| [documentsplitcriteria-object-set-split-mode-custom-page...](./documentsplitcriteria-object-set-split-mode-custom-page-ranges.cs) | `Document`, `Aspose`, `DocumentSplitCriteria` | Documentsplitcriteria object set split mode custom page ranges |
| [documentsplitcriteria-object-set-split-mode-headings](./documentsplitcriteria-object-set-split-mode-headings.cs) | `ParagraphFormat`, `Document`, `DocumentBuilder` | Documentsplitcriteria object set split mode headings |
| [documentsplitcriteria-object-set-split-mode-individual-...](./documentsplitcriteria-object-set-split-mode-individual-pages.cs) | `Document`, `HtmlSaveOptions`, `Aspose` | Documentsplitcriteria object set split mode individual pages |
| [documentsplitcriteria-object-set-split-mode-sections](./documentsplitcriteria-object-set-split-mode-sections.cs) | `Aspose`, `Document`, `DocumentBuilder` | Documentsplitcriteria object set split mode sections |
| [documentsplitcriteria-split-sections-then-each-part-net...](./documentsplitcriteria-split-sections-then-each-part-network-share-location.cs) | `Aspose`, `ArgumentException`, `Document` | Documentsplitcriteria split sections then each part network share location |
| [docx-source-document-document-class-before-splitting](./docx-source-document-document-class-before-splitting.cs) | `Document`, `Aspose`, `Source` | Docx source document document class before splitting |
| [ensure-split-parts-retain-complete-table-rows-when-orig...](./ensure-split-parts-retain-complete-table-rows-when-original-document-contains.cs) | `Aspose`, `Document`, `Words` | Ensure split parts retain complete table rows when original document contains |
| [execute-split-operation-multiple-documents-sequentially...](./execute-split-operation-multiple-documents-sequentially-storing-results-designated.cs) | `DocumentSplitCriteria`, `Aspose`, `ArgumentNullException` | Execute split operation multiple documents sequentially storing results desig... |
| [handle-exceptions-when-attempting-split-unsupported-mht...](./handle-exceptions-when-attempting-split-unsupported-mhtml-format.cs) | `Aspose`, `Document`, `DocumentBuilder` | Handle exceptions when attempting split unsupported mhtml format |
| [implement-documentpartsavingcallback-assign-filenames-b...](./implement-documentpartsavingcallback-assign-filenames-based-original-heading-text.cs) | `System`, `Document`, `Aspose` | Implement documentpartsavingcallback assign filenames based original heading... |
| [implement-documentpartsavingcallback-select-docx-even-p...](./implement-documentpartsavingcallback-select-docx-even-parts-pdf-odd-parts.cs) | `PartFormatCallback`, `Document`, `Aspose` | Implement documentpartsavingcallback select docx even parts pdf odd parts |
| [iterate-over-split-document-collection-each-part-docume...](./iterate-over-split-document-collection-each-part-documentpartsavingcallback.cs) | `DocumentSplitCriteria`, `Document`, `Aspose` | Iterate over split document collection each part documentpartsavingcallback |
| [merge-selected-split-documents-them-appenddocument-comb...](./merge-selected-split-documents-them-appenddocument-combined-file.cs) | `Document`, `Aspose`, `DocumentMerger` | Merge selected split documents them appenddocument combined file |
| [process-batch-docx-files-splitting-each-pages-pdfs-folder](./process-batch-docx-files-splitting-each-pages-pdfs-folder.cs) | `Document`, `Aspose`, `System` | Process batch docx files splitting each pages pdfs folder |
| [retain-original-page-orientation-each-split-part-preser...](./retain-original-page-orientation-each-split-part-preserving-landscape-pages.cs) | `Aspose`, `Document`, `PageExtractOptions` | Retain original page orientation each split part preserving landscape pages |
| [source-document-define-split-criteria-execute-split-ope...](./source-document-define-split-criteria-execute-split-operation-single-workflow.cs) | `Aspose`, `Document`, `HtmlSaveOptions` | Source document define split criteria execute split operation single workflow |
| [split-document-custom-page-ranges-like-1-3-5-7-export-e...](./split-document-custom-page-ranges-like-1-3-5-7-export-each-range-as-pdf.cs) | `Document`, `PageRange`, `PageSet` | Split document custom page ranges like 1 3 5 7 export each range as pdf |
| [split-epub-source-chapters-each-chapter-as-html-preserv...](./split-epub-source-chapters-each-chapter-as-html-preserving-layout.cs) | `Aspose`, `Document`, `HtmlSaveOptions` | Split epub source chapters each chapter as html preserving layout |
| [split-html-source-chapters-each-as-docx-while-preservin...](./split-html-source-chapters-each-as-docx-while-preserving-inline-styles.cs) | `Document`, `NodeImporter`, `System` | Split html source chapters each as docx while preserving inline styles |
| [split-large-word-file-50-page-chunks-each-chunk-as-pdf](./split-large-word-file-50-page-chunks-each-chunk-as-pdf.cs) | `Document`, `Aspose`, `WordSplitter` | Split large word file 50 page chunks each chunk as pdf |
| [split-parts-as-docx-files-preserving-original-formattin...](./split-parts-as-docx-files-preserving-original-formatting-page-orientation.cs) | `Document`, `Aspose` | Split parts as docx files preserving original formatting page orientation |
| [split-parts-as-pdf-files-while-preserving-original-docu...](./split-parts-as-pdf-files-while-preserving-original-document-styles-layout.cs) | `Aspose`, `Document`, `PageSplitCallback` | Split parts as pdf files while preserving original document styles layout |
| [validate-that-each-split-docx-file-maintains-original-h...](./validate-that-each-split-docx-file-maintains-original-header-footer-content-after.cs) | `Document`, `HeaderFooterType`, `Aspose` | Validate that each split docx file maintains original header footer content a... |

## Category Statistics
- Total examples: 29

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for split-document patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082619`
<!-- AUTOGENERATED:END -->
