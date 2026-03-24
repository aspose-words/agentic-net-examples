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

- `using System;` (23/23 files) ← category-specific
- `using Aspose.Words;` (23/23 files)
- `using Aspose.Words.Saving;` (20/23 files)
- `using System.IO;` (16/23 files)
- `using System.Collections.Generic;` (3/23 files)
- `using Aspose.Words.Drawing;` (2/23 files)
- `using System.Diagnostics;` (1/23 files)
- `using Aspose.Words.Tables;` (1/23 files)
- `using System.Linq;` (1/23 files)
- `using System.Text;` (1/23 files)
- `using Aspose.Words.Loading;` (1/23 files)

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
| [after-splitting-open-each-output-document-programmatica...](./after-splitting-open-each-output-document-programmatically-verify-headers-footers.cs) | `HeaderFooterType`, `Document`, `DocumentBuilder` | After splitting open each output document programmatically verify headers foo... |
| [combine-heading-section-flags-documentsplitcriteria-spl...](./combine-heading-section-flags-documentsplitcriteria-split-both-structures.cs) | `ParagraphFormat`, `StyleIdentifier`, `Document` | Combine heading section flags documentsplitcriteria split both structures |
| [combine-page-heading-flags-documentsplitcriteria-start-...](./combine-page-heading-flags-documentsplitcriteria-start-each-part-new-page.cs) | `Document`, `ParagraphFormat`, `StyleIdentifier` | Combine page heading flags documentsplitcriteria start each part new page |
| [combine-page-range-heading-criteria-produce-parts-that-...](./combine-page-range-heading-criteria-produce-parts-that-start-at-each-heading-new-page.cs) | `StyleIdentifier`, `ParagraphFormat`, `Document` | Combine page range heading criteria produce parts that start at each heading... |
| [documentsplitcriteria-enumeration-split-sections-then-m...](./documentsplitcriteria-enumeration-split-sections-then-merge-selected-parts.cs) | `Document`, `DocumentSplitCriteria`, `Section` | Documentsplitcriteria enumeration split sections then merge selected parts |
| [documentsplitcriteria-object-set-split-mode-custom-page...](./documentsplitcriteria-object-set-split-mode-custom-page-ranges.cs) | `ParagraphFormat`, `StyleIdentifier`, `Document` | Documentsplitcriteria object set split mode custom page ranges |
| [documentsplitcriteria-object-set-split-mode-headings](./documentsplitcriteria-object-set-split-mode-headings.cs) | `ParagraphFormat`, `Document`, `DocumentBuilder` | Documentsplitcriteria object set split mode headings |
| [documentsplitcriteria-object-set-split-mode-sections](./documentsplitcriteria-object-set-split-mode-sections.cs) | `Document`, `DocumentBuilder`, `BreakType` | Documentsplitcriteria object set split mode sections |
| [documentsplitcriteria-split-sections-then-each-part-net...](./documentsplitcriteria-split-sections-then-each-part-network-share-location.cs) | `Document`, `ArgumentException`, `DocumentBuilder` | Documentsplitcriteria split sections then each part network share location |
| [docx-source-document-document-class-before-splitting](./docx-source-document-document-class-before-splitting.cs) | `Document`, `DocumentBuilder`, `Sections` | Docx source document document class before splitting |
| [execute-split-operation-multiple-documents-sequentially...](./execute-split-operation-multiple-documents-sequentially-storing-results-designated.cs) | `DocumentSplitCriteria`, `Document`, `ArgumentNullException` | Execute split operation multiple documents sequentially storing results desig... |
| [handle-exceptions-when-attempting-split-unsupported-mht...](./handle-exceptions-when-attempting-split-unsupported-mhtml-format.cs) | `Document`, `DocumentBuilder`, `HtmlSaveOptions` | Handle exceptions when attempting split unsupported mhtml format |
| [implement-documentpartsavingcallback-assign-filenames-b...](./implement-documentpartsavingcallback-assign-filenames-based-original-heading-text.cs) | `Document`, `HeadingBasedDocumentPartRename`, `HtmlSaveOptions` | Implement documentpartsavingcallback assign filenames based original heading... |
| [iterate-over-split-document-collection-each-part-docume...](./iterate-over-split-document-collection-each-part-documentpartsavingcallback.cs) | `DocumentSplitCriteria`, `Document`, `DocumentBuilder` | Iterate over split document collection each part documentpartsavingcallback |
| [merge-selected-split-documents-them-appenddocument-comb...](./merge-selected-split-documents-them-appenddocument-combined-file.cs) | `Document`, `DocumentBuilder`, `DocumentMerger` | Merge selected split documents them appenddocument combined file |
| [process-batch-docx-files-splitting-each-pages-pdfs-folder](./process-batch-docx-files-splitting-each-pages-pdfs-folder.cs) | `Document`, `DocumentBuilder`, `AppContext` | Process batch docx files splitting each pages pdfs folder |
| [retain-original-page-orientation-each-split-part-preser...](./retain-original-page-orientation-each-split-part-preserving-landscape-pages.cs) | `Document`, `Orientation`, `PageSetup` | Retain original page orientation each split part preserving landscape pages |
| [source-document-define-split-criteria-execute-split-ope...](./source-document-define-split-criteria-execute-split-operation-single-workflow.cs) | `Document`, `Section`, `Body` | Source document define split criteria execute split operation single workflow |
| [split-document-custom-page-ranges-like-1-3-5-7-export-e...](./split-document-custom-page-ranges-like-1-3-5-7-export-each-range-as-pdf.cs) | `PageSet`, `Document`, `DocumentBuilder` | Split document custom page ranges like 1 3 5 7 export each range as pdf |
| [split-epub-source-chapters-each-chapter-as-html-preserv...](./split-epub-source-chapters-each-chapter-as-html-preserving-layout.cs) | `ParagraphFormat`, `StyleIdentifier`, `Document` | Split epub source chapters each chapter as html preserving layout |
| [split-html-source-chapters-each-as-docx-while-preservin...](./split-html-source-chapters-each-as-docx-while-preserving-inline-styles.cs) | `Document`, `NodeImporter`, `ImportFormatMode` | Split html source chapters each as docx while preserving inline styles |
| [split-large-word-file-50-page-chunks-each-chunk-as-pdf](./split-large-word-file-50-page-chunks-each-chunk-as-pdf.cs) | `Document`, `WordSplitter`, `Part_1` | Split large word file 50 page chunks each chunk as pdf |
| [split-parts-as-pdf-files-while-preserving-original-docu...](./split-parts-as-pdf-files-while-preserving-original-document-styles-layout.cs) | `Document`, `DocumentBuilder`, `BreakType` | Split parts as pdf files while preserving original document styles layout |

## Category Statistics
- Total examples: 23

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for split-document patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
