---
name: extraction
description: C# examples for extraction using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - extraction

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **extraction** category.
This folder contains standalone C# examples for extraction operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **extraction**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using System;` (30/30 files) ← category-specific
- `using Aspose.Words;` (30/30 files)
- `using Aspose.Words.Saving;` (10/30 files)
- `using System.IO;` (9/30 files)
- `using System.Collections.Generic;` (9/30 files)
- `using System.Linq;` (8/30 files)
- `using Aspose.Words.Drawing;` (7/30 files)
- `using Aspose.Words.Tables;` (6/30 files)
- `using System.Text;` (4/30 files)
- `using Aspose.Words.Fields;` (4/30 files)
- `using Aspose.Words.Markup;` (2/30 files)
- `using Aspose.Words.Notes;` (1/30 files)

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
| [automate-extraction-footnote-content-between-specified-...](./automate-extraction-footnote-content-between-specified-nodes-export-each-footnote-as.cs) | `Document`, `Aspose`, `System` | Automate extraction footnote content between specified nodes export each foot... |
| [batch-extract-images-shape-nodes-documents-csv-manifest...](./batch-extract-images-shape-nodes-documents-csv-manifest-listing-image-names-sources.cs) | `Aspose`, `ImageData`, `System` | Batch extract images shape nodes documents csv manifest listing image names s... |
| [batch-process-multiple-word-files-extracting-content-be...](./batch-process-multiple-word-files-extracting-content-between-specified-nodes-each.cs) | `Document`, `NodeImporter`, `Aspose` | Batch process multiple word files extracting content between specified nodes... |
| [command-line-tool-that-accepts-start-end-node-ids-outpu...](./command-line-tool-that-accepts-start-end-node-ids-outputs-extracted-segment-as-pdf.cs) | `Document`, `Aspose`, `ExtractSegmentTool` | Command line tool that accepts start end node ids outputs extracted segment a... |
| [develop-macro-that-calls-extraction-api-copy-selected-c...](./develop-macro-that-calls-extraction-api-copy-selected-content-clipboard-pasting.cs) | `Aspose`, `Document`, `DocumentBuilder` | Develop macro that calls extraction api copy selected content clipboard pasting |
| [docm-file-extract-content-between-macro-enabled-field-p...](./docm-file-extract-content-between-macro-enabled-field-paragraph-as-docx.cs) | `Document`, `NodeImporter`, `System` | Docm file extract content between macro enabled field paragraph as docx |
| [documentbuilder-insert-extracted-node-collection-new-do...](./documentbuilder-insert-extracted-node-collection-new-document-at-custom-bookmark.cs) | `NodeType`, `Document`, `DocumentBuilder` | Documentbuilder insert extracted node collection new document at custom bookmark |
| [documentbuilder-prepend-extracted-node-collection-begin...](./documentbuilder-prepend-extracted-node-collection-beginning-new-document-before.cs) | `Document`, `DocumentBuilder`, `NodeImporter` | Documentbuilder prepend extracted node collection beginning new document before |
| [docx-file-extract-content-between-two-paragraphs-result...](./docx-file-extract-content-between-two-paragraphs-result-as-new-docx.cs) | `Document`, `NodeImporter`, `InvalidOperationException` | Docx file extract content between two paragraphs result as new docx |
| [duplicate-extracted-content-between-table-field-node-wi...](./duplicate-extracted-content-between-table-field-node-within-original-document-without.cs) | `Aspose`, `Document`, `NodeImporter` | Duplicate extracted content between table field node within original document... |
| [extract-all-images-shape-nodes-across-document-collecti...](./extract-all-images-shape-nodes-across-document-collection-compile-them-single-zip.cs) | `System`, `Aspose`, `Words` | Extract all images shape nodes across document collection compile them single... |
| [extract-content-between-paragraph-comment-node-then-log...](./extract-content-between-paragraph-comment-node-then-log-extracted-text-monitoring.cs) | `Document`, `StringBuilder`, `Aspose` | Extract content between paragraph comment node then log extracted text monito... |
| [extract-content-between-run-node-following-table-then-c...](./extract-content-between-run-node-following-table-then-convert-extracted-portion-xps.cs) | `Document`, `Body`, `XpsSaveOptions` | Extract content between run node following table then convert extracted porti... |
| [extract-content-between-run-node-next-bookmark-then-con...](./extract-content-between-run-node-next-bookmark-then-convert-extracted-segment-html.cs) | `Document`, `FirstSection`, `Body` | Extract content between run node next bookmark then convert extracted segment... |
| [extract-content-between-two-bookmark-nodes-replace-orig...](./extract-content-between-two-bookmark-nodes-replace-original-range-placeholder-paragraph.cs) | `Document`, `NodeImporter`, `Paragraph` | Extract content between two bookmark nodes replace original range placeholder... |
| [extract-content-between-two-nodes-document-then-encrypt...](./extract-content-between-two-nodes-document-then-encrypt-resulting-file-password.cs) | `Document`, `Aspose`, `NodeImporter` | Extract content between two nodes document then encrypt resulting file password |
| [extract-document-segment-that-includes-nested-tables-en...](./extract-document-segment-that-includes-nested-tables-ensure-nested-structures-are.cs) | `Document`, `DocumentBuilder`, `Aspose` | Extract document segment that includes nested tables ensure nested structures... |
| [extract-images-shape-nodes-embed-them-directly-new-docx...](./extract-images-shape-nodes-embed-them-directly-new-docx-document.cs) | `Document`, `DocumentBuilder`, `System` | Extract images shape nodes embed them directly new docx document |
| [extract-mixed-node-range-that-starts-table-cell-ends-pa...](./extract-mixed-node-range-that-starts-table-cell-ends-paragraph-maintaining-layout.cs) | `Document`, `Section`, `Body` | Extract mixed node range that starts table cell ends paragraph maintaining la... |
| [extract-range-nodes-that-includes-tables-images-fields-...](./extract-range-nodes-that-includes-tables-images-fields-preserving-original-hierarchy.cs) | `NodeType`, `Document`, `Aspose` | Extract range nodes that includes tables images fields preserving original hi... |
| [extract-range-that-starts-inside-shape-s-image-ends-at-...](./extract-range-that-starts-inside-shape-s-image-ends-at-field-preserving-both-elements.cs) | `Document`, `Aspose`, `DocumentBuilder` | Extract range that starts inside shape s image ends at field preserving both... |
| [extracted-content-as-docx-file-while-preserving-embedde...](./extracted-content-as-docx-file-while-preserving-embedded-fields-their-evaluation.cs) | `Aspose`, `Document`, `OoxmlSaveOptions` | Extracted content as docx file while preserving embedded fields their evaluation |
| [extraction-api-copy-content-between-two-headings-insert...](./extraction-api-copy-content-between-two-headings-insert-it-template-document.cs) | `Document`, `NodeImporter`, `DocumentBuilder` | Extraction api copy content between two headings insert it template document |
| [identify-start-run-node-end-bookmark-node-then-extract-...](./identify-start-run-node-end-bookmark-node-then-extract-intervening-nodes-document.cs) | `Document`, `InvalidOperationException`, `NodeImporter` | Identify start run node end bookmark node then extract intervening nodes docu... |
| [implement-custom-node-filter-exclude-comments-while-ext...](./implement-custom-node-filter-exclude-comments-while-extracting-content-between-two.cs) | `VisitorAction`, `Aspose`, `StringBuilder` | Implement custom node filter exclude comments while extracting content betwee... |
| [implement-error-handling-cases-where-start-node-appears...](./implement-error-handling-cases-where-start-node-appears-after-end-node-during.cs) | `Document`, `InvalidOperationException`, `LayoutCollector` | Implement error handling cases where start node appears after end node during |
| [implement-parallel-processing-extract-node-ranges-multi...](./implement-parallel-processing-extract-node-ranges-multiple-documents-simultaneously.cs) | `System`, `Document`, `Parallel` | Implement parallel processing extract node ranges multiple documents simultan... |
| [programmatically-determine-start-end-nodes-based-paragr...](./programmatically-determine-start-end-nodes-based-paragraph-styles-then-extract-styled.cs) | `Document`, `StringBuilder`, `System` | Programmatically determine start end nodes based paragraph styles then extrac... |
| [reusable-extraction-utility-that-accepts-node-identifie...](./reusable-extraction-utility-that-accepts-node-identifiers-returns-document-containing.cs) | `Document`, `ArgumentNullException`, `ArgumentException` | Reusable extraction utility that accepts node identifiers returns document co... |
| [unit-test-that-verifies-extraction-content-between-two-...](./unit-test-that-verifies-extraction-content-between-two-specific-paragraphs-retains.cs) | `Font`, `ParagraphFormat`, `Document` | Unit test that verifies extraction content between two specific paragraphs re... |

## Category Statistics
- Total examples: 30

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for extraction patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082340`
<!-- AUTOGENERATED:END -->
