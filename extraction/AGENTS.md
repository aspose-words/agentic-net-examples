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

- `using System;` (27/27 files) ← category-specific
- `using Aspose.Words;` (27/27 files)
- `using System.IO;` (9/27 files)
- `using System.Collections.Generic;` (9/27 files)
- `using System.Linq;` (8/27 files)
- `using Aspose.Words.Saving;` (8/27 files)
- `using Aspose.Words.Tables;` (8/27 files)
- `using Aspose.Words.Drawing;` (6/27 files)
- `using System.Text;` (4/27 files)
- `using Aspose.Words.Fields;` (4/27 files)
- `using Aspose.Words.Notes;` (1/27 files)
- `using Aspose.Words.Vba;` (1/27 files)

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
| [automate-extraction-footnote-content-between-specified-...](./automate-extraction-footnote-content-between-specified-nodes-export-each-footnote-as.cs) | `Document`, `Input`, `NodeType` | Automate extraction footnote content between specified nodes export each foot... |
| [batch-extract-images-shape-nodes-documents-csv-manifest...](./batch-extract-images-shape-nodes-documents-csv-manifest-listing-image-names-sources.cs) | `ImageData`, `Document`, `DocumentBuilder` | Batch extract images shape nodes documents csv manifest listing image names s... |
| [command-line-tool-that-accepts-start-end-node-ids-outpu...](./command-line-tool-that-accepts-start-end-node-ids-outputs-extracted-segment-as-pdf.cs) | `Document`, `ExtractSegmentTool`, `PDF` | Command line tool that accepts start end node ids outputs extracted segment a... |
| [develop-macro-that-calls-extraction-api-copy-selected-c...](./develop-macro-that-calls-extraction-api-copy-selected-content-clipboard-pasting.cs) | `Document`, `DocumentBuilder`, `VbaProject` | Develop macro that calls extraction api copy selected content clipboard pasting |
| [docm-file-extract-content-between-macro-enabled-field-p...](./docm-file-extract-content-between-macro-enabled-field-paragraph-as-docx.cs) | `Document`, `DocumentBuilder`, `NodeImporter` | Docm file extract content between macro enabled field paragraph as docx |
| [documentbuilder-insert-extracted-node-collection-new-do...](./documentbuilder-insert-extracted-node-collection-new-document-at-custom-bookmark.cs) | `DocumentBuilder`, `Document`, `NodeType` | Documentbuilder insert extracted node collection new document at custom bookmark |
| [documentbuilder-prepend-extracted-node-collection-begin...](./documentbuilder-prepend-extracted-node-collection-beginning-new-document-before.cs) | `Document`, `DocumentBuilder`, `NodeImporter` | Documentbuilder prepend extracted node collection beginning new document before |
| [docx-file-extract-content-between-two-paragraphs-result...](./docx-file-extract-content-between-two-paragraphs-result-as-new-docx.cs) | `Document`, `NodeImporter`, `DocumentBuilder` | Docx file extract content between two paragraphs result as new docx |
| [duplicate-extracted-content-between-table-field-node-wi...](./duplicate-extracted-content-between-table-field-node-within-original-document-without.cs) | `Document`, `NodeImporter`, `NodeType` | Duplicate extracted content between table field node within original document... |
| [extract-all-images-shape-nodes-across-document-collecti...](./extract-all-images-shape-nodes-across-document-collection-compile-them-single-zip.cs) | `Document`, `ZipArchive`, `DocumentBuilder` | Extract all images shape nodes across document collection compile them single... |
| [extract-content-between-paragraph-comment-node-then-log...](./extract-content-between-paragraph-comment-node-then-log-extracted-text-monitoring.cs) | `Document`, `StringBuilder`, `NodeType` | Extract content between paragraph comment node then log extracted text monito... |
| [extract-content-between-run-node-following-table-then-c...](./extract-content-between-run-node-following-table-then-convert-extracted-portion-xps.cs) | `Document`, `DocumentBuilder`, `Section` | Extract content between run node following table then convert extracted porti... |
| [extract-content-between-run-node-next-bookmark-then-con...](./extract-content-between-run-node-next-bookmark-then-convert-extracted-segment-html.cs) | `Document`, `FirstSection`, `Body` | Extract content between run node next bookmark then convert extracted segment... |
| [extract-content-between-two-bookmark-nodes-replace-orig...](./extract-content-between-two-bookmark-nodes-replace-original-range-placeholder-paragraph.cs) | `Paragraph`, `Document`, `Run` | Extract content between two bookmark nodes replace original range placeholder... |
| [extract-content-between-two-nodes-document-then-encrypt...](./extract-content-between-two-nodes-document-then-encrypt-resulting-file-password.cs) | `Document`, `NodeImporter`, `DocumentBuilder` | Extract content between two nodes document then encrypt resulting file password |
| [extract-document-segment-that-includes-nested-tables-en...](./extract-document-segment-that-includes-nested-tables-ensure-nested-structures-are.cs) | `Document`, `DocumentBuilder`, `NodeImporter` | Extract document segment that includes nested tables ensure nested structures... |
| [extract-images-shape-nodes-embed-them-directly-new-docx...](./extract-images-shape-nodes-embed-them-directly-new-docx-document.cs) | `Document`, `DocumentBuilder`, `NodeType` | Extract images shape nodes embed them directly new docx document |
| [extract-mixed-node-range-that-starts-table-cell-ends-pa...](./extract-mixed-node-range-that-starts-table-cell-ends-paragraph-maintaining-layout.cs) | `Document`, `Section`, `Body` | Extract mixed node range that starts table cell ends paragraph maintaining la... |
| [extract-range-nodes-that-includes-tables-images-fields-...](./extract-range-nodes-that-includes-tables-images-fields-preserving-original-hierarchy.cs) | `NodeType`, `Document`, `NodeImporter` | Extract range nodes that includes tables images fields preserving original hi... |
| [extract-range-that-starts-inside-shape-s-image-ends-at-...](./extract-range-that-starts-inside-shape-s-image-ends-at-field-preserving-both-elements.cs) | `Document`, `DocumentBuilder`, `NodeImporter` | Extract range that starts inside shape s image ends at field preserving both... |
| [extracted-content-as-docx-file-while-preserving-embedde...](./extracted-content-as-docx-file-while-preserving-embedded-fields-their-evaluation.cs) | `Document`, `DocumentBuilder`, `OoxmlSaveOptions` | Extracted content as docx file while preserving embedded fields their evaluation |
| [extraction-api-copy-content-between-two-headings-insert...](./extraction-api-copy-content-between-two-headings-insert-it-template-document.cs) | `DocumentBuilder`, `Document`, `NodeImporter` | Extraction api copy content between two headings insert it template document |
| [identify-start-run-node-end-bookmark-node-then-extract-...](./identify-start-run-node-end-bookmark-node-then-extract-intervening-nodes-document.cs) | `Document`, `DocumentBuilder`, `InvalidOperationException` | Identify start run node end bookmark node then extract intervening nodes docu... |
| [implement-error-handling-cases-where-start-node-appears...](./implement-error-handling-cases-where-start-node-appears-after-end-node-during.cs) | `Document`, `DocumentBuilder`, `LayoutCollector` | Implement error handling cases where start node appears after end node during |
| [implement-parallel-processing-extract-node-ranges-multi...](./implement-parallel-processing-extract-node-ranges-multiple-documents-simultaneously.cs) | `Document`, `BreakType`, `StringBuilder` | Implement parallel processing extract node ranges multiple documents simultan... |
| [programmatically-determine-start-end-nodes-based-paragr...](./programmatically-determine-start-end-nodes-based-paragraph-styles-then-extract-styled.cs) | `Document`, `StringBuilder`, `NodeType` | Programmatically determine start end nodes based paragraph styles then extrac... |
| [reusable-extraction-utility-that-accepts-node-identifie...](./reusable-extraction-utility-that-accepts-node-identifiers-returns-document-containing.cs) | `Document`, `ArgumentNullException`, `ArgumentException` | Reusable extraction utility that accepts node identifiers returns document co... |

## Category Statistics
- Total examples: 27

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for extraction patterns.


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
Copy-Item ..\extraction\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

### Notes for Agents and Developers

- Treat every `.cs` file in `extraction/` as a full console program, not a snippet.
- Run one file at a time by copying it to `Program.cs`.
- If a sample needs input documents, images, fonts, or data files, place them in the temporary project directory before running.
- See the root [AGENTS.md](../AGENTS.md) for repository-wide prerequisites, project file template, and testing guidance.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
