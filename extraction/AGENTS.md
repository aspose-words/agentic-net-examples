---
name: extraction
description: Verified C# examples for extraction scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Extraction

## Purpose

This folder is a **live, curated example set** for extraction scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free extraction of text, tables, images, bookmarks, comments, fields, metadata, and related document content using direct Aspose.Words APIs.

## Non-negotiable conventions

- Use native Aspose.Words extraction APIs directly.
- Bootstrap local source documents, files, images, streams, or folders whenever the task implies an existing source.
- Enumerate actual document nodes and validate their types before extracting content.
- Use `Aspose.Words.Tables` for `Table`, `Row`, and `Cell` when structured extraction is required.
- Use `Newtonsoft.Json` for JSON serialization tasks and `Aspose.Drawing` instead of `System.Drawing` when drawing-related types are needed.
- Guard maybe-null values to avoid nullable-reference warnings such as `CS8600`, `CS8602`, and `CS8604`.

## Recommended workflow selection

- **Text / range extraction workflow**: 11 examples
- **Table / structured extraction workflow**: 6 examples
- **Image / shape extraction workflow**: 3 examples
- **Targeted node extraction workflow**: 9 examples
- **Input-bootstrap workflow**: 1 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. Source documents, files, images, streams, or folders must be bootstrapped locally whenever the task implies an existing input.
3. Extracted content must come from real document nodes, fields, bookmarks, comments, tables, images, or metadata objects.
4. Requested report or export files must actually be written.
5. Drawing-related types must use `Aspose.Drawing` and not `System.Drawing`.

## File-to-task reference

- `load-a-docx-file-extract-content-between-two-paragraphs-and-save-the-result-as-a-new-docx.cs`
  - Task: Load a DOCX file, extract content between two paragraphs, and save the result as a new DOCX.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: verified
- `load-a-docm-file-extract-content-between-a-macro-enabled-field-and-a-paragraph-and-save-as.cs`
  - Task: Load a DOCM file, extract content between a macro-enabled field and a paragraph, and save as DOCX.
  - Workflow: targeted-node-extraction
  - Outputs: docx, doc
  - Selected engine: verified
- `identify-a-start-run-node-and-an-end-bookmark-node-then-extract-the-intervening-nodes-into.cs`
  - Task: Identify a start run node and an end bookmark node, then extract the intervening nodes into a document.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: verified
- `programmatically-determine-start-and-end-nodes-based-on-paragraph-styles-then-extract-the.cs`
  - Task: Programmatically determine start and end nodes based on paragraph styles, then extract the styled content segment.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-a-mixed-node-range-that-starts-with-a-table-cell-and-ends-with-a-paragraph-maintai.cs`
  - Task: Extract a mixed node range that starts with a table cell and ends with a paragraph, maintaining layout.
  - Workflow: table-structured-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-a-range-of-nodes-that-includes-tables-images-and-fields-preserving-original-hierar.cs`
  - Task: Extract a range of nodes that includes tables, images, and fields, preserving original hierarchy in the output.
  - Workflow: table-structured-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-content-between-a-run-node-and-the-next-bookmark-then-convert-the-extracted-segmen.cs`
  - Task: Extract content between a run node and the next bookmark, then convert the extracted segment to HTML format.
  - Workflow: targeted-node-extraction
  - Outputs: html
  - Selected engine: verified
- `extract-content-between-a-run-node-and-the-following-table-then-convert-the-extracted-port.cs`
  - Task: Extract content between a run node and the following table, then convert the extracted portion to XPS format.
  - Workflow: table-structured-extraction
  - Outputs: xps
  - Selected engine: verified
- `use-the-extraction-api-to-copy-content-between-two-headings-and-insert-it-into-a-template.cs`
  - Task: Use the extraction API to copy content between two headings and insert it into a template document.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: verified
- `use-documentbuilder-to-prepend-extracted-node-collection-to-the-beginning-of-a-new-documen.cs`
  - Task: Use DocumentBuilder to prepend extracted node collection to the beginning of a new document before saving.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: verified
- `use-documentbuilder-to-insert-extracted-node-collection-into-a-new-document-at-a-custom-bo.cs`
  - Task: Use DocumentBuilder to insert extracted node collection into a new document at a custom bookmark location.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: verified
- `duplicate-extracted-content-between-a-table-and-a-field-node-within-the-original-document.cs`
  - Task: Duplicate extracted content between a table and a field node within the original document without altering formatting.
  - Workflow: table-structured-extraction
  - Outputs: docx
  - Selected engine: verified
- `save-extracted-content-as-a-docx-file-while-preserving-embedded-fields-and-their-evaluatio.cs`
  - Task: Save extracted content as a DOCX file while preserving embedded fields and their evaluation results.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: verified
- `batch-process-multiple-word-files-extracting-content-between-specified-nodes-and-saving-ea.cs`
  - Task: Batch process multiple Word files, extracting content between specified nodes and saving each extraction as an individual PDF.
  - Workflow: input-bootstrap
  - Outputs: pdf
  - Selected engine: verified
- `batch-extract-images-from-shape-nodes-in-documents-and-generate-a-csv-manifest-listing-ima.cs`
  - Task: Batch extract images from shape nodes in documents and generate a CSV manifest listing image names and sources.
  - Workflow: table-structured-extraction
  - Outputs: csv
  - Selected engine: verified
- `extract-all-images-from-shape-nodes-across-a-document-collection-and-compile-them-into-a-s.cs`
  - Task: Extract all images from shape nodes across a document collection and compile them into a single ZIP archive.
  - Workflow: image-shape-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-a-document-segment-that-includes-nested-tables-and-ensure-nested-structures-are-re.cs`
  - Task: Extract a document segment that includes nested tables and ensure nested structures are retained in the new file.
  - Workflow: table-structured-extraction
  - Outputs: docx
  - Selected engine: verified
- `create-a-reusable-extraction-utility-that-accepts-node-identifiers-and-returns-a-document.cs`
  - Task: Create a reusable extraction utility that accepts node identifiers and returns a Document containing the selected content.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: verified
- `implement-error-handling-for-cases-where-the-start-node-appears-after-the-end-node-during.cs`
  - Task: Implement error handling for cases where the start node appears after the end node during extraction.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: verified
- `implement-a-custom-node-filter-to-exclude-comments-while-extracting-content-between-two-pa.cs`
  - Task: Implement a custom node filter to exclude comments while extracting content between two paragraphs.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-content-between-two-bookmark-nodes-and-replace-the-original-range-with-a-placehold.cs`
  - Task: Extract content between two bookmark nodes and replace the original range with a placeholder paragraph.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-content-between-a-paragraph-and-a-comment-node-then-log-the-extracted-text-to-a-mo.cs`
  - Task: Extract content between a paragraph and a comment node, then log the extracted text to a monitoring system.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: verified
- `automate-extraction-of-footnote-content-between-specified-nodes-and-export-each-footnote-a.cs`
  - Task: Automate extraction of footnote content between specified nodes and export each footnote as a separate text file.
  - Workflow: targeted-node-extraction
  - Outputs: txt
  - Selected engine: verified
- `create-a-command-line-tool-that-accepts-start-and-end-node-ids-and-outputs-the-extracted-s.cs`
  - Task: Create a command‑line tool that accepts start and end node IDs and outputs the extracted segment as PDF.
  - Workflow: text-range-extraction
  - Outputs: pdf
  - Selected engine: verified
- `create-a-unit-test-that-verifies-extraction-of-content-between-two-specific-paragraphs-ret.cs`
  - Task: Create a unit test that verifies extraction of content between two specific paragraphs retains original text styling.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: verified
- `develop-a-macro-that-calls-the-extraction-api-to-copy-selected-content-into-the-clipboard.cs`
  - Task: Develop a macro that calls the extraction API to copy selected content into the clipboard for pasting elsewhere.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-a-range-that-starts-inside-a-shape-s-image-and-ends-at-a-field-preserving-both-ele.cs`
  - Task: Extract a range that starts inside a shape's image and ends at a field, preserving both elements.
  - Workflow: image-shape-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-content-between-two-nodes-in-a-document-then-encrypt-the-resulting-file-using-a-pa.cs`
  - Task: Extract content between two nodes in a document, then encrypt the resulting file using a password.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: verified
- `implement-parallel-processing-to-extract-node-ranges-from-multiple-documents-simultaneousl.cs`
  - Task: Implement parallel processing to extract node ranges from multiple documents simultaneously, improving performance.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: verified
- `extract-images-from-shape-nodes-and-embed-them-directly-into-a-new-docx-document.cs`
  - Task: Extract images from shape nodes and embed them directly into a new DOCX document.
  - Workflow: image-shape-extraction
  - Outputs: docx
  - Selected engine: verified

## Common failure patterns and preferred agent fixes

- **Invalid node insertion in destination document**
  - Symptom: Runtime failures such as 'Cannot insert a node of this type at this location'.
  - Preferred fix: Rebuild a valid destination document structure explicitly and insert cloned inline or block nodes only into supported parents.

- **Table namespace assumptions**
  - Symptom: Compile errors because Table, Row, or Cell were used without Aspose.Words.Tables.
  - Preferred fix: Import Aspose.Words.Tables or use fully qualified table node type names.

- **Weak bookmark boundary logic**
  - Symptom: Bookmark-bounded extraction or replacement fails because the wrong nodes are used or placeholder insertion is invalid.
  - Preferred fix: Use real BookmarkStart and BookmarkEnd boundaries and insert a valid placeholder Paragraph into a supported block container.

- **Footnote-specific API issues**
  - Symptom: Footnote export or enumeration fails because footnote APIs or namespaces are used incorrectly.
  - Preferred fix: Use Aspose.Words.Notes.Footnote and Aspose.Words.Notes.FootnoteType explicitly and enumerate actual footnote nodes.

- **Font or drawing ambiguity**
  - Symptom: Compile errors due to System.Drawing usage or ambiguous Font references.
  - Preferred fix: Use Aspose.Drawing only and prefer explicit Aspose.Drawing type names when ambiguity is possible.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`
- Additional package: `Aspose.Drawing.Common`
- Additional package: `Newtonsoft.Json`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.Drawing.Common
dotnet add package Newtonsoft.Json
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\extraction\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer direct Aspose.Words extraction APIs over speculative shortcuts.
