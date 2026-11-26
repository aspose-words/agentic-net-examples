---
name: extraction
description: Verified C# examples for extraction scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Extraction

## Purpose

Provide an AI coding assistant with concise guidance for building reliable node-range extraction utilities that respect Aspose.Words document hierarchies, avoid cross-document node insertion errors, and produce correctly formatted output files.

## Non-negotiable conventions

- Do not reference System.Drawing; use Aspose.Drawing for all drawing operations.
- Do not assume external files exist; always create sample documents or streams within the example.
- Do not fabricate helper APIs; rely solely on documented Aspose.Words types such as DocumentBuilder, NodeImporter, and the classes in Aspose.Words.Tables.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Text Range Extraction workflow**: 11 examples
- **Targeted Node Extraction workflow**: 9 examples
- **Table Structured Extraction workflow**: 6 examples
- **Input Bootstrap workflow**: 1 examples
- **Image Shape Extraction workflow**: 3 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Validate that the extracted document contains the expected number of nodes and that the output file exists on disk.

## File-to-task reference

- `load-a-docx-file-extract-content-between-two-paragraphs-and-save-the-result-as-a-new-docx.cs`
  - Task: Load a DOCX file, extract content between two paragraphs, and save the result as a new DOCX.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: mcp
- `load-a-docm-file-extract-content-between-a-macro-enabled-field-and-a-paragraph-and-save-as.cs`
  - Task: Load a DOCM file, extract content between a macro-enabled field and a paragraph, and save as DOCX.
  - Workflow: targeted-node-extraction
  - Outputs: docx, doc
  - Selected engine: mcp
- `identify-a-start-run-node-and-an-end-bookmark-node-then-extract-the-intervening-nodes-into.cs`
  - Task: Identify a start run node and an end bookmark node, then extract the intervening nodes into a document.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-determine-start-and-end-nodes-based-on-paragraph-styles-then-extract-the.cs`
  - Task: Programmatically determine start and end nodes based on paragraph styles, then extract the styled content segment.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: mcp
- `extract-a-mixed-node-range-that-starts-with-a-table-cell-and-ends-with-a-paragraph-maintai.cs`
  - Task: Extract a mixed node range that starts with a table cell and ends with a paragraph, maintaining layout.
  - Workflow: table-structured-extraction
  - Outputs: docx
  - Selected engine: mcp
- `extract-a-range-of-nodes-that-includes-tables-images-and-fields-preserving-original-hierar.cs`
  - Task: Extract a range of nodes that includes tables, images, and fields, preserving original hierarchy in the output.
  - Workflow: table-structured-extraction
  - Outputs: docx
  - Selected engine: mcp
- `extract-content-between-a-run-node-and-the-next-bookmark-then-convert-the-extracted-segmen.cs`
  - Task: Extract content between a run node and the next bookmark, then convert the extracted segment to HTML format.
  - Workflow: targeted-node-extraction
  - Outputs: html
  - Selected engine: mcp
- `extract-content-between-a-run-node-and-the-following-table-then-convert-the-extracted-port.cs`
  - Task: Extract content between a run node and the following table, then convert the extracted portion to XPS format.
  - Workflow: table-structured-extraction
  - Outputs: xps
  - Selected engine: existing_repo
- `use-the-extraction-api-to-copy-content-between-two-headings-and-insert-it-into-a-template.cs`
  - Task: Use the extraction API to copy content between two headings and insert it into a template document.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: mcp
- `use-documentbuilder-to-prepend-extracted-node-collection-to-the-beginning-of-a-new-documen.cs`
  - Task: Use DocumentBuilder to prepend extracted node collection to the beginning of a new document before saving.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: mcp
- `use-documentbuilder-to-insert-extracted-node-collection-into-a-new-document-at-a-custom-bo.cs`
  - Task: Use DocumentBuilder to insert extracted node collection into a new document at a custom bookmark location.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: existing_repo
- `duplicate-extracted-content-between-a-table-and-a-field-node-within-the-original-document.cs`
  - Task: Duplicate extracted content between a table and a field node within the original document without altering formatting.
  - Workflow: table-structured-extraction
  - Outputs: docx
  - Selected engine: mcp
- `save-extracted-content-as-a-docx-file-while-preserving-embedded-fields-and-their-evaluatio.cs`
  - Task: Save extracted content as a DOCX file while preserving embedded fields and their evaluation results.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: existing_repo
- `batch-process-multiple-word-files-extracting-content-between-specified-nodes-and-saving-ea.cs`
  - Task: Batch process multiple Word files, extracting content between specified nodes and saving each extraction as an individual PDF.
  - Workflow: input-bootstrap
  - Outputs: pdf
  - Selected engine: mcp
- `batch-extract-images-from-shape-nodes-in-documents-and-generate-a-csv-manifest-listing-ima.cs`
  - Task: Batch extract images from shape nodes in documents and generate a CSV manifest listing image names and sources.
  - Workflow: table-structured-extraction
  - Outputs: csv
  - Selected engine: mcp
- `extract-all-images-from-shape-nodes-across-a-document-collection-and-compile-them-into-a-s.cs`
  - Task: Extract all images from shape nodes across a document collection and compile them into a single ZIP archive.
  - Workflow: image-shape-extraction
  - Outputs: docx
  - Selected engine: mcp
- `extract-a-document-segment-that-includes-nested-tables-and-ensure-nested-structures-are-re.cs`
  - Task: Extract a document segment that includes nested tables and ensure nested structures are retained in the new file.
  - Workflow: table-structured-extraction
  - Outputs: docx
  - Selected engine: mcp
- `create-a-reusable-extraction-utility-that-accepts-node-identifiers-and-returns-a-document.cs`
  - Task: Create a reusable extraction utility that accepts node identifiers and returns a Document containing the selected content.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: mcp
- `implement-error-handling-for-cases-where-the-start-node-appears-after-the-end-node-during.cs`
  - Task: Implement error handling for cases where the start node appears after the end node during extraction.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: mcp
- `implement-a-custom-node-filter-to-exclude-comments-while-extracting-content-between-two-pa.cs`
  - Task: Implement a custom node filter to exclude comments while extracting content between two paragraphs.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: existing_repo
- `extract-content-between-two-bookmark-nodes-and-replace-the-original-range-with-a-placehold.cs`
  - Task: Extract content between two bookmark nodes and replace the original range with a placeholder paragraph.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: llm
- `extract-content-between-a-paragraph-and-a-comment-node-then-log-the-extracted-text-to-a-mo.cs`
  - Task: Extract content between a paragraph and a comment node, then log the extracted text to a monitoring system.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: existing_repo
- `automate-extraction-of-footnote-content-between-specified-nodes-and-export-each-footnote-a.cs`
  - Task: Automate extraction of footnote content between specified nodes and export each footnote as a separate text file.
  - Workflow: targeted-node-extraction
  - Outputs: txt
  - Selected engine: mcp
- `create-a-command-line-tool-that-accepts-start-and-end-node-ids-and-outputs-the-extracted-s.cs`
  - Task: Create a command-line tool that accepts start and end node IDs and outputs the extracted segment as PDF.
  - Workflow: text-range-extraction
  - Outputs: pdf
  - Selected engine: mcp
- `create-a-unit-test-that-verifies-extraction-of-content-between-two-specific-paragraphs-ret.cs`
  - Task: Create a unit test that verifies extraction of content between two specific paragraphs retains original text styling.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: llm
- `develop-a-macro-that-calls-the-extraction-api-to-copy-selected-content-into-the-clipboard.cs`
  - Task: Develop a macro that calls the extraction API to copy selected content into the clipboard for pasting elsewhere.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: mcp
- `extract-a-range-that-starts-inside-a-shape-s-image-and-ends-at-a-field-preserving-both-ele.cs`
  - Task: Extract a range that starts inside a shape's image and ends at a field, preserving both elements.
  - Workflow: image-shape-extraction
  - Outputs: docx
  - Selected engine: mcp
- `extract-content-between-two-nodes-in-a-document-then-encrypt-the-resulting-file-using-a-pa.cs`
  - Task: Extract content between two nodes in a document, then encrypt the resulting file using a password.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: mcp
- `implement-parallel-processing-to-extract-node-ranges-from-multiple-documents-simultaneousl.cs`
  - Task: Implement parallel processing to extract node ranges from multiple documents simultaneously, improving performance.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: mcp
- `extract-images-from-shape-nodes-and-embed-them-directly-into-a-new-docx-document.cs`
  - Task: Extract images from shape nodes and embed them directly into a new DOCX document.
  - Workflow: image-shape-extraction
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Cross-document node insertion**
  - Symptom: Runtime ArgumentException: "The newChild was created from a different document than the one that created this node."
  - Root cause: Nodes are appended directly from the source document without being imported into the destination document.
  - Preferred fix: Create a NodeImporter (new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting)) and call ImportNode for each node before insertion.
  - Prevention: Always import or clone nodes before adding them to a different Document instance.

- **Missing Importing namespace**
  - Symptom: Compile error CS0234: "The type or namespace name 'Importing' does not exist in the namespace 'Aspose.Words'."
  - Root cause: The source code references NodeImporter without adding the required using directive or assembly reference.
  - Preferred fix: Add `using Aspose.Words.Importing;` and ensure the Aspose.Words assembly version includes the Importing namespace.
  - Prevention: Reference the correct namespace whenever NodeImporter is used.

- **Table type namespace omission**
  - Symptom: Compile errors CS1061 or CS0246 when using Table, Row, or Cell types.
  - Root cause: The code does not import Aspose.Words.Tables, so table-related classes are undefined.
  - Preferred fix: Add `using Aspose.Words.Tables;` at the top of the file.
  - Prevention: Include the Tables namespace whenever table structures are manipulated.

- **Undefined paragraph style**
  - Symptom: InvalidOperationException: "Cannot find style 'StartStyle'."
  - Root cause: The extraction logic searches for a style that has not been created in the document.
  - Preferred fix: Create the custom style programmatically before applying it, or verify that the style name exists in the document's Styles collection.
  - Prevention: Always ensure custom styles are defined prior to lookup.

- **Invalid bookmark insertion point**
  - Symptom: InvalidOperationException: "The node must be a block or an inline." when moving DocumentBuilder to a bookmark node.
  - Root cause: DocumentBuilder.MoveTo is called with a BookmarkStart node instead of its containing Paragraph.
  - Preferred fix: Move to `bookmark.BookmarkStart.ParentParagraph` or another block-level node.
  - Prevention: Only pass block or inline nodes to DocumentBuilder.MoveTo.

- **Comments API missing**
  - Symptom: Compile error CS0234: "The type or namespace name 'Comments' does not exist in the namespace 'Aspose.Words'."
  - Root cause: The code attempts to use a Comments namespace that is not part of the referenced Aspose.Words version.
  - Preferred fix: Use `using Aspose.Words.Comments;` (or the correct namespace for the version) and reference the Comments assembly if required.
  - Prevention: Verify the Aspose.Words version supports comment APIs and import the proper namespace.

- **Incorrect NodeImporter method name**
  - Symptom: CS1061: 'NodeImporter' does not contain a definition for 'Import'.
  - Root cause: The code calls a non-existent Import method instead of the correct ImportNode method.
  - Preferred fix: Replace calls to `importer.Import(...)` with `importer.ImportNode(node, true)`.
  - Prevention: Consult the API reference for the exact method signature of NodeImporter.

- **ImageData format property misuse**
  - Symptom: CS1061: 'ImageData' does not contain a definition for 'ImageFormat'.
  - Root cause: The code accesses a property that was removed or renamed in the current Aspose.Words version.
  - Preferred fix: Use `imageData.ImageBytes` together with `imageData.ImageType` to determine the format, or call `imageData.Save` with a desired SaveFormat.
  - Prevention: Refer to the latest Aspose.Words documentation for image handling properties.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
dotnet add package Aspose.Drawing.Common --version 26.10.0
dotnet add package Newtonsoft.Json
```

```powershell
Copy-Item ..\extraction\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Periodically verify that the Aspose.Words version used in the examples still contains the referenced namespaces and method signatures, updating using directives and API calls as needed.
