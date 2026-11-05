---
name: extraction
description: Verified C# examples for extraction scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Extraction

## Purpose

Provide an AI coding assistant with a concise, end-to-end reference for building reliable extraction utilities that locate start and end markers, import nodes safely, and persist the result using only supported Aspose.Words APIs.

## Non-negotiable conventions

- Do not reference System.Drawing, assume no pre-existing files, and avoid inventing custom extraction helpers.
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

1. Ensure every extracted node is imported before insertion and that the saved file is present on disk.

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
  - Selected engine: existing_repo
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
  - Selected engine: llm
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
  - Selected engine: llm
- `use-documentbuilder-to-prepend-extracted-node-collection-to-the-beginning-of-a-new-documen.cs`
  - Task: Use DocumentBuilder to prepend extracted node collection to the beginning of a new document before saving.
  - Workflow: text-range-extraction
  - Outputs: docx
  - Selected engine: mcp
- `use-documentbuilder-to-insert-extracted-node-collection-into-a-new-document-at-a-custom-bo.cs`
  - Task: Use DocumentBuilder to insert extracted node collection into a new document at a custom bookmark location.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: llm
- `duplicate-extracted-content-between-a-table-and-a-field-node-within-the-original-document.cs`
  - Task: Duplicate extracted content between a table and a field node within the original document without altering formatting.
  - Workflow: table-structured-extraction
  - Outputs: docx
  - Selected engine: existing_repo
- `save-extracted-content-as-a-docx-file-while-preserving-embedded-fields-and-their-evaluatio.cs`
  - Task: Save extracted content as a DOCX file while preserving embedded fields and their evaluation results.
  - Workflow: targeted-node-extraction
  - Outputs: docx
  - Selected engine: mcp
- `batch-process-multiple-word-files-extracting-content-between-specified-nodes-and-saving-ea.cs`
  - Task: Batch process multiple Word files, extracting content between specified nodes and saving each extraction as an individual PDF.
  - Workflow: input-bootstrap
  - Outputs: pdf
  - Selected engine: llm
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
  - Selected engine: existing_repo
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
  - Selected engine: llm
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
  - Selected engine: llm
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
  - Symptom: Runtime throws ArgumentException: "The newChild was created from a different document than the one that created this node."
  - Root cause: Nodes from the source Document are inserted directly into a destination Document without using NodeImporter, leaving them bound to the original document.
  - Preferred fix: Create a NodeImporter (ImportFormatMode.KeepSourceFormatting) and import each node before appending or inserting it into the target document.
  - Prevention: Always import or clone nodes via NodeImporter whenever a node moves between documents.

- **Missing API members**
  - Symptom: Compile errors CS1061 or CS0234 indicating undefined members such as FieldStart.EndNode, Document.ExtractContent, or DocumentImporter.
  - Root cause: Code references members that do not exist in the current Aspose.Words version or omits required namespaces.
  - Preferred fix: Replace invalid members with the correct API (e.g., use Field.StartNode, Document.Range, and avoid non-existent extension methods) and add the proper using directives.
  - Prevention: Consult the official Aspose.Words documentation for the exact member names and required namespaces before coding.

- **Incorrect type casting**
  - Symptom: Compiler reports CS0266: cannot implicitly convert type 'Node' to 'Paragraph'.
  - Root cause: Assuming a generic Node is a Paragraph without checking its NodeType.
  - Preferred fix: Verify node.NodeType == NodeType.Paragraph and cast explicitly (e.g., (Paragraph)node) or use pattern matching.
  - Prevention: Validate node types before casting and use explicit casts only after confirming the type.

- **Table namespace omission**
  - Symptom: Compile errors because Table, Row, or Cell types are not recognized.
  - Root cause: Missing the Aspose.Words.Tables namespace in the source file.
  - Preferred fix: Add `using Aspose.Words.Tables;` or fully qualify the types (e.g., Aspose.Words.Tables.Table).
  - Prevention: Include all required Aspose.Words sub-namespaces at the top of each file.

- **Bookmark boundary misuse**
  - Symptom: InvalidOperationException when extracting between bookmarks or when inserting placeholder paragraphs.
  - Root cause: Using incorrect start/end nodes (e.g., BookmarkStart vs. BookmarkEnd) or inserting a Paragraph into a non-block container.
  - Preferred fix: Reference BookmarkStart and BookmarkEnd correctly and insert placeholder Paragraphs only into a Body or Section.
  - Prevention: Validate bookmark boundaries and ensure placeholder nodes are added to a block-level container.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.10.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
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

- Update all sample snippets to use NodeImporter for every node transfer and verify that the Aspose.Words version (26.10.0) matches the API signatures used.
