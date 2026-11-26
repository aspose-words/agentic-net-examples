# Extraction Examples for Aspose.Words for .NET

The Extraction category demonstrates how to isolate and copy specific node ranges from Word documents using Aspose.Words APIs. Verified examples show end-to-end workflows that create sample inputs, locate start and end boundaries, import nodes safely, and persist the result in the desired format.

## Snapshot

- Category: Extraction
- Slug: extraction
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261124_190740_1c00f5
- Text Range Extraction examples: 11
- Targeted Node Extraction examples: 9
- Table Structured Extraction examples: 6
- Input Bootstrap examples: 1
- Image Shape Extraction examples: 3

## Category rules that shaped these examples

- Do not reference System.Drawing; use Aspose.Drawing for all drawing operations.
- Do not assume external files exist; always create sample documents or streams within the example.
- Do not fabricate helper APIs; rely solely on documented Aspose.Words types such as DocumentBuilder, NodeImporter, and the classes in Aspose.Words.Tables.
- Enumerate actual document nodes and verify NodeType before processing.
- When extracting a range, import each node with NodeImporter before appending it to the destination body.
- Use Aspose.Words.Tables for Table, Row, and Cell handling, and Aspose.Words.Drawing for image manipulation.
- Initialize all non-nullable reference-type fields (e.g., Document, Section) at declaration to silence CS8618.
- Guard any potentially null node lookups with explicit null checks before dereferencing.
- Declare locals as nullable when a null value is a legitimate outcome and test for null before use.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.11.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
dotnet add package Aspose.Drawing.Common --version 26.10.0
dotnet add package Newtonsoft.Json

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\extraction\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `extraction/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0

# PowerShell example
Copy-Item ..\extraction\load-a-docx-file-extract-content-between-two-paragraphs-and-save-the-result-as-a-new-docx.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-extract-content-between-two-paragraphs-and-save-the-result-as-a-new-docx.cs` | Load a DOCX file, extract content between two paragraphs, and save the result as a new DOCX. | text-range-extraction | docx | mcp |
| 2 | `load-a-docm-file-extract-content-between-a-macro-enabled-field-and-a-paragraph-and-save-as.cs` | Load a DOCM file, extract content between a macro-enabled field and a paragraph, and save as DOCX. | targeted-node-extraction | docx, doc | mcp |
| 3 | `identify-a-start-run-node-and-an-end-bookmark-node-then-extract-the-intervening-nodes-into.cs` | Identify a start run node and an end bookmark node, then extract the intervening nodes into a document. | targeted-node-extraction | docx | mcp |
| 4 | `programmatically-determine-start-and-end-nodes-based-on-paragraph-styles-then-extract-the.cs` | Programmatically determine start and end nodes based on paragraph styles, then extract the styled content segment. | text-range-extraction | docx | mcp |
| 5 | `extract-a-mixed-node-range-that-starts-with-a-table-cell-and-ends-with-a-paragraph-maintai.cs` | Extract a mixed node range that starts with a table cell and ends with a paragraph, maintaining layout. | table-structured-extraction | docx | mcp |
| 6 | `extract-a-range-of-nodes-that-includes-tables-images-and-fields-preserving-original-hierar.cs` | Extract a range of nodes that includes tables, images, and fields, preserving original hierarchy in the output. | table-structured-extraction | docx | mcp |
| 7 | `extract-content-between-a-run-node-and-the-next-bookmark-then-convert-the-extracted-segmen.cs` | Extract content between a run node and the next bookmark, then convert the extracted segment to HTML format. | targeted-node-extraction | html | mcp |
| 8 | `extract-content-between-a-run-node-and-the-following-table-then-convert-the-extracted-port.cs` | Extract content between a run node and the following table, then convert the extracted portion to XPS format. | table-structured-extraction | xps | existing_repo |
| 9 | `use-the-extraction-api-to-copy-content-between-two-headings-and-insert-it-into-a-template.cs` | Use the extraction API to copy content between two headings and insert it into a template document. | text-range-extraction | docx | mcp |
| 10 | `use-documentbuilder-to-prepend-extracted-node-collection-to-the-beginning-of-a-new-documen.cs` | Use DocumentBuilder to prepend extracted node collection to the beginning of a new document before saving. | text-range-extraction | docx | mcp |
| 11 | `use-documentbuilder-to-insert-extracted-node-collection-into-a-new-document-at-a-custom-bo.cs` | Use DocumentBuilder to insert extracted node collection into a new document at a custom bookmark location. | targeted-node-extraction | docx | existing_repo |
| 12 | `duplicate-extracted-content-between-a-table-and-a-field-node-within-the-original-document.cs` | Duplicate extracted content between a table and a field node within the original document without altering formatting. | table-structured-extraction | docx | mcp |
| 13 | `save-extracted-content-as-a-docx-file-while-preserving-embedded-fields-and-their-evaluatio.cs` | Save extracted content as a DOCX file while preserving embedded fields and their evaluation results. | targeted-node-extraction | docx | existing_repo |
| 14 | `batch-process-multiple-word-files-extracting-content-between-specified-nodes-and-saving-ea.cs` | Batch process multiple Word files, extracting content between specified nodes and saving each extraction as an individual PDF. | input-bootstrap | pdf | mcp |
| 15 | `batch-extract-images-from-shape-nodes-in-documents-and-generate-a-csv-manifest-listing-ima.cs` | Batch extract images from shape nodes in documents and generate a CSV manifest listing image names and sources. | table-structured-extraction | csv | mcp |
| 16 | `extract-all-images-from-shape-nodes-across-a-document-collection-and-compile-them-into-a-s.cs` | Extract all images from shape nodes across a document collection and compile them into a single ZIP archive. | image-shape-extraction | docx | mcp |
| 17 | `extract-a-document-segment-that-includes-nested-tables-and-ensure-nested-structures-are-re.cs` | Extract a document segment that includes nested tables and ensure nested structures are retained in the new file. | table-structured-extraction | docx | mcp |
| 18 | `create-a-reusable-extraction-utility-that-accepts-node-identifiers-and-returns-a-document.cs` | Create a reusable extraction utility that accepts node identifiers and returns a Document containing the selected content. | text-range-extraction | docx | mcp |
| 19 | `implement-error-handling-for-cases-where-the-start-node-appears-after-the-end-node-during.cs` | Implement error handling for cases where the start node appears after the end node during extraction. | text-range-extraction | docx | mcp |
| 20 | `implement-a-custom-node-filter-to-exclude-comments-while-extracting-content-between-two-pa.cs` | Implement a custom node filter to exclude comments while extracting content between two paragraphs. | targeted-node-extraction | docx | existing_repo |
| 21 | `extract-content-between-two-bookmark-nodes-and-replace-the-original-range-with-a-placehold.cs` | Extract content between two bookmark nodes and replace the original range with a placeholder paragraph. | targeted-node-extraction | docx | llm |
| 22 | `extract-content-between-a-paragraph-and-a-comment-node-then-log-the-extracted-text-to-a-mo.cs` | Extract content between a paragraph and a comment node, then log the extracted text to a monitoring system. | targeted-node-extraction | docx | existing_repo |
| 23 | `automate-extraction-of-footnote-content-between-specified-nodes-and-export-each-footnote-a.cs` | Automate extraction of footnote content between specified nodes and export each footnote as a separate text file. | targeted-node-extraction | txt | mcp |
| 24 | `create-a-command-line-tool-that-accepts-start-and-end-node-ids-and-outputs-the-extracted-s.cs` | Create a command-line tool that accepts start and end node IDs and outputs the extracted segment as PDF. | text-range-extraction | pdf | mcp |
| 25 | `create-a-unit-test-that-verifies-extraction-of-content-between-two-specific-paragraphs-ret.cs` | Create a unit test that verifies extraction of content between two specific paragraphs retains original text styling. | text-range-extraction | docx | llm |
| 26 | `develop-a-macro-that-calls-the-extraction-api-to-copy-selected-content-into-the-clipboard.cs` | Develop a macro that calls the extraction API to copy selected content into the clipboard for pasting elsewhere. | text-range-extraction | docx | mcp |
| 27 | `extract-a-range-that-starts-inside-a-shape-s-image-and-ends-at-a-field-preserving-both-ele.cs` | Extract a range that starts inside a shape's image and ends at a field, preserving both elements. | image-shape-extraction | docx | mcp |
| 28 | `extract-content-between-two-nodes-in-a-document-then-encrypt-the-resulting-file-using-a-pa.cs` | Extract content between two nodes in a document, then encrypt the resulting file using a password. | text-range-extraction | docx | mcp |
| 29 | `implement-parallel-processing-to-extract-node-ranges-from-multiple-documents-simultaneousl.cs` | Implement parallel processing to extract node ranges from multiple documents simultaneously, improving performance. | text-range-extraction | docx | mcp |
| 30 | `extract-images-from-shape-nodes-and-embed-them-directly-into-a-new-docx-document.cs` | Extract images from shape nodes and embed them directly into a new DOCX document. | image-shape-extraction | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Cross-document node insertion

- Symptom: Runtime ArgumentException: "The newChild was created from a different document than the one that created this node."
- Root cause: Nodes are appended directly from the source document without being imported into the destination document.
- Preferred fix: Create a NodeImporter (new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting)) and call ImportNode for each node before insertion.
- Prevention: Always import or clone nodes before adding them to a different Document instance.

### Missing Importing namespace

- Symptom: Compile error CS0234: "The type or namespace name 'Importing' does not exist in the namespace 'Aspose.Words'."
- Root cause: The source code references NodeImporter without adding the required using directive or assembly reference.
- Preferred fix: Add `using Aspose.Words.Importing;` and ensure the Aspose.Words assembly version includes the Importing namespace.
- Prevention: Reference the correct namespace whenever NodeImporter is used.

### Table type namespace omission

- Symptom: Compile errors CS1061 or CS0246 when using Table, Row, or Cell types.
- Root cause: The code does not import Aspose.Words.Tables, so table-related classes are undefined.
- Preferred fix: Add `using Aspose.Words.Tables;` at the top of the file.
- Prevention: Include the Tables namespace whenever table structures are manipulated.

### Undefined paragraph style

- Symptom: InvalidOperationException: "Cannot find style 'StartStyle'."
- Root cause: The extraction logic searches for a style that has not been created in the document.
- Preferred fix: Create the custom style programmatically before applying it, or verify that the style name exists in the document's Styles collection.
- Prevention: Always ensure custom styles are defined prior to lookup.

### Invalid bookmark insertion point

- Symptom: InvalidOperationException: "The node must be a block or an inline." when moving DocumentBuilder to a bookmark node.
- Root cause: DocumentBuilder.MoveTo is called with a BookmarkStart node instead of its containing Paragraph.
- Preferred fix: Move to `bookmark.BookmarkStart.ParentParagraph` or another block-level node.
- Prevention: Only pass block or inline nodes to DocumentBuilder.MoveTo.

### Comments API missing

- Symptom: Compile error CS0234: "The type or namespace name 'Comments' does not exist in the namespace 'Aspose.Words'."
- Root cause: The code attempts to use a Comments namespace that is not part of the referenced Aspose.Words version.
- Preferred fix: Use `using Aspose.Words.Comments;` (or the correct namespace for the version) and reference the Comments assembly if required.
- Prevention: Verify the Aspose.Words version supports comment APIs and import the proper namespace.

### Incorrect NodeImporter method name

- Symptom: CS1061: 'NodeImporter' does not contain a definition for 'Import'.
- Root cause: The code calls a non-existent Import method instead of the correct ImportNode method.
- Preferred fix: Replace calls to `importer.Import(...)` with `importer.ImportNode(node, true)`.
- Prevention: Consult the API reference for the exact method signature of NodeImporter.

### ImageData format property misuse

- Symptom: CS1061: 'ImageData' does not contain a definition for 'ImageFormat'.
- Root cause: The code accesses a property that was removed or renamed in the current Aspose.Words version.
- Preferred fix: Use `imageData.ImageBytes` together with `imageData.ImageType` to determine the format, or call `imageData.Save` with a desired SaveFormat.
- Prevention: Refer to the latest Aspose.Words documentation for image handling properties.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Periodically verify that the Aspose.Words version used in the examples still contains the referenced namespaces and method signatures, updating using directives and API calls as needed.
