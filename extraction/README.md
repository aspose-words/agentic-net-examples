# Extraction Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Extraction category. Each file is a standalone console example selected from the verified 26.5.0 run.

## Snapshot

- Category: Extraction
- Slug: extraction
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20260619_131835_59df5f
- Image Shape Extraction examples: 3
- Input Bootstrap examples: 1
- Table Structured Extraction examples: 6
- Targeted Node Extraction examples: 9
- Text Range Extraction examples: 11

## Category rules that shaped these examples

- Do not use System.Drawing in this category.
- Do not assume source documents or assets already exist; bootstrap them locally when needed.
- Do not invent helper APIs for extracting tables, images, comments, bookmarks, or metadata.
- Do not skip writing the requested output artifact when the task expects one.
- Enumerate actual document nodes and validate node types before extracting content.
- Create local sample DOC, DOCX, TXT, HTML, XML, image, stream, or folder inputs whenever the task implies an existing source.
- Use Aspose.Words.Tables for Table, Row, and Cell types when structured extraction is needed.
- Use Newtonsoft.Json for JSON serialization tasks and Aspose.Drawing instead of System.Drawing when drawing-related types are needed.
- Initialize all non-nullable reference type properties to avoid CS8618 warnings.
- Avoid CS8600, CS8602, and CS8604 by guarding maybe-null values before dereference or assignment.
- Declare nullable locals when a value may legitimately be null and null-check before use.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.5.0
- Newtonsoft.Json

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0
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
dotnet add package Aspose.Words --version 26.5.0
dotnet add package Newtonsoft.Json

# PowerShell example
Copy-Item ..\extraction\load-a-docx-file-extract-content-between-two-paragraphs-and-save-the-result-as-a-new-docx.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-extract-content-between-two-paragraphs-and-save-the-result-as-a-new-docx.cs` | Load a DOCX file, extract content between two paragraphs, and save the result as a new DOCX. | Text Range Extraction | docx | llm |
| 2 | `load-a-docm-file-extract-content-between-a-macro-enabled-field-and-a-paragraph-and-save-as.cs` | Load a DOCM file, extract content between a macro-enabled field and a paragraph, and save as DOCX. | Targeted Node Extraction | docx, doc | llm |
| 3 | `identify-a-start-run-node-and-an-end-bookmark-node-then-extract-the-intervening-nodes-into.cs` | Identify a start run node and an end bookmark node, then extract the intervening nodes into a document. | Targeted Node Extraction | docx | mcp |
| 4 | `programmatically-determine-start-and-end-nodes-based-on-paragraph-styles-then-extract-the.cs` | Programmatically determine start and end nodes based on paragraph styles, then extract the styled content segment. | Text Range Extraction | docx | mcp |
| 5 | `extract-a-mixed-node-range-that-starts-with-a-table-cell-and-ends-with-a-paragraph-maintai.cs` | Extract a mixed node range that starts with a table cell and ends with a paragraph, maintaining layout. | Table Structured Extraction | docx | mcp |
| 6 | `extract-a-range-of-nodes-that-includes-tables-images-and-fields-preserving-original-hierar.cs` | Extract a range of nodes that includes tables, images, and fields, preserving original hierarchy in the output. | Table Structured Extraction | docx | llm |
| 7 | `extract-content-between-a-run-node-and-the-next-bookmark-then-convert-the-extracted-segmen.cs` | Extract content between a run node and the next bookmark, then convert the extracted segment to HTML format. | Targeted Node Extraction | html | mcp |
| 8 | `extract-content-between-a-run-node-and-the-following-table-then-convert-the-extracted-port.cs` | Extract content between a run node and the following table, then convert the extracted portion to XPS format. | Table Structured Extraction | xps | mcp |
| 9 | `use-the-extraction-api-to-copy-content-between-two-headings-and-insert-it-into-a-template.cs` | Use the extraction API to copy content between two headings and insert it into a template document. | Text Range Extraction | docx | mcp |
| 10 | `use-documentbuilder-to-prepend-extracted-node-collection-to-the-beginning-of-a-new-documen.cs` | Use DocumentBuilder to prepend extracted node collection to the beginning of a new document before saving. | Text Range Extraction | docx | mcp |
| 11 | `use-documentbuilder-to-insert-extracted-node-collection-into-a-new-document-at-a-custom-bo.cs` | Use DocumentBuilder to insert extracted node collection into a new document at a custom bookmark location. | Targeted Node Extraction | docx | llm |
| 12 | `duplicate-extracted-content-between-a-table-and-a-field-node-within-the-original-document.cs` | Duplicate extracted content between a table and a field node within the original document without altering formatting. | Table Structured Extraction | docx | mcp |
| 13 | `save-extracted-content-as-a-docx-file-while-preserving-embedded-fields-and-their-evaluatio.cs` | Save extracted content as a DOCX file while preserving embedded fields and their evaluation results. | Targeted Node Extraction | docx | mcp |
| 14 | `batch-process-multiple-word-files-extracting-content-between-specified-nodes-and-saving-ea.cs` | Batch process multiple Word files, extracting content between specified nodes and saving each extraction as an individual PDF. | Input Bootstrap | pdf | mcp |
| 15 | `batch-extract-images-from-shape-nodes-in-documents-and-generate-a-csv-manifest-listing-ima.cs` | Batch extract images from shape nodes in documents and generate a CSV manifest listing image names and sources. | Table Structured Extraction | csv | mcp |
| 16 | `extract-all-images-from-shape-nodes-across-a-document-collection-and-compile-them-into-a-s.cs` | Extract all images from shape nodes across a document collection and compile them into a single ZIP archive. | Image Shape Extraction | docx | mcp |
| 17 | `extract-a-document-segment-that-includes-nested-tables-and-ensure-nested-structures-are-re.cs` | Extract a document segment that includes nested tables and ensure nested structures are retained in the new file. | Table Structured Extraction | docx | mcp |
| 18 | `create-a-reusable-extraction-utility-that-accepts-node-identifiers-and-returns-a-document.cs` | Create a reusable extraction utility that accepts node identifiers and returns a Document containing the selected content. | Text Range Extraction | docx | mcp |
| 19 | `implement-error-handling-for-cases-where-the-start-node-appears-after-the-end-node-during.cs` | Implement error handling for cases where the start node appears after the end node during extraction. | Text Range Extraction | docx | mcp |
| 20 | `implement-a-custom-node-filter-to-exclude-comments-while-extracting-content-between-two-pa.cs` | Implement a custom node filter to exclude comments while extracting content between two paragraphs. | Targeted Node Extraction | docx | mcp |
| 21 | `extract-content-between-two-bookmark-nodes-and-replace-the-original-range-with-a-placehold.cs` | Extract content between two bookmark nodes and replace the original range with a placeholder paragraph. | Targeted Node Extraction | docx | existing_repo |
| 22 | `extract-content-between-a-paragraph-and-a-comment-node-then-log-the-extracted-text-to-a-mo.cs` | Extract content between a paragraph and a comment node, then log the extracted text to a monitoring system. | Targeted Node Extraction | docx | mcp |
| 23 | `automate-extraction-of-footnote-content-between-specified-nodes-and-export-each-footnote-a.cs` | Automate extraction of footnote content between specified nodes and export each footnote as a separate text file. | Targeted Node Extraction | txt | mcp |
| 24 | `create-a-command-line-tool-that-accepts-start-and-end-node-ids-and-outputs-the-extracted-s.cs` | Create a command-line tool that accepts start and end node IDs and outputs the extracted segment as PDF. | Text Range Extraction | pdf | llm |
| 25 | `create-a-unit-test-that-verifies-extraction-of-content-between-two-specific-paragraphs-ret.cs` | Create a unit test that verifies extraction of content between two specific paragraphs retains original text styling. | Text Range Extraction | docx | existing_repo |
| 26 | `develop-a-macro-that-calls-the-extraction-api-to-copy-selected-content-into-the-clipboard.cs` | Develop a macro that calls the extraction API to copy selected content into the clipboard for pasting elsewhere. | Text Range Extraction | docx | llm |
| 27 | `extract-a-range-that-starts-inside-a-shape-s-image-and-ends-at-a-field-preserving-both-ele.cs` | Extract a range that starts inside a shape's image and ends at a field, preserving both elements. | Image Shape Extraction | docx | existing_repo |
| 28 | `extract-content-between-two-nodes-in-a-document-then-encrypt-the-resulting-file-using-a-pa.cs` | Extract content between two nodes in a document, then encrypt the resulting file using a password. | Text Range Extraction | docx | existing_repo |
| 29 | `implement-parallel-processing-to-extract-node-ranges-from-multiple-documents-simultaneousl.cs` | Implement parallel processing to extract node ranges from multiple documents simultaneously, improving performance. | Text Range Extraction | docx | mcp |
| 30 | `extract-images-from-shape-nodes-and-embed-them-directly-into-a-new-docx-document.cs` | Extract images from shape nodes and embed them directly into a new DOCX document. | Image Shape Extraction | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Invalid node insertion in destination document

- Symptom: Runtime failures such as 'Cannot insert a node of this type at this location'.
- Fix: Rebuild a valid destination document structure explicitly and insert cloned inline or block nodes only into supported parents.

### Table namespace assumptions

- Symptom: Compile errors because Table, Row, or Cell were used without Aspose.Words.Tables.
- Fix: Import Aspose.Words.Tables or use fully qualified table node type names.

### Weak bookmark boundary logic

- Symptom: Bookmark-bounded extraction or replacement fails because the wrong nodes are used or placeholder insertion is invalid.
- Fix: Use real BookmarkStart and BookmarkEnd boundaries and insert a valid placeholder Paragraph into a supported block container.

### Footnote-specific API issues

- Symptom: Footnote export or enumeration fails because footnote APIs or namespaces are used incorrectly.
- Fix: Use Aspose.Words.Notes.Footnote and Aspose.Words.Notes.FootnoteType explicitly and enumerate actual footnote nodes.

### Font or drawing ambiguity

- Symptom: Compile errors due to System.Drawing usage or ambiguous Font references.
- Fix: Use Aspose.Drawing only and prefer explicit Aspose.Drawing type names when ambiguity is possible.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.5.0 run.
- Preserve file-to-task traceability when updating this folder.
- Keep examples standalone and bootstrap local inputs inside the example whenever external sources are mentioned.
