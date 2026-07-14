# Watermark Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Watermark category. Each file is a standalone console example selected from the verified 26.5.0 run.

## Snapshot

- Category: Watermark
- Slug: watermark
- Total examples: 34
- Publish-ready successful examples: 34 / 34
- Source run: 20260619_131835_59df5f
- Watermark API examples: 1
- Watermark Workflow examples: 33

## Category rules that shaped these examples

- Use native Aspose.Words APIs directly.
- Create local sample source documents or images when a task refers to an existing file, folder, stream, template, or input asset.
- Do not assume external files or folders already exist.
- Prefer documented `Document.Watermark` APIs when they directly fit the task.
- Keep validation narrow and task-specific.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.5.0

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\watermark\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `watermark/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

# PowerShell example
Copy-Item ..\watermark\load-a-word-document-from-a-file-path-and-add-a-text-watermark-using-watermark-settext.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-word-document-from-a-file-path-and-add-a-text-watermark-using-watermark-settext.cs` | Load a Word document from a file path and add a text watermark using Watermark.SetText. | Watermark Workflow | doc | mcp |
| 2 | `load-a-word-document-from-a-file-path-and-add-an-image-watermark-using-watermark-setimage.cs` | Load a Word document from a file path and add an image watermark using Watermark.SetImage. | Watermark Workflow | doc | mcp |
| 3 | `load-a-word-document-from-a-memory-stream-and-apply-a-text-watermark-without-writing-to-di.cs` | Load a Word document from a memory stream and apply a text watermark without writing to disk. | Watermark Workflow | doc | mcp |
| 4 | `insert-a-text-watermark-into-a-specific-table-cell-within-a-word-document-using-the-waterm.cs` | Insert a text watermark into a specific table cell within a Word document using the Watermark class. | Watermark Workflow | doc | mcp |
| 5 | `add-a-watermark-to-a-table-cell-that-spans-multiple-rows-and-columns-in-a-complex-word-tab.cs` | Add a watermark to a table cell that spans multiple rows and columns in a complex Word table. | Watermark Workflow | docx | mcp |
| 6 | `insert-a-watermark-into-a-table-cell-that-contains-merged-cells-without-disrupting-the-tab.cs` | Insert a watermark into a table cell that contains merged cells without disrupting the table layout. | Watermark Workflow | docx | mcp |
| 7 | `insert-a-watermark-into-each-cell-of-the-first-row-of-a-table-using-the-watermark-class.cs` | Insert a watermark into each cell of the first row of a table using the Watermark class. | Watermark Workflow | docx | mcp |
| 8 | `add-a-text-watermark-to-a-docx-document-using-watermark-settext-with-custom-font-settings.cs` | Add a text watermark to a DOCX document using Watermark.SetText with custom font settings. | Watermark Workflow | docx, doc | mcp |
| 9 | `use-watermark-settext-with-textwatermarkoptions-to-set-watermark-font-size-color-and-spaci.cs` | Use Watermark.SetText with TextWatermarkOptions to set watermark font size, color, and spacing. | Watermark Workflow | docx | mcp |
| 10 | `add-a-confidential-text-watermark-to-all-new-documents-created-by-an-automated-report-gene.cs` | Add a confidential text watermark to all new documents created by an automated report generator. | Watermark Workflow | doc | mcp |
| 11 | `combine-text-and-image-watermarks-by-first-setting-a-text-watermark-then-overlaying-an-ima.cs` | Combine text and image watermarks by first setting a text watermark then overlaying an image watermark. | Watermark Workflow | docx | mcp |
| 12 | `customize-image-watermark-opacity-and-scaling-by-configuring-imagewatermarkoptions-before.cs` | Customize image watermark opacity and scaling by configuring ImageWatermarkOptions before insertion. | Watermark Workflow | docx | mcp |
| 13 | `insert-an-image-watermark-from-a-system-drawing-image-object-into-a-document-after-calling.cs` | Insert an image watermark from a System.Drawing.Image object into a document after calling Optimize. | Watermark API | image | mcp |
| 14 | `insert-an-image-watermark-from-a-file-path-into-a-word-document-after-optimizing-the-docum.cs` | Insert an image watermark from a file path into a Word document after optimizing the document. | Watermark Workflow | doc | mcp |
| 15 | `use-watermark-setimage-with-a-stream-to-embed-a-logo-watermark-into-a-document-stored-in-a.cs` | Use Watermark.SetImage with a stream to embed a logo watermark into a document stored in Azure Blob storage. | Watermark Workflow | doc | mcp |
| 16 | `use-watermark-setimage-with-a-byte-array-stream-to-embed-a-dynamically-generated-barcode-w.cs` | Use Watermark.SetImage with a byte array stream to embed a dynamically generated barcode watermark. | Watermark Workflow | docx | existing_repo |
| 17 | `optimize-a-large-docx-file-before-applying-an-image-watermark-to-improve-performance-and-m.cs` | Optimize a large DOCX file before applying an image watermark to improve performance and memory usage. | Watermark Workflow | docx, doc | mcp |
| 18 | `remove-all-existing-watermarks-from-a-loaded-word-document-using-the-watermark-remove-meth.cs` | Remove all existing watermarks from a loaded Word document using the Watermark.Remove method. | Watermark Workflow | doc | mcp |
| 19 | `create-a-utility-method-that-removes-all-watermarks-from-a-document-using-watermark-remove.cs` | Create a utility method that removes all watermarks from a Document using Watermark.Remove. | Watermark Workflow | doc | mcp |
| 20 | `validate-that-a-document-contains-no-watermarks-before-publishing-by-using-watermarktype-n.cs` | Validate that a document contains no watermarks before publishing by using WatermarkType.None check. | Watermark Workflow | doc | mcp |
| 21 | `use-watermarktype-enumeration-to-verify-a-document-has-no-watermark-before-adding-a-new-on.cs` | Use WatermarkType enumeration to verify a document has no watermark before adding a new one. | Watermark Workflow | doc | mcp |
| 22 | `use-watermarktype-enumeration-to-switch-between-text-and-image-watermarks-based-on-user-se.cs` | Use WatermarkType enumeration to switch between text and image watermarks based on user selection. | Watermark Workflow | docx | mcp |
| 23 | `save-a-watermarked-word-document-directly-to-pdf-format-while-preserving-watermark-appeara.cs` | Save a watermarked Word document directly to PDF format while preserving watermark appearance. | Watermark Workflow | doc, pdf | mcp |
| 24 | `apply-a-text-watermark-to-a-word-document-and-then-save-the-document-as-docx.cs` | Apply a text watermark to a Word document and then save the document as DOCX. | Watermark Workflow | docx, doc | mcp |
| 25 | `apply-an-image-watermark-to-a-word-document-and-then-save-the-document-as-docx.cs` | Apply an image watermark to a Word document and then save the document as DOCX. | Watermark Workflow | docx, doc | existing_repo |
| 26 | `batch-process-a-folder-of-doc-files-to-add-the-same-image-watermark-to-each-document.cs` | Batch process a folder of DOC files to add the same image watermark to each document. | Watermark Workflow | doc | mcp |
| 27 | `batch-process-multiple-word-documents-in-a-directory-to-add-a-text-watermark-to-each-file.cs` | Batch process multiple Word documents in a directory to add a text watermark to each file. | Watermark Workflow | doc | mcp |
| 28 | `batch-process-multiple-word-documents-in-a-directory-to-remove-existing-watermarks-from-ea.cs` | Batch process multiple Word documents in a directory to remove existing watermarks from each file. | Watermark Workflow | doc | mcp |
| 29 | `batch-convert-docx-files-to-pdf-while-adding-a-corporate-logo-image-watermark-to-each-pdf.cs` | Batch convert DOCX files to PDF while adding a corporate logo image watermark to each PDF. | Watermark Workflow | docx, doc, pdf | mcp |
| 30 | `create-a-command-line-tool-that-accepts-a-directory-path-and-adds-a-specified-watermark-to.cs` | Create a command-line tool that accepts a directory path and adds a specified watermark to each file. | Watermark Workflow | docx | mcp |
| 31 | `use-a-configuration-file-to-define-watermark-text-font-and-opacity-then-apply-it-to-multip.cs` | Use a configuration file to define watermark text, font, and opacity, then apply it to multiple documents. | Watermark Workflow | doc | mcp |
| 32 | `implement-a-unit-test-that-verifies-watermark-remove-successfully-deletes-a-previously-add.cs` | Implement a unit test that verifies Watermark.Remove successfully deletes a previously added text watermark. | Watermark Workflow | docx | mcp |
| 33 | `add-a-watermark-to-a-document-opened-from-a-network-share-ensuring-proper-disposal-of-file.cs` | Add a watermark to a document opened from a network share, ensuring proper disposal of file handles. | Watermark Workflow | doc | mcp |
| 34 | `create-a-reusable-method-that-adds-a-configurable-text-watermark-to-any-document-object.cs` | Create a reusable method that adds a configurable text watermark to any Document object. | Watermark Workflow | doc | mcp |

## Common failure patterns seen during generation and how they were corrected

### Unsupported API invention

- Symptom: Generated code references members that do not exist in the selected package version.
- Fix: Replace invented members with documented Aspose.Words APIs already proven in this category.

### Missing local bootstrap inputs

- Symptom: The example assumes source files, folders, images, or data already exist.
- Fix: Create deterministic local inputs before loading, processing, or validating them.

### Over-broad validation

- Symptom: The example fails at runtime while checking unrelated document internals.
- Fix: Validate only the requested behavior and the existence of expected outputs.

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
