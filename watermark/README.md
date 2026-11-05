# Watermark Examples for Aspose.Words for .NET

The Watermark category demonstrates how to apply text and image watermarks to Word documents using Aspose.Words 26.10.0. Verified examples cover document-level, table-cell-level, batch, and combined watermark scenarios with correct API usage.

## Snapshot

- Category: Watermark
- Slug: watermark
- Total examples: 33
- Publish-ready successful examples: 33 / 33
- Source run: 20261102_031946_08fb5b
- Watermark Workflow examples: 33

## Category rules that shaped these examples

- Do not reference members that are not present in the Aspose.Words 26.10.0 API, such as WatermarkLayout, TextBox.Text, or Shape.Transparency.
- Use Document.Watermark.SetText(string, TextWatermarkOptions) for text watermarks and Document.Watermark.SetImage(string, ImageWatermarkOptions) for image watermarks.
- When inserting a watermark into a table cell, retrieve the watermark shape from the primary header, detach it, set WrapType to Inline, and append it to the target cell's first paragraph.
- Validate that image files exist before calling SetImage, and ensure that options objects are instantiated before use.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.10.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
dotnet add package SkiaSharp --version 4.152.1

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
dotnet add package Aspose.Words --version 26.10.0

# PowerShell example
Copy-Item ..\watermark\load-a-word-document-from-a-file-path-and-add-a-text-watermark-using-watermark-settext.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-word-document-from-a-file-path-and-add-a-text-watermark-using-watermark-settext.cs` | Load a Word document from a file path and add a text watermark using Watermark.SetText. | watermark-workflow | doc | mcp |
| 2 | `load-a-word-document-from-a-file-path-and-add-an-image-watermark-using-watermark-setimage.cs` | Load a Word document from a file path and add an image watermark using Watermark.SetImage. | watermark-workflow | doc | existing_repo |
| 3 | `load-a-word-document-from-a-memory-stream-and-apply-a-text-watermark-without-writing-to-di.cs` | Load a Word document from a memory stream and apply a text watermark without writing to disk. | watermark-workflow | doc | mcp |
| 4 | `insert-a-text-watermark-into-a-specific-table-cell-within-a-word-document-using-the-waterm.cs` | Insert a text watermark into a specific table cell within a Word document using the Watermark class. | watermark-workflow | doc | llm |
| 5 | `add-a-watermark-to-a-table-cell-that-spans-multiple-rows-and-columns-in-a-complex-word-tab.cs` | Add a watermark to a table cell that spans multiple rows and columns in a complex Word table. | watermark-workflow | docx | mcp |
| 6 | `insert-a-watermark-into-a-table-cell-that-contains-merged-cells-without-disrupting-the-tab.cs` | Insert a watermark into a table cell that contains merged cells without disrupting the table layout. | watermark-workflow | docx | mcp |
| 7 | `insert-a-watermark-into-each-cell-of-the-first-row-of-a-table-using-the-watermark-class.cs` | Insert a watermark into each cell of the first row of a table using the Watermark class. | watermark-workflow | docx | mcp |
| 8 | `add-a-text-watermark-to-a-docx-document-using-watermark-settext-with-custom-font-settings.cs` | Add a text watermark to a DOCX document using Watermark.SetText with custom font settings. | watermark-workflow | docx, doc | mcp |
| 9 | `use-watermark-settext-with-textwatermarkoptions-to-set-watermark-font-size-color-and-spaci.cs` | Use Watermark.SetText with TextWatermarkOptions to set watermark font size, color, and spacing. | watermark-workflow | docx | existing_repo |
| 10 | `add-a-confidential-text-watermark-to-all-new-documents-created-by-an-automated-report-gene.cs` | Add a confidential text watermark to all new documents created by an automated report generator. | watermark-workflow | doc | mcp |
| 11 | `combine-text-and-image-watermarks-by-first-setting-a-text-watermark-then-overlaying-an-ima.cs` | Combine text and image watermarks by first setting a text watermark then overlaying an image watermark. | watermark-workflow | docx | existing_repo |
| 12 | `customize-image-watermark-opacity-and-scaling-by-configuring-imagewatermarkoptions-before.cs` | Customize image watermark opacity and scaling by configuring ImageWatermarkOptions before insertion. | watermark-workflow | docx | existing_repo |
| 13 | `insert-an-image-watermark-from-a-file-path-into-a-word-document-after-optimizing-the-docum.cs` | Insert an image watermark from a file path into a Word document after optimizing the document. | watermark-workflow | doc | existing_repo |
| 14 | `use-watermark-setimage-with-a-stream-to-embed-a-logo-watermark-into-a-document-stored-in-a.cs` | Use Watermark.SetImage with a stream to embed a logo watermark into a document stored in Azure Blob storage. | watermark-workflow | doc | existing_repo |
| 15 | `use-watermark-setimage-with-a-byte-array-stream-to-embed-a-dynamically-generated-barcode-w.cs` | Use Watermark.SetImage with a byte array stream to embed a dynamically generated barcode watermark. | watermark-workflow | docx | existing_repo |
| 16 | `optimize-a-large-docx-file-before-applying-an-image-watermark-to-improve-performance-and-m.cs` | Optimize a large DOCX file before applying an image watermark to improve performance and memory usage. | watermark-workflow | docx, doc | existing_repo |
| 17 | `remove-all-existing-watermarks-from-a-loaded-word-document-using-the-watermark-remove-meth.cs` | Remove all existing watermarks from a loaded Word document using the Watermark.Remove method. | watermark-workflow | doc | mcp |
| 18 | `create-a-utility-method-that-removes-all-watermarks-from-a-document-using-watermark-remove.cs` | Create a utility method that removes all watermarks from a Document using Watermark.Remove. | watermark-workflow | doc | mcp |
| 19 | `validate-that-a-document-contains-no-watermarks-before-publishing-by-using-watermarktype-n.cs` | Validate that a document contains no watermarks before publishing by using WatermarkType.None check. | watermark-workflow | doc | llm |
| 20 | `use-watermarktype-enumeration-to-verify-a-document-has-no-watermark-before-adding-a-new-on.cs` | Use WatermarkType enumeration to verify a document has no watermark before adding a new one. | watermark-workflow | doc | mcp |
| 21 | `use-watermarktype-enumeration-to-switch-between-text-and-image-watermarks-based-on-user-se.cs` | Use WatermarkType enumeration to switch between text and image watermarks based on user selection. | watermark-workflow | docx | existing_repo |
| 22 | `save-a-watermarked-word-document-directly-to-pdf-format-while-preserving-watermark-appeara.cs` | Save a watermarked Word document directly to PDF format while preserving watermark appearance. | watermark-workflow | doc, pdf | mcp |
| 23 | `apply-a-text-watermark-to-a-word-document-and-then-save-the-document-as-docx.cs` | Apply a text watermark to a Word document and then save the document as DOCX. | watermark-workflow | docx, doc | mcp |
| 24 | `apply-an-image-watermark-to-a-word-document-and-then-save-the-document-as-docx.cs` | Apply an image watermark to a Word document and then save the document as DOCX. | watermark-workflow | docx, doc | existing_repo |
| 25 | `batch-process-a-folder-of-doc-files-to-add-the-same-image-watermark-to-each-document.cs` | Batch process a folder of DOC files to add the same image watermark to each document. | watermark-workflow | doc | llm |
| 26 | `batch-process-multiple-word-documents-in-a-directory-to-add-a-text-watermark-to-each-file.cs` | Batch process multiple Word documents in a directory to add a text watermark to each file. | watermark-workflow | doc | mcp |
| 27 | `batch-process-multiple-word-documents-in-a-directory-to-remove-existing-watermarks-from-ea.cs` | Batch process multiple Word documents in a directory to remove existing watermarks from each file. | watermark-workflow | doc | mcp |
| 28 | `batch-convert-docx-files-to-pdf-while-adding-a-corporate-logo-image-watermark-to-each-pdf.cs` | Batch convert DOCX files to PDF while adding a corporate logo image watermark to each PDF. | watermark-workflow | docx, doc, pdf | existing_repo |
| 29 | `create-a-command-line-tool-that-accepts-a-directory-path-and-adds-a-specified-watermark-to.cs` | Create a command-line tool that accepts a directory path and adds a specified watermark to each file. | watermark-workflow | docx | mcp |
| 30 | `use-a-configuration-file-to-define-watermark-text-font-and-opacity-then-apply-it-to-multip.cs` | Use a configuration file to define watermark text, font, and opacity, then apply it to multiple documents. | watermark-workflow | doc | existing_repo |
| 31 | `implement-a-unit-test-that-verifies-watermark-remove-successfully-deletes-a-previously-add.cs` | Implement a unit test that verifies Watermark.Remove successfully deletes a previously added text watermark. | watermark-workflow | docx | mcp |
| 32 | `add-a-watermark-to-a-document-opened-from-a-network-share-ensuring-proper-disposal-of-file.cs` | Add a watermark to a document opened from a network share, ensuring proper disposal of file handles. | watermark-workflow | doc | mcp |
| 33 | `create-a-reusable-method-that-adds-a-configurable-text-watermark-to-any-document-object.cs` | Create a reusable method that adds a configurable text watermark to any Document object. | watermark-workflow | doc | existing_repo |

## Common failure patterns seen during generation and how they were corrected

### Incorrect overload for SetImage

- Symptom: Compilation fails with CS1503 because a string or stream is passed where an SKBitmap is expected.
- Root cause: The code uses the overload that requires a SkiaSharp.SKBitmap instead of the overload that accepts a file path or stream.
- Preferred fix: Call Document.Watermark.SetImage(string path, ImageWatermarkOptions) or SetImage(Stream stream, ImageWatermarkOptions) as demonstrated in the verified examples.
- Prevention: Always select the overload that matches the argument type; consult the API reference for the correct signatures.

### Reference to non-existent types or members

- Symptom: Build errors such as CS0246, CS1061, or CS0117 indicate missing types like Table, Cell, WatermarkLayout, or members like Shape.Text.
- Root cause: The code invents or assumes APIs that are not part of the current Aspose.Words version.
- Preferred fix: Remove invented references and use only documented types (DocumentBuilder, Table, Cell, Shape) and members (WrapType, FillColor, etc.).
- Prevention: Rely exclusively on the official Aspose.Words documentation and the verified code snippets for supported members.

### Misuse of using directives

- Symptom: Compilation fails with CS0138 because a type is treated as a namespace in a using statement.
- Root cause: The code attempts to write "using Watermark;" or similar, treating the Watermark class as a namespace.
- Preferred fix: Remove the invalid using directive; reference the Watermark class directly via the Document instance.
- Prevention: Only include using statements for actual namespaces; never use a class name in a using directive.

### Assignment to read-only TextBox property

- Symptom: CS0200 errors occur when the code tries to assign to Shape.TextBox or its Text property.
- Root cause: The TextBox property of Shape is read-only, and its Text member does not exist in this API version.
- Preferred fix: Create a Shape, set its properties (Width, Height, FillColor, etc.), and add a Run with the desired text inside a Paragraph appended to the shape.
- Prevention: Do not attempt to modify Shape.TextBox; instead build the watermark content using Paragraph and Run objects.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Periodically verify that the Watermark API signatures have not changed in newer Aspose.Words releases and update the examples accordingly.
