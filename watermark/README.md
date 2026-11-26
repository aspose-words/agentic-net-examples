# Watermark Examples for Aspose.Words for .NET

The Watermark category demonstrates how to apply text and image watermarks to Word documents using Aspose.Words 26.11.0. Verified examples show correct usage of Document.Watermark APIs for document-wide, table-cell, and stream-based scenarios.

## Snapshot

- Category: Watermark
- Slug: watermark
- Total examples: 33
- Publish-ready successful examples: 33 / 33
- Source run: 20261124_190740_1c00f5
- Watermark Workflow examples: 33

## Category rules that shaped these examples

- Do not reference members that are not present in Aspose.Words 26.11.0, such as WatermarkLayout members, ImageWatermarkOptions inside the Watermark type, or System.Drawing APIs for image generation.
- Use Document.Watermark.SetText(string, TextWatermarkOptions) for customizable text watermarks.
- Use Document.Watermark.SetImage(string, ImageWatermarkOptions) or the stream overload for image watermarks.
- For table-cell watermarks, create a Shape, set WrapType to None, and append it to the cell's first paragraph.
- Validate that all objects passed to Watermark methods are non-null and of the expected type before invoking the API.

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
dotnet add package Aspose.Words --version 26.11.0

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
| 4 | `insert-a-text-watermark-into-a-specific-table-cell-within-a-word-document-using-the-waterm.cs` | Insert a text watermark into a specific table cell within a Word document using the Watermark class. | watermark-workflow | doc | mcp |
| 5 | `add-a-watermark-to-a-table-cell-that-spans-multiple-rows-and-columns-in-a-complex-word-tab.cs` | Add a watermark to a table cell that spans multiple rows and columns in a complex Word table. | watermark-workflow | docx | mcp |
| 6 | `insert-a-watermark-into-a-table-cell-that-contains-merged-cells-without-disrupting-the-tab.cs` | Insert a watermark into a table cell that contains merged cells without disrupting the table layout. | watermark-workflow | docx | mcp |
| 7 | `insert-a-watermark-into-each-cell-of-the-first-row-of-a-table-using-the-watermark-class.cs` | Insert a watermark into each cell of the first row of a table using the Watermark class. | watermark-workflow | docx | mcp |
| 8 | `add-a-text-watermark-to-a-docx-document-using-watermark-settext-with-custom-font-settings.cs` | Add a text watermark to a DOCX document using Watermark.SetText with custom font settings. | watermark-workflow | docx, doc | existing_repo |
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
| 19 | `validate-that-a-document-contains-no-watermarks-before-publishing-by-using-watermarktype-n.cs` | Validate that a document contains no watermarks before publishing by using WatermarkType.None check. | watermark-workflow | doc | mcp |
| 20 | `use-watermarktype-enumeration-to-verify-a-document-has-no-watermark-before-adding-a-new-on.cs` | Use WatermarkType enumeration to verify a document has no watermark before adding a new one. | watermark-workflow | doc | mcp |
| 21 | `use-watermarktype-enumeration-to-switch-between-text-and-image-watermarks-based-on-user-se.cs` | Use WatermarkType enumeration to switch between text and image watermarks based on user selection. | watermark-workflow | docx | existing_repo |
| 22 | `save-a-watermarked-word-document-directly-to-pdf-format-while-preserving-watermark-appeara.cs` | Save a watermarked Word document directly to PDF format while preserving watermark appearance. | watermark-workflow | doc, pdf | mcp |
| 23 | `apply-a-text-watermark-to-a-word-document-and-then-save-the-document-as-docx.cs` | Apply a text watermark to a Word document and then save the document as DOCX. | watermark-workflow | docx, doc | mcp |
| 24 | `apply-an-image-watermark-to-a-word-document-and-then-save-the-document-as-docx.cs` | Apply an image watermark to a Word document and then save the document as DOCX. | watermark-workflow | docx, doc | existing_repo |
| 25 | `batch-process-a-folder-of-doc-files-to-add-the-same-image-watermark-to-each-document.cs` | Batch process a folder of DOC files to add the same image watermark to each document. | watermark-workflow | doc | existing_repo |
| 26 | `batch-process-multiple-word-documents-in-a-directory-to-add-a-text-watermark-to-each-file.cs` | Batch process multiple Word documents in a directory to add a text watermark to each file. | watermark-workflow | doc | mcp |
| 27 | `batch-process-multiple-word-documents-in-a-directory-to-remove-existing-watermarks-from-ea.cs` | Batch process multiple Word documents in a directory to remove existing watermarks from each file. | watermark-workflow | doc | mcp |
| 28 | `batch-convert-docx-files-to-pdf-while-adding-a-corporate-logo-image-watermark-to-each-pdf.cs` | Batch convert DOCX files to PDF while adding a corporate logo image watermark to each PDF. | watermark-workflow | docx, doc, pdf | llm |
| 29 | `create-a-command-line-tool-that-accepts-a-directory-path-and-adds-a-specified-watermark-to.cs` | Create a command-line tool that accepts a directory path and adds a specified watermark to each file. | watermark-workflow | docx | mcp |
| 30 | `use-a-configuration-file-to-define-watermark-text-font-and-opacity-then-apply-it-to-multip.cs` | Use a configuration file to define watermark text, font, and opacity, then apply it to multiple documents. | watermark-workflow | doc | existing_repo |
| 31 | `implement-a-unit-test-that-verifies-watermark-remove-successfully-deletes-a-previously-add.cs` | Implement a unit test that verifies Watermark.Remove successfully deletes a previously added text watermark. | watermark-workflow | docx | llm |
| 32 | `add-a-watermark-to-a-document-opened-from-a-network-share-ensuring-proper-disposal-of-file.cs` | Add a watermark to a document opened from a network share, ensuring proper disposal of file handles. | watermark-workflow | doc | mcp |
| 33 | `create-a-reusable-method-that-adds-a-configurable-text-watermark-to-any-document-object.cs` | Create a reusable method that adds a configurable text watermark to any Document object. | watermark-workflow | doc | existing_repo |

## Common failure patterns seen during generation and how they were corrected

### Incorrect argument type for Watermark.SetImage

- Symptom: The compiler reports CS1503 because a string, FileStream, or MemoryStream is supplied where an SKBitmap is required, leading to runtime ArgumentNullException.
- Root cause: Code passes file paths or generic streams directly to SetImage instead of providing an SKBitmap instance or using the overload that accepts a file path.
- Preferred fix: Use the overload Document.Watermark.SetImage(string imagePath, ImageWatermarkOptions) or convert the stream to an SKBitmap before calling SetImage.
- Prevention: Always select the overload that matches the input type and avoid implicit conversions; consult the API reference for accepted parameter types.

### Missing table and cell types

- Symptom: CS0246 errors indicate that Table, Cell, Row, or related types cannot be resolved, causing build failures.
- Root cause: The required namespace Aspose.Words.Tables is omitted from the using directives.
- Preferred fix: Add "using Aspose.Words.Tables;" and, when needed, "using System.Drawing;" for color definitions.
- Prevention: Include all necessary Aspose.Words namespaces at the top of the file before referencing table-related classes.

### Invalid Watermark option types

- Symptom: CS0426 and CS0117 errors show that TextWatermarkOptions, ImageWatermarkOptions, or WatermarkLayout members are not found within the Watermark type.
- Root cause: The code treats Watermark as a namespace and attempts to access option types that belong to the Aspose.Words.Drawing namespace.
- Preferred fix: Reference Aspose.Words.Drawing and instantiate TextWatermarkOptions or ImageWatermarkOptions directly; use WatermarkLayout from the same namespace.
- Prevention: Follow the documented pattern: var options = new TextWatermarkOptions { FontFamily = "Arial", FontSize = 48, Color = Color.Blue, Layout = WatermarkLayout.Diagonal };

### Misuse of Watermark as a namespace and overbroad validation

- Symptom: CS0138 errors arise from using "using Watermark;" and validation code checks unrelated document internals.
- Root cause: Developers incorrectly add a using directive for the Watermark class and write validation that inspects unrelated nodes.
- Preferred fix: Remove the erroneous using directive and limit validation to checking the existence of the watermark and the output file.
- Prevention: Treat Watermark as a static class accessed via Document.Watermark and keep validation scoped to the task's output.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Update example images to a minimal embedded PNG to avoid external dependencies.
- Periodically verify that the Watermark API signatures have not changed in newer Aspose.Words releases.
- Ensure that all sample code includes explicit using directives for Aspose.Words.Drawing and Aspose.Words.Tables.
