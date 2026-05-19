# Watermark Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **Watermark** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Watermark**
- Slug: **watermark**
- Total examples: **34**
- Workflow examples: **34 / 34** use the standard watermark workflow

## Category rules that shaped these examples

- Use native Aspose.Words APIs directly.
- Create local sample source documents or images when a task refers to an existing file, folder, stream, template, or input asset.
- Do not assume external files or folders already exist.
- Prefer documented `Document.Watermark` APIs when they directly fit the task.
- Keep validation narrow and task-specific.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\watermark\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# PowerShell example
Copy-Item ..\watermark\load-a-word-document-from-a-file-path-and-add-a-text-watermark-using-watermark-settext.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-word-document-from-a-file-path-and-add-a-text-watermark-using-watermark-settext.cs` | Load a Word document from a file path and add a text watermark using Watermark.SetText. | watermark-workflow | doc | mcp |
| 2 | `load-a-word-document-from-a-file-path-and-add-an-image-watermark-using-watermark-setimage.cs` | Load a Word document from a file path and add an image watermark using Watermark.SetImage. | watermark-workflow | doc | mcp |
| 3 | `load-a-word-document-from-a-memory-stream-and-apply-a-text-watermark-without-writing-to-di.cs` | Load a Word document from a memory stream and apply a text watermark without writing to disk. | watermark-workflow | doc | mcp |
| 4 | `insert-a-text-watermark-into-a-specific-table-cell-within-a-word-document-using-the-waterm.cs` | Insert a text watermark into a specific table cell within a Word document using the Watermark class. | watermark-workflow | doc | mcp |
| 5 | `add-a-watermark-to-a-table-cell-that-spans-multiple-rows-and-columns-in-a-complex-word-tab.cs` | Add a watermark to a table cell that spans multiple rows and columns in a complex Word table. | watermark-workflow | docx | mcp |
| 6 | `insert-a-watermark-into-a-table-cell-that-contains-merged-cells-without-disrupting-the-tab.cs` | Insert a watermark into a table cell that contains merged cells without disrupting the table layout. | watermark-workflow | docx | mcp |
| 7 | `insert-a-watermark-into-each-cell-of-the-first-row-of-a-table-using-the-watermark-class.cs` | Insert a watermark into each cell of the first row of a table using the Watermark class. | watermark-workflow | docx | mcp |
| 8 | `add-a-text-watermark-to-a-docx-document-using-watermark-settext-with-custom-font-settings.cs` | Add a text watermark to a DOCX document using Watermark.SetText with custom font settings. | watermark-workflow | docx, doc | mcp |
| 9 | `use-watermark-settext-with-textwatermarkoptions-to-set-watermark-font-size-color-and-spaci.cs` | Use Watermark.SetText with TextWatermarkOptions to set watermark font size, color, and spacing. | watermark-workflow | docx | mcp |
| 10 | `add-a-confidential-text-watermark-to-all-new-documents-created-by-an-automated-report-gene.cs` | Add a confidential text watermark to all new documents created by an automated report generator. | watermark-workflow | doc | mcp |
| 11 | `combine-text-and-image-watermarks-by-first-setting-a-text-watermark-then-overlaying-an-ima.cs` | Combine text and image watermarks by first setting a text watermark then overlaying an image watermark. | watermark-workflow | docx | mcp |
| 12 | `customize-image-watermark-opacity-and-scaling-by-configuring-imagewatermarkoptions-before.cs` | Customize image watermark opacity and scaling by configuring ImageWatermarkOptions before insertion. | watermark-workflow | docx | mcp |
| 13 | `insert-an-image-watermark-from-a-system-drawing-image-object-into-a-document-after-calling.cs` | Insert an image watermark from a System.Drawing.Image object into a document after calling Optimize. | watermark-workflow | doc | mcp |
| 14 | `insert-an-image-watermark-from-a-file-path-into-a-word-document-after-optimizing-the-docum.cs` | Insert an image watermark from a file path into a Word document after optimizing the document. | watermark-workflow | doc | mcp |
| 15 | `use-watermark-setimage-with-a-stream-to-embed-a-logo-watermark-into-a-document-stored-in-a.cs` | Use Watermark.SetImage with a stream to embed a logo watermark into a document stored in Azure Blob storage. | watermark-workflow | doc | mcp |
| 16 | `use-watermark-setimage-with-a-byte-array-stream-to-embed-a-dynamically-generated-barcode-w.cs` | Use Watermark.SetImage with a byte array stream to embed a dynamically generated barcode watermark. | watermark-workflow | docx | mcp |
| 17 | `optimize-a-large-docx-file-before-applying-an-image-watermark-to-improve-performance-and-m.cs` | Optimize a large DOCX file before applying an image watermark to improve performance and memory usage. | watermark-workflow | docx, doc | mcp |
| 18 | `remove-all-existing-watermarks-from-a-loaded-word-document-using-the-watermark-remove-meth.cs` | Remove all existing watermarks from a loaded Word document using the Watermark.Remove method. | watermark-workflow | doc | mcp |
| 19 | `create-a-utility-method-that-removes-all-watermarks-from-a-document-using-watermark-remove.cs` | Create a utility method that removes all watermarks from a Document using Watermark.Remove. | watermark-workflow | doc | mcp |
| 20 | `validate-that-a-document-contains-no-watermarks-before-publishing-by-using-watermarktype-n.cs` | Validate that a document contains no watermarks before publishing by using WatermarkType.None check. | watermark-workflow | doc | mcp |
| 21 | `use-watermarktype-enumeration-to-verify-a-document-has-no-watermark-before-adding-a-new-on.cs` | Use WatermarkType enumeration to verify a document has no watermark before adding a new one. | watermark-workflow | doc | mcp |
| 22 | `use-watermarktype-enumeration-to-switch-between-text-and-image-watermarks-based-on-user-se.cs` | Use WatermarkType enumeration to switch between text and image watermarks based on user selection. | watermark-workflow | docx | mcp |
| 23 | `save-a-watermarked-word-document-directly-to-pdf-format-while-preserving-watermark-appeara.cs` | Save a watermarked Word document directly to PDF format while preserving watermark appearance. | watermark-workflow | doc, pdf | mcp |
| 24 | `apply-a-text-watermark-to-a-word-document-and-then-save-the-document-as-docx.cs` | Apply a text watermark to a Word document and then save the document as DOCX. | watermark-workflow | docx, doc | mcp |
| 25 | `apply-an-image-watermark-to-a-word-document-and-then-save-the-document-as-docx.cs` | Apply an image watermark to a Word document and then save the document as DOCX. | watermark-workflow | docx, doc | mcp |
| 26 | `batch-process-a-folder-of-doc-files-to-add-the-same-image-watermark-to-each-document.cs` | Batch process a folder of DOC files to add the same image watermark to each document. | watermark-workflow | doc | mcp |
| 27 | `batch-process-multiple-word-documents-in-a-directory-to-add-a-text-watermark-to-each-file.cs` | Batch process multiple Word documents in a directory to add a text watermark to each file. | watermark-workflow | doc | mcp |
| 28 | `batch-process-multiple-word-documents-in-a-directory-to-remove-existing-watermarks-from-ea.cs` | Batch process multiple Word documents in a directory to remove existing watermarks from each file. | watermark-workflow | doc | mcp |
| 29 | `batch-convert-docx-files-to-pdf-while-adding-a-corporate-logo-image-watermark-to-each-pdf.cs` | Batch convert DOCX files to PDF while adding a corporate logo image watermark to each PDF. | watermark-workflow | docx, doc, pdf | mcp |
| 30 | `create-a-command-line-tool-that-accepts-a-directory-path-and-adds-a-specified-watermark-to.cs` | Create a command‑line tool that accepts a directory path and adds a specified watermark to each file. | watermark-workflow | docx | mcp |
| 31 | `use-a-configuration-file-to-define-watermark-text-font-and-opacity-then-apply-it-to-multip.cs` | Use a configuration file to define watermark text, font, and opacity, then apply it to multiple documents. | watermark-workflow | doc | mcp |
| 32 | `implement-a-unit-test-that-verifies-watermark-remove-successfully-deletes-a-previously-add.cs` | Implement a unit test that verifies Watermark.Remove successfully deletes a previously added text watermark. | watermark-workflow | docx | mcp |
| 33 | `add-a-watermark-to-a-document-opened-from-a-network-share-ensuring-proper-disposal-of-file.cs` | Add a watermark to a document opened from a network share, ensuring proper disposal of file handles. | watermark-workflow | doc | mcp |
| 34 | `create-a-reusable-method-that-adds-a-configurable-text-watermark-to-any-document-object.cs` | Create a reusable method that adds a configurable text watermark to any Document object. | watermark-workflow | doc | mcp |

## Common failure patterns seen during generation and how they were corrected

### Using unsupported System.Drawing-based image generation

- Symptom: Build failures caused by `Bitmap`, `Graphics`, or other `System.Drawing` APIs not available in the verifier environment.
- Fix: Use compile-safe local image files or streams instead of System.Drawing drawing logic.

### Treating Watermark as a namespace

- Symptom: Build failures caused by lines such as `using Watermark;`.
- Fix: Access watermark functionality through `Document.Watermark` only.

### Inventing unsupported WordArt or TextPath APIs for cell-level watermarks

- Symptom: Build failures caused by `TextPath.FontSize`, `TextPath.FillColor`, `ShapeType.WordArt`, or unsupported positioning enums.
- Fix: Use compile-safe shape or image workflows inside the target cell rather than unsupported WordArt-style approaches.

## Notes for maintainers

- The selected file for each task is the verified winner recorded in the batch run.
- This category performed best with light primary rules plus narrow patches for image generation, cell-scoped watermark handling, and environment-specific stream workflows.
- Preserve exact file-to-task traceability when updating the category.
- Bootstrap all sample input files locally inside the example when the task refers to an existing asset.
