---
name: watermark
description: Verified C# examples for watermark scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Watermark

## Purpose

Provide an AI coding agent with clear guidance to generate reliable watermark code that respects Aspose.Words APIs, creates any required sample assets, and validates only the presence of the expected watermark output.

## Non-negotiable conventions

- Do not reference members that are not present in Aspose.Words 26.11.0, such as WatermarkLayout members, ImageWatermarkOptions inside the Watermark type, or System.Drawing APIs for image generation.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Watermark Workflow workflow**: 33 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Confirm that the output file exists and that Document.Watermark.ContainsWatermark returns true for the expected WatermarkType.

## File-to-task reference

- `load-a-word-document-from-a-file-path-and-add-a-text-watermark-using-watermark-settext.cs`
  - Task: Load a Word document from a file path and add a text watermark using Watermark.SetText.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `load-a-word-document-from-a-file-path-and-add-an-image-watermark-using-watermark-setimage.cs`
  - Task: Load a Word document from a file path and add an image watermark using Watermark.SetImage.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `load-a-word-document-from-a-memory-stream-and-apply-a-text-watermark-without-writing-to-di.cs`
  - Task: Load a Word document from a memory stream and apply a text watermark without writing to disk.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `insert-a-text-watermark-into-a-specific-table-cell-within-a-word-document-using-the-waterm.cs`
  - Task: Insert a text watermark into a specific table cell within a Word document using the Watermark class.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `add-a-watermark-to-a-table-cell-that-spans-multiple-rows-and-columns-in-a-complex-word-tab.cs`
  - Task: Add a watermark to a table cell that spans multiple rows and columns in a complex Word table.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-watermark-into-a-table-cell-that-contains-merged-cells-without-disrupting-the-tab.cs`
  - Task: Insert a watermark into a table cell that contains merged cells without disrupting the table layout.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-watermark-into-each-cell-of-the-first-row-of-a-table-using-the-watermark-class.cs`
  - Task: Insert a watermark into each cell of the first row of a table using the Watermark class.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: mcp
- `add-a-text-watermark-to-a-docx-document-using-watermark-settext-with-custom-font-settings.cs`
  - Task: Add a text watermark to a DOCX document using Watermark.SetText with custom font settings.
  - Workflow: watermark-workflow
  - Outputs: docx, doc
  - Selected engine: existing_repo
- `use-watermark-settext-with-textwatermarkoptions-to-set-watermark-font-size-color-and-spaci.cs`
  - Task: Use Watermark.SetText with TextWatermarkOptions to set watermark font size, color, and spacing.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: existing_repo
- `add-a-confidential-text-watermark-to-all-new-documents-created-by-an-automated-report-gene.cs`
  - Task: Add a confidential text watermark to all new documents created by an automated report generator.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `combine-text-and-image-watermarks-by-first-setting-a-text-watermark-then-overlaying-an-ima.cs`
  - Task: Combine text and image watermarks by first setting a text watermark then overlaying an image watermark.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: existing_repo
- `customize-image-watermark-opacity-and-scaling-by-configuring-imagewatermarkoptions-before.cs`
  - Task: Customize image watermark opacity and scaling by configuring ImageWatermarkOptions before insertion.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: existing_repo
- `insert-an-image-watermark-from-a-file-path-into-a-word-document-after-optimizing-the-docum.cs`
  - Task: Insert an image watermark from a file path into a Word document after optimizing the document.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `use-watermark-setimage-with-a-stream-to-embed-a-logo-watermark-into-a-document-stored-in-a.cs`
  - Task: Use Watermark.SetImage with a stream to embed a logo watermark into a document stored in Azure Blob storage.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `use-watermark-setimage-with-a-byte-array-stream-to-embed-a-dynamically-generated-barcode-w.cs`
  - Task: Use Watermark.SetImage with a byte array stream to embed a dynamically generated barcode watermark.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: existing_repo
- `optimize-a-large-docx-file-before-applying-an-image-watermark-to-improve-performance-and-m.cs`
  - Task: Optimize a large DOCX file before applying an image watermark to improve performance and memory usage.
  - Workflow: watermark-workflow
  - Outputs: docx, doc
  - Selected engine: existing_repo
- `remove-all-existing-watermarks-from-a-loaded-word-document-using-the-watermark-remove-meth.cs`
  - Task: Remove all existing watermarks from a loaded Word document using the Watermark.Remove method.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-utility-method-that-removes-all-watermarks-from-a-document-using-watermark-remove.cs`
  - Task: Create a utility method that removes all watermarks from a Document using Watermark.Remove.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `validate-that-a-document-contains-no-watermarks-before-publishing-by-using-watermarktype-n.cs`
  - Task: Validate that a document contains no watermarks before publishing by using WatermarkType.None check.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-watermarktype-enumeration-to-verify-a-document-has-no-watermark-before-adding-a-new-on.cs`
  - Task: Use WatermarkType enumeration to verify a document has no watermark before adding a new one.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-watermarktype-enumeration-to-switch-between-text-and-image-watermarks-based-on-user-se.cs`
  - Task: Use WatermarkType enumeration to switch between text and image watermarks based on user selection.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: existing_repo
- `save-a-watermarked-word-document-directly-to-pdf-format-while-preserving-watermark-appeara.cs`
  - Task: Save a watermarked Word document directly to PDF format while preserving watermark appearance.
  - Workflow: watermark-workflow
  - Outputs: doc, pdf
  - Selected engine: mcp
- `apply-a-text-watermark-to-a-word-document-and-then-save-the-document-as-docx.cs`
  - Task: Apply a text watermark to a Word document and then save the document as DOCX.
  - Workflow: watermark-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `apply-an-image-watermark-to-a-word-document-and-then-save-the-document-as-docx.cs`
  - Task: Apply an image watermark to a Word document and then save the document as DOCX.
  - Workflow: watermark-workflow
  - Outputs: docx, doc
  - Selected engine: existing_repo
- `batch-process-a-folder-of-doc-files-to-add-the-same-image-watermark-to-each-document.cs`
  - Task: Batch process a folder of DOC files to add the same image watermark to each document.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `batch-process-multiple-word-documents-in-a-directory-to-add-a-text-watermark-to-each-file.cs`
  - Task: Batch process multiple Word documents in a directory to add a text watermark to each file.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `batch-process-multiple-word-documents-in-a-directory-to-remove-existing-watermarks-from-ea.cs`
  - Task: Batch process multiple Word documents in a directory to remove existing watermarks from each file.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `batch-convert-docx-files-to-pdf-while-adding-a-corporate-logo-image-watermark-to-each-pdf.cs`
  - Task: Batch convert DOCX files to PDF while adding a corporate logo image watermark to each PDF.
  - Workflow: watermark-workflow
  - Outputs: docx, doc, pdf
  - Selected engine: llm
- `create-a-command-line-tool-that-accepts-a-directory-path-and-adds-a-specified-watermark-to.cs`
  - Task: Create a command-line tool that accepts a directory path and adds a specified watermark to each file.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: mcp
- `use-a-configuration-file-to-define-watermark-text-font-and-opacity-then-apply-it-to-multip.cs`
  - Task: Use a configuration file to define watermark text, font, and opacity, then apply it to multiple documents.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `implement-a-unit-test-that-verifies-watermark-remove-successfully-deletes-a-previously-add.cs`
  - Task: Implement a unit test that verifies Watermark.Remove successfully deletes a previously added text watermark.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: llm
- `add-a-watermark-to-a-document-opened-from-a-network-share-ensuring-proper-disposal-of-file.cs`
  - Task: Add a watermark to a document opened from a network share, ensuring proper disposal of file handles.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-reusable-method-that-adds-a-configurable-text-watermark-to-any-document-object.cs`
  - Task: Create a reusable method that adds a configurable text watermark to any Document object.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: existing_repo

## Common failure patterns and preferred agent fixes

- **Incorrect argument type for Watermark.SetImage**
  - Symptom: The compiler reports CS1503 because a string, FileStream, or MemoryStream is supplied where an SKBitmap is required, leading to runtime ArgumentNullException.
  - Root cause: Code passes file paths or generic streams directly to SetImage instead of providing an SKBitmap instance or using the overload that accepts a file path.
  - Preferred fix: Use the overload Document.Watermark.SetImage(string imagePath, ImageWatermarkOptions) or convert the stream to an SKBitmap before calling SetImage.
  - Prevention: Always select the overload that matches the input type and avoid implicit conversions; consult the API reference for accepted parameter types.

- **Missing table and cell types**
  - Symptom: CS0246 errors indicate that Table, Cell, Row, or related types cannot be resolved, causing build failures.
  - Root cause: The required namespace Aspose.Words.Tables is omitted from the using directives.
  - Preferred fix: Add "using Aspose.Words.Tables;" and, when needed, "using System.Drawing;" for color definitions.
  - Prevention: Include all necessary Aspose.Words namespaces at the top of the file before referencing table-related classes.

- **Invalid Watermark option types**
  - Symptom: CS0426 and CS0117 errors show that TextWatermarkOptions, ImageWatermarkOptions, or WatermarkLayout members are not found within the Watermark type.
  - Root cause: The code treats Watermark as a namespace and attempts to access option types that belong to the Aspose.Words.Drawing namespace.
  - Preferred fix: Reference Aspose.Words.Drawing and instantiate TextWatermarkOptions or ImageWatermarkOptions directly; use WatermarkLayout from the same namespace.
  - Prevention: Follow the documented pattern: var options = new TextWatermarkOptions { FontFamily = "Arial", FontSize = 48, Color = Color.Blue, Layout = WatermarkLayout.Diagonal };

- **Misuse of Watermark as a namespace and overbroad validation**
  - Symptom: CS0138 errors arise from using "using Watermark;" and validation code checks unrelated document internals.
  - Root cause: Developers incorrectly add a using directive for the Watermark class and write validation that inspects unrelated nodes.
  - Preferred fix: Remove the erroneous using directive and limit validation to checking the existence of the watermark and the output file.
  - Prevention: Treat Watermark as a static class accessed via Document.Watermark and keep validation scoped to the task's output.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
dotnet add package SkiaSharp --version 4.152.1
```

```powershell
Copy-Item ..\watermark\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update example images to a minimal embedded PNG to avoid external dependencies.
- Periodically verify that the Watermark API signatures have not changed in newer Aspose.Words releases.
- Ensure that all sample code includes explicit using directives for Aspose.Words.Drawing and Aspose.Words.Tables.
