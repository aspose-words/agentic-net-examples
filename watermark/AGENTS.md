---
name: watermark
description: Verified C# examples for watermark workflows in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Watermark

## Purpose

This folder is a **live, curated example set** for watermark scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free use of documented Aspose.Words APIs for inserting, removing, replacing, validating, and exporting text and image watermarks.

## Non-negotiable conventions

- Always use documented Aspose.Words APIs directly.
- Always create local sample source documents or images when a task refers to an existing file, folder, stream, template, or input asset.
- Prefer `Document.Watermark` APIs when they directly fit the task.
- Keep validation narrow and task-specific.
- Do not invent watermark helper APIs or unsupported WordArt/TextPath members.

## Recommended workflow selection

- **Watermark workflow**: 34 examples

This category performed best with light primary rules plus narrow patches for image generation, cell-scoped watermarks, and stream-based image workflows.

## Validation priorities

1. The code must compile and run without manual input.
2. Required sample inputs must be bootstrapped locally inside the example.
3. Requested watermark or output files must be produced successfully.
4. Validation should focus only on the exact requested watermark presence, type, replacement, removal, cell scope, or saved output.

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
  - Selected engine: mcp
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
  - Selected engine: mcp
- `use-watermark-settext-with-textwatermarkoptions-to-set-watermark-font-size-color-and-spaci.cs`
  - Task: Use Watermark.SetText with TextWatermarkOptions to set watermark font size, color, and spacing.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: mcp
- `add-a-confidential-text-watermark-to-all-new-documents-created-by-an-automated-report-gene.cs`
  - Task: Add a confidential text watermark to all new documents created by an automated report generator.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `combine-text-and-image-watermarks-by-first-setting-a-text-watermark-then-overlaying-an-ima.cs`
  - Task: Combine text and image watermarks by first setting a text watermark then overlaying an image watermark.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: mcp
- `customize-image-watermark-opacity-and-scaling-by-configuring-imagewatermarkoptions-before.cs`
  - Task: Customize image watermark opacity and scaling by configuring ImageWatermarkOptions before insertion.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-an-image-watermark-from-a-system-drawing-image-object-into-a-document-after-calling.cs`
  - Task: Insert an image watermark from a System.Drawing.Image object into a document after calling Optimize.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `insert-an-image-watermark-from-a-file-path-into-a-word-document-after-optimizing-the-docum.cs`
  - Task: Insert an image watermark from a file path into a Word document after optimizing the document.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-watermark-setimage-with-a-stream-to-embed-a-logo-watermark-into-a-document-stored-in-a.cs`
  - Task: Use Watermark.SetImage with a stream to embed a logo watermark into a document stored in Azure Blob storage.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-watermark-setimage-with-a-byte-array-stream-to-embed-a-dynamically-generated-barcode-w.cs`
  - Task: Use Watermark.SetImage with a byte array stream to embed a dynamically generated barcode watermark.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: mcp
- `optimize-a-large-docx-file-before-applying-an-image-watermark-to-improve-performance-and-m.cs`
  - Task: Optimize a large DOCX file before applying an image watermark to improve performance and memory usage.
  - Workflow: watermark-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
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
  - Selected engine: mcp
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
  - Selected engine: mcp
- `batch-process-a-folder-of-doc-files-to-add-the-same-image-watermark-to-each-document.cs`
  - Task: Batch process a folder of DOC files to add the same image watermark to each document.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
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
  - Selected engine: mcp
- `create-a-command-line-tool-that-accepts-a-directory-path-and-adds-a-specified-watermark-to.cs`
  - Task: Create a command‑line tool that accepts a directory path and adds a specified watermark to each file.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: mcp
- `use-a-configuration-file-to-define-watermark-text-font-and-opacity-then-apply-it-to-multip.cs`
  - Task: Use a configuration file to define watermark text, font, and opacity, then apply it to multiple documents.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `implement-a-unit-test-that-verifies-watermark-remove-successfully-deletes-a-previously-add.cs`
  - Task: Implement a unit test that verifies Watermark.Remove successfully deletes a previously added text watermark.
  - Workflow: watermark-workflow
  - Outputs: docx
  - Selected engine: mcp
- `add-a-watermark-to-a-document-opened-from-a-network-share-ensuring-proper-disposal-of-file.cs`
  - Task: Add a watermark to a document opened from a network share, ensuring proper disposal of file handles.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-reusable-method-that-adds-a-configurable-text-watermark-to-any-document-object.cs`
  - Task: Create a reusable method that adds a configurable text watermark to any Document object.
  - Workflow: watermark-workflow
  - Outputs: doc
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Unsupported image generation**
  - Symptom: Build failures caused by `System.Drawing` APIs in the verifier environment.
  - Preferred fix: Use a compile-safe local image file or stream instead of drawing the image at runtime with unsupported APIs.
- **Incorrect Watermark namespace assumptions**
  - Symptom: Build failures from `using Watermark;`.
  - Preferred fix: Access watermark features through `Document.Watermark`.
- **Unsupported cell-level WordArt approaches**
  - Symptom: Build failures or unsupported runtime behavior from WordArt or unsupported `TextPath` members.
  - Preferred fix: Use compile-safe shape or image insertion inside the target cell or merged-cell scope.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required package

```bash
dotnet add package Aspose.Words --version 26.3.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\watermark\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve exact file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the verified winner from the latest batch report rather than a merely compiling draft.
- Bootstrap file-based inputs locally instead of depending on machine-specific paths.
