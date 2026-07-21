---
name: watermark
description: Verified C# examples for watermark scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Watermark

## Purpose

This folder is a live, curated example set for Watermark scenarios. Treat every `.cs` file as a standalone console application. The goal is correct, warning-free examples that use documented Aspose APIs and match the original task intent.

## Non-negotiable conventions

- Always use documented Aspose.Words APIs directly.
- Always create local sample source documents or images when a task refers to an existing file, folder, stream, template, or input asset.
- Prefer `Document.Watermark` APIs when they directly fit the task.
- Keep validation narrow and task-specific.
- Do not invent watermark helper APIs or unsupported WordArt/TextPath members.

## Recommended workflow selection

- Watermark Workflow workflow: 33 examples
- Watermark Api workflow: 1 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. API usage must be supported by the configured package versions.
3. Exported outputs must actually be written by the example.
4. Validation scenarios must inspect only the behavior requested by the task.
5. Examples that depend on files, folders, streams, images, or data should bootstrap those inputs locally during the example run.

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
  - Workflow: watermark-api
  - Outputs: image
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
  - Selected engine: mcp
- `create-a-command-line-tool-that-accepts-a-directory-path-and-adds-a-specified-watermark-to.cs`
  - Task: Create a command-line tool that accepts a directory path and adds a specified watermark to each file.
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

- **Unsupported API invention**
  - Symptom: Generated code references members that do not exist in the selected package version.
  - Preferred fix: Replace invented members with documented Aspose.Words APIs already proven in this category.

- **Missing local bootstrap inputs**
  - Symptom: The example assumes source files, folders, images, or data already exist.
  - Preferred fix: Create deterministic local inputs before loading, processing, or validating them.

- **Over-broad validation**
  - Symptom: The example fails at runtime while checking unrelated document internals.
  - Preferred fix: Validate only the requested behavior and the existence of expected outputs.

## Build and run contract

- Target framework: `net8.0`
- Package: `Aspose.Words` `26.6.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.6.0
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

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer documented Aspose APIs over speculative shortcuts.
