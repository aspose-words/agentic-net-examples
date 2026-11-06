---
name: barcode-image
description: Verified C# examples for barcode-image scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - BarCode Image

## Purpose

Provide an AI coding assistant with clear guidance on creating, configuring, and rendering barcode fields in Aspose.Words documents while avoiding common compilation and runtime pitfalls, and ensuring that generated PDFs contain correctly rendered barcode images.

## Non-negotiable conventions

- Never reference Font, Color, or BarcodeParameters without their full Aspose.Drawing or Aspose.Words.Fields qualification.
- Do not use builder.InsertField with a raw DISPLAYBARCODE string or any string-based field switches.
- Never rely on the DISPLAYBARCODE field rendering path when the target format is PDF or an image.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Word Field Only workflow**: 5 examples
- **Custom Generator workflow**: 25 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Compile the sample with the exact package versions listed; any ambiguous type or missing member error must be resolved before runtime testing.

## File-to-task reference

- `create-a-new-document-insert-a-displaybarcode-field-then-save-the-document-as-docx.cs`
  - Task: Create a new Document, insert a DISPLAYBARCODE field, then save the document as DOCX.
  - Workflow: word-field-only
  - Outputs: docx
  - Selected engine: mcp
- `use-documentbuilder-to-insert-a-displaybarcode-field-with-a-datamatrix-barcode-type-switch.cs`
  - Task: Use DocumentBuilder to insert a DISPLAYBARCODE field with a DataMatrix barcode type switch.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `create-a-reusable-method-to-insert-a-displaybarcode-field-with-customizable-height-width-a.cs`
  - Task: Create a reusable method to insert a DISPLAYBARCODE field with customizable height, width, and type switches.
  - Workflow: word-field-only
  - Outputs: docx
  - Selected engine: mcp
- `create-a-macro-to-insert-displaybarcode-fields-with-predefined-switches-for-various-barcod.cs`
  - Task: Create a macro to insert DISPLAYBARCODE fields with predefined switches for various barcode types.
  - Workflow: word-field-only
  - Outputs: docx
  - Selected engine: mcp
- `configure-barcode-height-and-width-via-field-switches-in-the-displaybarcode-field-definiti.cs`
  - Task: Configure barcode height and width via field switches in the DISPLAYBARCODE field definition.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: llm
- `customize-barcode-color-and-background-via-additional-field-switches-and-verify-visual-app.cs`
  - Task: Customize barcode color and background via additional field switches and verify visual appearance in saved PDF.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `set-barcode-orientation-to-vertical-via-field-switches-and-verify-correct-rendering-in-pdf.cs`
  - Task: Set barcode orientation to vertical via field switches and verify correct rendering in PDF output.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: existing_repo
- `apply-different-barcode-types-to-separate-displaybarcode-fields-in-the-same-document-and-v.cs`
  - Task: Apply different barcode types to separate DISPLAYBARCODE fields in the same document and verify correct rendering.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `replace-placeholder-text-in-a-displaybarcode-field-with-dynamic-values-prior-to-barcode-ge.cs`
  - Task: Replace placeholder text in a DISPLAYBARCODE field with dynamic values prior to barcode generation.
  - Workflow: word-field-only
  - Outputs: docx
  - Selected engine: mcp
- `use-document-range-replace-to-update-the-data-string-of-an-existing-displaybarcode-field-b.cs`
  - Task: Use Document.Range.Replace to update the data string of an existing DISPLAYBARCODE field before rendering.
  - Workflow: word-field-only
  - Outputs: docx
  - Selected engine: mcp
- `implement-error-handling-for-missing-barcode-data-in-displaybarcode-fields-to-avoid-docume.cs`
  - Task: Implement error handling for missing barcode data in DISPLAYBARCODE fields to avoid document save failures.
  - Workflow: custom-generator
  - Outputs: docx
  - Selected engine: mcp
- `implement-the-ibarcodegenerator-interface-to-generate-code128-barcodes-from-field-data.cs`
  - Task: Implement the IBarcodeGenerator interface to generate Code128 barcodes from field data.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `integrate-the-aspose-barcode-library-to-enable-qr-code-generation-for-displaybarcode-field.cs`
  - Task: Integrate the Aspose.BarCode library to enable QR code generation for DISPLAYBARCODE fields.
  - Workflow: custom-generator
  - Outputs: docx
  - Selected engine: mcp
- `configure-the-custom-barcode-generator-to-cache-images-for-repeated-field-values-improving.cs`
  - Task: Configure the custom barcode generator to cache images for repeated field values, improving performance.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `configure-the-barcode-generator-to-produce-high-resolution-images-suitable-for-large-forma.cs`
  - Task: Configure the barcode generator to produce high-resolution images suitable for large-format PDF printing.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `generate-barcode-images-on-the-fly-during-save-by-assigning-the-custom-generator-to-docume.cs`
  - Task: Generate barcode images on the fly during save by assigning the custom generator to Document.BarcodeGenerator.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `load-an-rtf-document-assign-a-custom-generator-and-save-the-result-as-docx-with-barcodes.cs`
  - Task: Load an RTF document, assign a custom generator, and save the result as DOCX with barcodes.
  - Workflow: custom-generator
  - Outputs: docx
  - Selected engine: mcp
- `load-an-existing-docx-with-displaybarcode-fields-assign-a-custom-generator-and-export-to-p.cs`
  - Task: Load an existing DOCX with DISPLAYBARCODE fields, assign a custom generator, and export to PDF.
  - Workflow: custom-generator
  - Outputs: docx, pdf
  - Selected engine: mcp
- `load-a-docx-template-populate-displaybarcode-fields-with-user-data-and-export-the-document.cs`
  - Task: Load a DOCX template, populate DISPLAYBARCODE fields with user data, and export the document as PDF.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `batch-process-a-folder-of-doc-files-render-barcodes-and-save-each-document-as-pdf.cs`
  - Task: Batch process a folder of DOC files, render barcodes, and save each document as PDF.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `process-multiple-docx-files-in-parallel-each-with-its-own-barcode-generator-and-output-pdf.cs`
  - Task: Process multiple DOCX files in parallel, each with its own barcode generator, and output PDFs concurrently.
  - Workflow: custom-generator
  - Outputs: docx, pdf
  - Selected engine: mcp
- `create-a-console-application-that-accepts-a-directory-path-processes-supported-files-and-g.cs`
  - Task: Create a console application that accepts a directory path, processes supported files, and generates PDFs with barcodes.
  - Workflow: custom-generator
  - Outputs: docx, pdf
  - Selected engine: mcp
- `save-a-document-with-displaybarcode-fields-as-rtf-ensuring-barcodes-render-as-images-in-ou.cs`
  - Task: Save a document with DISPLAYBARCODE fields as RTF, ensuring barcodes render as images in output.
  - Workflow: custom-generator
  - Outputs: rtf
  - Selected engine: mcp
- `validate-barcode-images-are-correctly-embedded-in-pdf-by-extracting-them-and-comparing-dim.cs`
  - Task: Validate barcode images are correctly embedded in PDF by extracting them and comparing dimensions.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `validate-that-barcode-images-maintain-correct-aspect-ratio-after-converting-a-document-fro.cs`
  - Task: Validate that barcode images maintain correct aspect ratio after converting a document from DOC to PDF.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `test-barcode-rendering-when-saving-a-document-to-pdf-a-format-ensuring-archival-compliance.cs`
  - Task: Test barcode rendering when saving a document to PDF/A format, ensuring archival compliance.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `add-a-logging-mechanism-to-record-each-barcode-generation-event-with-field-name-and-image.cs`
  - Task: Add a logging mechanism to record each barcode generation event with field name and image size.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `create-a-unit-test-that-loads-a-doc-file-renders-barcodes-and-asserts-pdf-output-contains.cs`
  - Task: Create a unit test that loads a DOC file, renders barcodes, and asserts PDF output contains expected images.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `generate-barcodes-with-variable-widths-based-on-input-string-length-by-programmatically-ad.cs`
  - Task: Generate barcodes with variable widths based on input string length by programmatically adjusting field switch parameters.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `implement-a-feature-to-disable-barcode-rendering-for-specific-fields-during-pdf-export-whi.cs`
  - Task: Implement a feature to disable barcode rendering for specific fields during PDF export while retaining them in DOCX.
  - Workflow: custom-generator
  - Outputs: docx, pdf
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Ambiguous type references**
  - Symptom: Compiler reports an ambiguous reference for Font or BarcodeParameters.
  - Root cause: Both Aspose.Drawing and Aspose.Words (or Aspose.BarCode) define types with the same name, and the code uses an unqualified identifier.
  - Preferred fix: Qualify the type with its full namespace (Aspose.Drawing.Font, Aspose.Words.Fields.BarcodeParameters) everywhere barcode code is written.
  - Prevention: Avoid adding broad using directives for conflicting namespaces; always use fully qualified names for Font, Color, and BarcodeParameters.

- **Use of non-existent FieldDisplayBarcode members**
  - Symptom: CS1061 errors such as "FieldDisplayBarcode does not contain a definition for 'Height'" or "'Switches'".
  - Root cause: Legacy examples rely on field-code string switches that are not exposed as properties on the typed FieldDisplayBarcode class.
  - Preferred fix: Set only the supported typed properties (BarcodeType, BarcodeValue, ForegroundColor, BackgroundColor, ScalingFactor, SymbolRotation) and let the custom generator handle size and orientation.
  - Prevention: Consult the Aspose.Words API reference for FieldDisplayBarcode and remove any code that accesses undefined members.

- **Incorrect namespace usage**
  - Symptom: CS0234 errors indicating that namespaces such as Aspose.Words.Building, Aspose.Words.Nodes, or Aspose.Words.Builder do not exist.
  - Root cause: Typographical errors or outdated documentation introduced invalid namespace names.
  - Preferred fix: Replace the incorrect namespaces with the correct ones (e.g., Aspose.Words.Loading, Aspose.Words.Saving, Aspose.Words.Replacing).
  - Prevention: Reference the official Aspose.Words documentation for the exact namespace hierarchy before adding using directives.

- **Invalid EncodeTypes enum members**
  - Symptom: CS0117 or CS0722 errors stating that EncodeTypes does not contain definitions such as Code39Standard or that EncodeTypes is used as a static type.
  - Root cause: The code uses obsolete or non-existent enum members from an older version of Aspose.BarCode.
  - Preferred fix: Use the current enum members (e.g., EncodeTypes.Code39, EncodeTypes.QR, SymbologyEncodeType) that exist in version 26.10.0.
  - Prevention: Verify enum members against the installed Aspose.BarCode package version and avoid hard-coding values that are not present.

- **Missing input assets at runtime**
  - Symptom: FileNotFoundException or DirectoryNotFoundException when loading template.docx, input.docx, or processing a folder of documents.
  - Root cause: Sample code assumes external files exist in the execution directory without creating or checking them.
  - Preferred fix: Create temporary sample files programmatically before loading, or add existence checks and fallback sample content.
  - Prevention: Always bootstrap required assets in examples or guard file-system operations with File.Exists / Directory.Exists checks.

- **Transient MCP gateway timeout**
  - Symptom: 504 Gateway Time-out errors during generation.
  - Root cause: Intermittent network or service availability issues unrelated to the source code.
  - Preferred fix: Retry the generation or allow the alternative engine to produce the output when the primary engine fails.
  - Prevention: Implement retry logic with exponential back-off for network-bound operations; this does not require code changes to the barcode logic.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.10.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
dotnet add package Aspose.BarCode --version 26.10.0
dotnet add package Aspose.Drawing.Common --version 26.10.0
```

```powershell
Copy-Item ..\barcode-image\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update the fully qualified type references if future Aspose releases introduce new conflicting namespaces.
- Periodically verify that the custom IBarcodeGenerator implementation still matches the latest Aspose.BarCode API signatures.
