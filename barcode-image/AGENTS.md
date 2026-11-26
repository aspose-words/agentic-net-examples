---
name: barcode-image
description: Verified C# examples for barcode-image scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - BarCode Image

## Purpose

Guide an AI coding agent to generate, customize, and validate barcode fields in Word documents while adhering to strict API usage rules and employing a custom barcode generator for rendered outputs such as PDF.

## Non-negotiable conventions

- Never use unqualified BarcodeParameters; always reference Aspose.Words.Fields.BarcodeParameters.
- Never use unqualified Font; always reference Aspose.Drawing.Font.
- Never use unqualified Color; always reference Aspose.Drawing.Color.
- Never import System.Drawing types in this category.
- Do not insert DISPLAYBARCODE fields using string syntax or builder.InsertField with a raw field code.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Word Field Only workflow**: 5 examples
- **Custom Generator workflow**: 25 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. After updating fields, verify that the generated PDF contains barcode images of the expected type and dimensions.

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
  - Selected engine: llm
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
  - Selected engine: llm
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
  - Selected engine: llm
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

- **Ambiguous type reference**
  - Symptom: Compiler reports an ambiguous reference for 'Font' or 'BarcodeParameters' between Aspose.Drawing and Aspose.Words namespaces.
  - Root cause: Unqualified type names cause the C# compiler to see multiple matching types from different assemblies.
  - Preferred fix: Qualify the type with its full namespace (e.g., Aspose.Drawing.Font, Aspose.Words.Fields.BarcodeParameters).
  - Prevention: Always use fully qualified names or explicit using aliases for these types.

- **Invalid FieldDisplayBarcode member usage**
  - Symptom: CS1061 errors for Height, Width, SetSwitch, SetFieldCode, or Orientation members that do not exist on FieldDisplayBarcode.
  - Root cause: Attempting to control barcode appearance through legacy field switches instead of the typed API.
  - Preferred fix: Use supported typed properties such as ScalingFactor, ForegroundColor, BackgroundColor, and rely on a custom generator for size adjustments.
  - Prevention: Do not call non-existent members; consult the API reference for available FieldDisplayBarcode properties.

- **EncodeTypes enum misuse**
  - Symptom: CS0117 or CS0723 errors when assigning SymbologyEncodeType values to EncodeTypes or using undefined EncodeTypes members.
  - Root cause: Confusing the static EncodeTypes class with the SymbologyEncodeType enum used by Aspose.BarCode.
  - Preferred fix: Map barcode type strings to SymbologyEncodeType values (e.g., EncodeTypes.Code128) and avoid treating EncodeTypes as an enum.
  - Prevention: Always use SymbologyEncodeType for generator construction and keep the mapping logic isolated.

- **Missing Aspose.Words namespace**
  - Symptom: CS0234 errors for namespaces such as Aspose.Words.Building, Aspose.Words.Nodes, or Aspose.Words.LoadOptions.
  - Root cause: Incorrect or outdated using directives referencing non-existent namespaces in the installed package version.
  - Preferred fix: Remove invalid namespaces and import only existing ones like Aspose.Words.Drawing and Aspose.Words.Replacing.
  - Prevention: Verify namespace availability against the current Aspose.Words version before adding using statements.

- **Missing input file or directory**
  - Symptom: Runtime FileNotFoundException or DirectoryNotFoundException when loading template documents or processing folders.
  - Root cause: Example code assumes external assets without creating or checking them.
  - Preferred fix: Create temporary sample files/folders within the example or guard file-system operations with existence checks.
  - Prevention: Ensure all required assets are provisioned programmatically before they are accessed.

- **Undefined ImageFormat**
  - Symptom: CS0103 error indicating that 'ImageFormat' does not exist in the current context.
  - Root cause: Missing reference to Aspose.Drawing.Imaging or accidental reliance on System.Drawing types.
  - Preferred fix: Add 'using Aspose.Drawing.Imaging;' and reference Aspose.Drawing.Imaging.ImageFormat.
  - Prevention: Always import the Aspose.Drawing.Imaging namespace when working with image formats.

- **BarcodeParameters property mismatch**
  - Symptom: CS1061 errors for missing members such as BarcodeSymbology, BarcodeForeColor, or FieldName on BarcodeParameters.
  - Root cause: Using the wrong BarcodeParameters type (Aspose.BarCode.Generation) or assuming properties that are not part of Aspose.Words.Fields.BarcodeParameters.
  - Preferred fix: Reference Aspose.Words.Fields.BarcodeParameters and set only the properties defined there.
  - Prevention: Consult the Aspose.Words.Fields.BarcodeParameters documentation to use supported members only.

- **Transient gateway timeout**
  - Symptom: 504 Gateway Time-out during code generation or execution.
  - Root cause: Temporary network or service latency unrelated to the source code.
  - Preferred fix: Retry the operation or allow the alternate engine result when the example otherwise passes verification.
  - Prevention: Implement retry logic with exponential back-off for external service calls.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
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

- Keep the Aspose.Words, Aspose.BarCode, and Aspose.Drawing.Common package versions synchronized and update using directives when libraries are upgraded.
