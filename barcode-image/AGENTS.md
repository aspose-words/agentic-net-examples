---
name: barcode-image
description: Verified C# examples for barcode image scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - BarCode Image

## Purpose

This folder is a **live, curated example set** for barcode-image scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free generation and rendering of barcodes in Word-centric workflows, with a strong preference for typed field APIs and a custom barcode generator for rendered output.

## Non-negotiable conventions

- Always use `Aspose.Words.Fields.BarcodeParameters` explicitly.
- Always use `Aspose.Drawing.Font` and `Aspose.Drawing.Color` explicitly.
- Never rely on string-based `DISPLAYBARCODE` field syntax in this category.
- Create typed barcode fields via `DocumentBuilder.InsertField(FieldType.FieldDisplayBarcode, true)` and cast to `FieldDisplayBarcode`.
- For PDF, RTF, image, extraction, validation, or rendered-output workflows, always register `doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();`.
- Use the category reference implementation for `CustomBarcodeGeneratorUtils` and `CustomBarcodeGenerator`.

## Recommended workflow selection

- **Custom generator workflow**: 25 examples
- **Word-field-only workflow**: 5 examples

Use the simpler word-field-only pattern only for pure DOCX scenarios with no rendered-output requirement. Otherwise, default to the custom-generator workflow.

## Validation priorities

1. The code must compile and run without manual input.
2. Rendered outputs must not contain the fallback error text `Error! Bar code generator is not set.`
3. Font, color, and barcode parameter types must be fully qualified where ambiguity exists.
4. Examples that depend on templates or folders should bootstrap those inputs locally during the example run.

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
  - Selected engine: mcp
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
  - Selected engine: llm
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
  - Task: Configure the barcode generator to produce high‑resolution images suitable for large‑format PDF printing.
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
  - Selected engine: llm
- `load-an-existing-docx-with-displaybarcode-fields-assign-a-custom-generator-and-export-to-p.cs`
  - Task: Load an existing DOCX with DISPLAYBARCODE fields, assign a custom generator, and export to PDF.
  - Workflow: custom-generator
  - Outputs: docx, pdf
  - Selected engine: mcp
- `load-a-docx-template-populate-displaybarcode-fields-with-user-data-and-export-the-document.cs`
  - Task: Load a DOCX template, populate DISPLAYBARCODE fields with user data, and export the document as PDF.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: llm
- `batch-process-a-folder-of-doc-files-render-barcodes-and-save-each-document-as-pdf.cs`
  - Task: Batch process a folder of DOC files, render barcodes, and save each document as PDF.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: llm
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
  - Selected engine: llm
- `test-barcode-rendering-when-saving-a-document-to-pdf-a-format-ensuring-archival-compliance.cs`
  - Task: Test barcode rendering when saving a document to PDF/A format, ensuring archival compliance.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: mcp
- `add-a-logging-mechanism-to-record-each-barcode-generation-event-with-field-name-and-image.cs`
  - Task: Add a logging mechanism to record each barcode generation event with field name and image size.
  - Workflow: custom-generator
  - Outputs: pdf
  - Selected engine: llm
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

- **Ambiguous Font type**
  - Symptom: Compiler error: 'Font' is an ambiguous reference between Aspose.Drawing.Font and Aspose.Words.Font.
  - Preferred fix: Always use Aspose.Drawing.Font explicitly in DrawErrorImage and any barcode-image drawing code.
- **Ambiguous BarcodeParameters type**
  - Symptom: Compiler error: 'BarcodeParameters' is ambiguous between Aspose.Words.Fields and Aspose.BarCode.Generation.
  - Preferred fix: Always use Aspose.Words.Fields.BarcodeParameters in IBarcodeGenerator signatures and in any verification object construction.
- **Incorrect FieldDisplayBarcode.BarcodeParameters usage**
  - Symptom: Compiler error: FieldDisplayBarcode does not expose a BarcodeParameters property.
  - Preferred fix: Set typed FieldDisplayBarcode properties directly. When a parameter object is needed, create a new Aspose.Words.Fields.BarcodeParameters instance instead of reading a non-existent field property.
- **Wrong Aspose.Drawing FontStyle / constructor usage**
  - Symptom: Compiler/runtime errors around Aspose.Drawing.Drawing2D.FontStyle or invalid Aspose.Drawing.Font constructor arguments.
  - Preferred fix: Use new Aspose.Drawing.Font("Microsoft Sans Serif", 8f, Aspose.Drawing.FontStyle.Regular) and avoid Aspose.Drawing.Drawing2D.FontStyle.
- **Invalid EncodeTypes member**
  - Symptom: Compiler error: EncodeTypes does not contain a definition for Code39Standard.
  - Preferred fix: Use only valid Aspose.BarCode.Generation.EncodeTypes members and keep the Word-to-BarCode mapping method constrained to supported values.
- **Verifier missing input asset/path**
  - Symptom: Runtime failures caused by missing Template.docx or missing input directories during sandbox verification.
  - Preferred fix: Create temporary sample input assets inside the example before loading them, or guard file-system input scenarios with existence checks and bootstrap sample files/folders.
- **Transient MCP gateway timeout**
  - Symptom: 504 Gateway Time-out during generation, not a code-quality issue.
  - Preferred fix: Treat as infrastructure noise. Re-run generation or allow the alternate engine result when the produced example is otherwise verified.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`
- Rendered barcode scenarios also require:
  - `Aspose.BarCode` `26.3.0`
  - `Aspose.Drawing.Common` `25.11.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.BarCode --version 26.3.0
dotnet add package Aspose.Drawing.Common --version 25.11.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\barcode-image\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```
## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the verified winner from the latest batch report rather than a merely compiling draft.
- If two engines both pass, retaining the current selected winner is acceptable unless the alternate output is materially cleaner or more maintainable.
