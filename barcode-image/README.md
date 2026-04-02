# BarCode Image Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **BarCode Image** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **BarCode Image**
- Slug: **barcode-image**
- Total examples: **30**
- Verified winners: **both=19**, **mcp-only=4**, **llm-only=7**, **none=0**
- Custom generator workflow examples: **25 / 30**
- Word-field-only DOCX examples: **5 / 30**

## Category rules that shaped these examples

- Always qualify ambiguous APIs explicitly, especially `Aspose.Words.Fields.BarcodeParameters`, `Aspose.Drawing.Font`, and `Aspose.Drawing.Color`.
- Do not use string-based `DISPLAYBARCODE` field construction. Prefer typed `FieldDisplayBarcode` creation and explicit property assignment.
- Use `CustomBarcodeGenerator` whenever rendered output is involved, especially PDF/RTF/image workflows or any validation/extraction scenario.
- Use Aspose.BarCode through a supported Word-to-BarCode mapping method that returns `SymbologyEncodeType` and uses valid `EncodeTypes` values only.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`
- For rendered barcode scenarios, Aspose.BarCode for .NET `26.3.0`
- For drawing/image helpers used by the custom generator, Aspose.Drawing.Common `25.11.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.BarCode --version 26.3.0
dotnet add package Aspose.Drawing.Common --version 25.11.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\barcode-image\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `barcode-image/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.BarCode --version 26.3.0
dotnet add package Aspose.Drawing.Common --version 25.11.0

# PowerShell example
Copy-Item ..\barcode-image\implement-the-ibarcodegenerator-interface-to-generate-code128-barcodes-from-field-data.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `create-a-new-document-insert-a-displaybarcode-field-then-save-the-document-as-docx.cs` | Create a new Document, insert a DISPLAYBARCODE field, then save the document as DOCX. | word-field-only | docx | mcp |
| 2 | `use-documentbuilder-to-insert-a-displaybarcode-field-with-a-datamatrix-barcode-type-switch.cs` | Use DocumentBuilder to insert a DISPLAYBARCODE field with a DataMatrix barcode type switch. | custom-generator | pdf | mcp |
| 3 | `create-a-reusable-method-to-insert-a-displaybarcode-field-with-customizable-height-width-a.cs` | Create a reusable method to insert a DISPLAYBARCODE field with customizable height, width, and type switches. | word-field-only | docx | mcp |
| 4 | `create-a-macro-to-insert-displaybarcode-fields-with-predefined-switches-for-various-barcod.cs` | Create a macro to insert DISPLAYBARCODE fields with predefined switches for various barcode types. | word-field-only | docx | mcp |
| 5 | `configure-barcode-height-and-width-via-field-switches-in-the-displaybarcode-field-definiti.cs` | Configure barcode height and width via field switches in the DISPLAYBARCODE field definition. | custom-generator | pdf | mcp |
| 6 | `customize-barcode-color-and-background-via-additional-field-switches-and-verify-visual-app.cs` | Customize barcode color and background via additional field switches and verify visual appearance in saved PDF. | custom-generator | pdf | mcp |
| 7 | `set-barcode-orientation-to-vertical-via-field-switches-and-verify-correct-rendering-in-pdf.cs` | Set barcode orientation to vertical via field switches and verify correct rendering in PDF output. | custom-generator | pdf | llm |
| 8 | `apply-different-barcode-types-to-separate-displaybarcode-fields-in-the-same-document-and-v.cs` | Apply different barcode types to separate DISPLAYBARCODE fields in the same document and verify correct rendering. | custom-generator | pdf | llm |
| 9 | `replace-placeholder-text-in-a-displaybarcode-field-with-dynamic-values-prior-to-barcode-ge.cs` | Replace placeholder text in a DISPLAYBARCODE field with dynamic values prior to barcode generation. | word-field-only | docx | mcp |
| 10 | `use-document-range-replace-to-update-the-data-string-of-an-existing-displaybarcode-field-b.cs` | Use Document.Range.Replace to update the data string of an existing DISPLAYBARCODE field before rendering. | word-field-only | docx | mcp |
| 11 | `implement-error-handling-for-missing-barcode-data-in-displaybarcode-fields-to-avoid-docume.cs` | Implement error handling for missing barcode data in DISPLAYBARCODE fields to avoid document save failures. | custom-generator | docx | mcp |
| 12 | `implement-the-ibarcodegenerator-interface-to-generate-code128-barcodes-from-field-data.cs` | Implement the IBarcodeGenerator interface to generate Code128 barcodes from field data. | custom-generator | pdf | mcp |
| 13 | `integrate-the-aspose-barcode-library-to-enable-qr-code-generation-for-displaybarcode-field.cs` | Integrate the Aspose.BarCode library to enable QR code generation for DISPLAYBARCODE fields. | custom-generator | docx | mcp |
| 14 | `configure-the-custom-barcode-generator-to-cache-images-for-repeated-field-values-improving.cs` | Configure the custom barcode generator to cache images for repeated field values, improving performance. | custom-generator | pdf | mcp |
| 15 | `configure-the-barcode-generator-to-produce-high-resolution-images-suitable-for-large-forma.cs` | Configure the barcode generator to produce high‑resolution images suitable for large‑format PDF printing. | custom-generator | pdf | mcp |
| 16 | `generate-barcode-images-on-the-fly-during-save-by-assigning-the-custom-generator-to-docume.cs` | Generate barcode images on the fly during save by assigning the custom generator to Document.BarcodeGenerator. | custom-generator | pdf | mcp |
| 17 | `load-an-rtf-document-assign-a-custom-generator-and-save-the-result-as-docx-with-barcodes.cs` | Load an RTF document, assign a custom generator, and save the result as DOCX with barcodes. | custom-generator | docx | llm |
| 18 | `load-an-existing-docx-with-displaybarcode-fields-assign-a-custom-generator-and-export-to-p.cs` | Load an existing DOCX with DISPLAYBARCODE fields, assign a custom generator, and export to PDF. | custom-generator | docx, pdf | mcp |
| 19 | `load-a-docx-template-populate-displaybarcode-fields-with-user-data-and-export-the-document.cs` | Load a DOCX template, populate DISPLAYBARCODE fields with user data, and export the document as PDF. | custom-generator | pdf | llm |
| 20 | `batch-process-a-folder-of-doc-files-render-barcodes-and-save-each-document-as-pdf.cs` | Batch process a folder of DOC files, render barcodes, and save each document as PDF. | custom-generator | pdf | llm |
| 21 | `process-multiple-docx-files-in-parallel-each-with-its-own-barcode-generator-and-output-pdf.cs` | Process multiple DOCX files in parallel, each with its own barcode generator, and output PDFs concurrently. | custom-generator | docx, pdf | mcp |
| 22 | `create-a-console-application-that-accepts-a-directory-path-processes-supported-files-and-g.cs` | Create a console application that accepts a directory path, processes supported files, and generates PDFs with barcodes. | custom-generator | docx, pdf | mcp |
| 23 | `save-a-document-with-displaybarcode-fields-as-rtf-ensuring-barcodes-render-as-images-in-ou.cs` | Save a document with DISPLAYBARCODE fields as RTF, ensuring barcodes render as images in output. | custom-generator | rtf | mcp |
| 24 | `validate-barcode-images-are-correctly-embedded-in-pdf-by-extracting-them-and-comparing-dim.cs` | Validate barcode images are correctly embedded in PDF by extracting them and comparing dimensions. | custom-generator | pdf | mcp |
| 25 | `validate-that-barcode-images-maintain-correct-aspect-ratio-after-converting-a-document-fro.cs` | Validate that barcode images maintain correct aspect ratio after converting a document from DOC to PDF. | custom-generator | pdf | llm |
| 26 | `test-barcode-rendering-when-saving-a-document-to-pdf-a-format-ensuring-archival-compliance.cs` | Test barcode rendering when saving a document to PDF/A format, ensuring archival compliance. | custom-generator | pdf | mcp |
| 27 | `add-a-logging-mechanism-to-record-each-barcode-generation-event-with-field-name-and-image.cs` | Add a logging mechanism to record each barcode generation event with field name and image size. | custom-generator | pdf | llm |
| 28 | `create-a-unit-test-that-loads-a-doc-file-renders-barcodes-and-asserts-pdf-output-contains.cs` | Create a unit test that loads a DOC file, renders barcodes, and asserts PDF output contains expected images. | custom-generator | pdf | mcp |
| 29 | `generate-barcodes-with-variable-widths-based-on-input-string-length-by-programmatically-ad.cs` | Generate barcodes with variable widths based on input string length by programmatically adjusting field switch parameters. | custom-generator | pdf | mcp |
| 30 | `implement-a-feature-to-disable-barcode-rendering-for-specific-fields-during-pdf-export-whi.cs` | Implement a feature to disable barcode rendering for specific fields during PDF export while retaining them in DOCX. | custom-generator | docx, pdf | mcp |

## Common failure patterns seen during generation and how they were corrected

### Ambiguous Font type

- Seen in verification: **3** case(s)
- Symptom: Compiler error: 'Font' is an ambiguous reference between Aspose.Drawing.Font and Aspose.Words.Font.
- Fix: Always use Aspose.Drawing.Font explicitly in DrawErrorImage and any barcode-image drawing code.

### Ambiguous BarcodeParameters type

- Seen in verification: **1** case(s)
- Symptom: Compiler error: 'BarcodeParameters' is ambiguous between Aspose.Words.Fields and Aspose.BarCode.Generation.
- Fix: Always use Aspose.Words.Fields.BarcodeParameters in IBarcodeGenerator signatures and in any verification object construction.

### Incorrect FieldDisplayBarcode.BarcodeParameters usage

- Seen in verification: **1** case(s)
- Symptom: Compiler error: FieldDisplayBarcode does not expose a BarcodeParameters property.
- Fix: Set typed FieldDisplayBarcode properties directly. When a parameter object is needed, create a new Aspose.Words.Fields.BarcodeParameters instance instead of reading a non-existent field property.

### Wrong Aspose.Drawing FontStyle / constructor usage

- Seen in verification: **1** case(s)
- Symptom: Compiler/runtime errors around Aspose.Drawing.Drawing2D.FontStyle or invalid Aspose.Drawing.Font constructor arguments.
- Fix: Use new Aspose.Drawing.Font("Microsoft Sans Serif", 8f, Aspose.Drawing.FontStyle.Regular) and avoid Aspose.Drawing.Drawing2D.FontStyle.

### Invalid EncodeTypes member

- Seen in verification: **1** case(s)
- Symptom: Compiler error: EncodeTypes does not contain a definition for Code39Standard.
- Fix: Use only valid Aspose.BarCode.Generation.EncodeTypes members and keep the Word-to-BarCode mapping method constrained to supported values.

### Verifier missing input asset/path

- Seen in verification: **2** case(s)
- Symptom: Runtime failures caused by missing Template.docx or missing input directories during sandbox verification.
- Fix: Create temporary sample input assets inside the example before loading them, or guard file-system input scenarios with existence checks and bootstrap sample files/folders.

### Transient MCP gateway timeout

- Seen in verification: **2** case(s)
- Symptom: 504 Gateway Time-out during generation, not a code-quality issue.
- Fix: Treat as infrastructure noise. Re-run generation or allow the alternate engine result when the produced example is otherwise verified.

## Notes for maintainers

- The selected file for each task is the verified winner recorded in the batch report.
- When updating this category, preserve the current typed-field and custom-generator conventions. Regressions usually happen when ambiguous drawing APIs are used or when PDF workflows skip `doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();`.
- Template- or directory-based samples should create temporary local assets for verification instead of assuming machine-specific paths.