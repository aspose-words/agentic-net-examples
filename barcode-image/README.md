# BarCode Image Examples for Aspose.Words for .NET

The BarCode Image category demonstrates how to embed and render barcodes in Word documents using Aspose.Words and Aspose.BarCode. Verified examples show typed FieldDisplayBarcode usage, custom IBarcodeGenerator implementations, and PDF export with correctly rendered barcode images.

## Snapshot

- Category: BarCode Image
- Slug: barcode-image
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261124_190740_1c00f5
- Word Field Only examples: 5
- Custom Generator examples: 25

## Category rules that shaped these examples

- Never use unqualified BarcodeParameters; always reference Aspose.Words.Fields.BarcodeParameters.
- Never use unqualified Font; always reference Aspose.Drawing.Font.
- Never use unqualified Color; always reference Aspose.Drawing.Color.
- Never import System.Drawing types in this category.
- Do not insert DISPLAYBARCODE fields using string syntax or builder.InsertField with a raw field code.
- Create barcode fields with the typed FieldDisplayBarcode object, assign properties directly, register a custom IBarcodeGenerator, update fields, and save the document.
- Validate that BarcodeValue, BarcodeType, and color strings are non-null before assigning them to a FieldDisplayBarcode.

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
dotnet add package Aspose.BarCode --version 26.10.0
dotnet add package Aspose.Drawing.Common --version 26.10.0

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
dotnet add package Aspose.Words --version 26.11.0

# PowerShell example
Copy-Item ..\barcode-image\create-a-new-document-insert-a-displaybarcode-field-then-save-the-document-as-docx.cs .\Program.cs

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
| 5 | `configure-barcode-height-and-width-via-field-switches-in-the-displaybarcode-field-definiti.cs` | Configure barcode height and width via field switches in the DISPLAYBARCODE field definition. | custom-generator | pdf | llm |
| 6 | `customize-barcode-color-and-background-via-additional-field-switches-and-verify-visual-app.cs` | Customize barcode color and background via additional field switches and verify visual appearance in saved PDF. | custom-generator | pdf | mcp |
| 7 | `set-barcode-orientation-to-vertical-via-field-switches-and-verify-correct-rendering-in-pdf.cs` | Set barcode orientation to vertical via field switches and verify correct rendering in PDF output. | custom-generator | pdf | llm |
| 8 | `apply-different-barcode-types-to-separate-displaybarcode-fields-in-the-same-document-and-v.cs` | Apply different barcode types to separate DISPLAYBARCODE fields in the same document and verify correct rendering. | custom-generator | pdf | mcp |
| 9 | `replace-placeholder-text-in-a-displaybarcode-field-with-dynamic-values-prior-to-barcode-ge.cs` | Replace placeholder text in a DISPLAYBARCODE field with dynamic values prior to barcode generation. | word-field-only | docx | mcp |
| 10 | `use-document-range-replace-to-update-the-data-string-of-an-existing-displaybarcode-field-b.cs` | Use Document.Range.Replace to update the data string of an existing DISPLAYBARCODE field before rendering. | word-field-only | docx | mcp |
| 11 | `implement-error-handling-for-missing-barcode-data-in-displaybarcode-fields-to-avoid-docume.cs` | Implement error handling for missing barcode data in DISPLAYBARCODE fields to avoid document save failures. | custom-generator | docx | mcp |
| 12 | `implement-the-ibarcodegenerator-interface-to-generate-code128-barcodes-from-field-data.cs` | Implement the IBarcodeGenerator interface to generate Code128 barcodes from field data. | custom-generator | pdf | mcp |
| 13 | `integrate-the-aspose-barcode-library-to-enable-qr-code-generation-for-displaybarcode-field.cs` | Integrate the Aspose.BarCode library to enable QR code generation for DISPLAYBARCODE fields. | custom-generator | docx | mcp |
| 14 | `configure-the-custom-barcode-generator-to-cache-images-for-repeated-field-values-improving.cs` | Configure the custom barcode generator to cache images for repeated field values, improving performance. | custom-generator | pdf | mcp |
| 15 | `configure-the-barcode-generator-to-produce-high-resolution-images-suitable-for-large-forma.cs` | Configure the barcode generator to produce high-resolution images suitable for large-format PDF printing. | custom-generator | pdf | mcp |
| 16 | `generate-barcode-images-on-the-fly-during-save-by-assigning-the-custom-generator-to-docume.cs` | Generate barcode images on the fly during save by assigning the custom generator to Document.BarcodeGenerator. | custom-generator | pdf | mcp |
| 17 | `load-an-rtf-document-assign-a-custom-generator-and-save-the-result-as-docx-with-barcodes.cs` | Load an RTF document, assign a custom generator, and save the result as DOCX with barcodes. | custom-generator | docx | mcp |
| 18 | `load-an-existing-docx-with-displaybarcode-fields-assign-a-custom-generator-and-export-to-p.cs` | Load an existing DOCX with DISPLAYBARCODE fields, assign a custom generator, and export to PDF. | custom-generator | docx, pdf | mcp |
| 19 | `load-a-docx-template-populate-displaybarcode-fields-with-user-data-and-export-the-document.cs` | Load a DOCX template, populate DISPLAYBARCODE fields with user data, and export the document as PDF. | custom-generator | pdf | mcp |
| 20 | `batch-process-a-folder-of-doc-files-render-barcodes-and-save-each-document-as-pdf.cs` | Batch process a folder of DOC files, render barcodes, and save each document as PDF. | custom-generator | pdf | mcp |
| 21 | `process-multiple-docx-files-in-parallel-each-with-its-own-barcode-generator-and-output-pdf.cs` | Process multiple DOCX files in parallel, each with its own barcode generator, and output PDFs concurrently. | custom-generator | docx, pdf | mcp |
| 22 | `create-a-console-application-that-accepts-a-directory-path-processes-supported-files-and-g.cs` | Create a console application that accepts a directory path, processes supported files, and generates PDFs with barcodes. | custom-generator | docx, pdf | mcp |
| 23 | `save-a-document-with-displaybarcode-fields-as-rtf-ensuring-barcodes-render-as-images-in-ou.cs` | Save a document with DISPLAYBARCODE fields as RTF, ensuring barcodes render as images in output. | custom-generator | rtf | mcp |
| 24 | `validate-barcode-images-are-correctly-embedded-in-pdf-by-extracting-them-and-comparing-dim.cs` | Validate barcode images are correctly embedded in PDF by extracting them and comparing dimensions. | custom-generator | pdf | llm |
| 25 | `validate-that-barcode-images-maintain-correct-aspect-ratio-after-converting-a-document-fro.cs` | Validate that barcode images maintain correct aspect ratio after converting a document from DOC to PDF. | custom-generator | pdf | mcp |
| 26 | `test-barcode-rendering-when-saving-a-document-to-pdf-a-format-ensuring-archival-compliance.cs` | Test barcode rendering when saving a document to PDF/A format, ensuring archival compliance. | custom-generator | pdf | mcp |
| 27 | `add-a-logging-mechanism-to-record-each-barcode-generation-event-with-field-name-and-image.cs` | Add a logging mechanism to record each barcode generation event with field name and image size. | custom-generator | pdf | mcp |
| 28 | `create-a-unit-test-that-loads-a-doc-file-renders-barcodes-and-asserts-pdf-output-contains.cs` | Create a unit test that loads a DOC file, renders barcodes, and asserts PDF output contains expected images. | custom-generator | pdf | llm |
| 29 | `generate-barcodes-with-variable-widths-based-on-input-string-length-by-programmatically-ad.cs` | Generate barcodes with variable widths based on input string length by programmatically adjusting field switch parameters. | custom-generator | pdf | mcp |
| 30 | `implement-a-feature-to-disable-barcode-rendering-for-specific-fields-during-pdf-export-whi.cs` | Implement a feature to disable barcode rendering for specific fields during PDF export while retaining them in DOCX. | custom-generator | docx, pdf | mcp |

## Common failure patterns seen during generation and how they were corrected

### Ambiguous type reference

- Symptom: Compiler reports an ambiguous reference for 'Font' or 'BarcodeParameters' between Aspose.Drawing and Aspose.Words namespaces.
- Root cause: Unqualified type names cause the C# compiler to see multiple matching types from different assemblies.
- Preferred fix: Qualify the type with its full namespace (e.g., Aspose.Drawing.Font, Aspose.Words.Fields.BarcodeParameters).
- Prevention: Always use fully qualified names or explicit using aliases for these types.

### Invalid FieldDisplayBarcode member usage

- Symptom: CS1061 errors for Height, Width, SetSwitch, SetFieldCode, or Orientation members that do not exist on FieldDisplayBarcode.
- Root cause: Attempting to control barcode appearance through legacy field switches instead of the typed API.
- Preferred fix: Use supported typed properties such as ScalingFactor, ForegroundColor, BackgroundColor, and rely on a custom generator for size adjustments.
- Prevention: Do not call non-existent members; consult the API reference for available FieldDisplayBarcode properties.

### EncodeTypes enum misuse

- Symptom: CS0117 or CS0723 errors when assigning SymbologyEncodeType values to EncodeTypes or using undefined EncodeTypes members.
- Root cause: Confusing the static EncodeTypes class with the SymbologyEncodeType enum used by Aspose.BarCode.
- Preferred fix: Map barcode type strings to SymbologyEncodeType values (e.g., EncodeTypes.Code128) and avoid treating EncodeTypes as an enum.
- Prevention: Always use SymbologyEncodeType for generator construction and keep the mapping logic isolated.

### Missing Aspose.Words namespace

- Symptom: CS0234 errors for namespaces such as Aspose.Words.Building, Aspose.Words.Nodes, or Aspose.Words.LoadOptions.
- Root cause: Incorrect or outdated using directives referencing non-existent namespaces in the installed package version.
- Preferred fix: Remove invalid namespaces and import only existing ones like Aspose.Words.Drawing and Aspose.Words.Replacing.
- Prevention: Verify namespace availability against the current Aspose.Words version before adding using statements.

### Missing input file or directory

- Symptom: Runtime FileNotFoundException or DirectoryNotFoundException when loading template documents or processing folders.
- Root cause: Example code assumes external assets without creating or checking them.
- Preferred fix: Create temporary sample files/folders within the example or guard file-system operations with existence checks.
- Prevention: Ensure all required assets are provisioned programmatically before they are accessed.

### Undefined ImageFormat

- Symptom: CS0103 error indicating that 'ImageFormat' does not exist in the current context.
- Root cause: Missing reference to Aspose.Drawing.Imaging or accidental reliance on System.Drawing types.
- Preferred fix: Add 'using Aspose.Drawing.Imaging;' and reference Aspose.Drawing.Imaging.ImageFormat.
- Prevention: Always import the Aspose.Drawing.Imaging namespace when working with image formats.

### BarcodeParameters property mismatch

- Symptom: CS1061 errors for missing members such as BarcodeSymbology, BarcodeForeColor, or FieldName on BarcodeParameters.
- Root cause: Using the wrong BarcodeParameters type (Aspose.BarCode.Generation) or assuming properties that are not part of Aspose.Words.Fields.BarcodeParameters.
- Preferred fix: Reference Aspose.Words.Fields.BarcodeParameters and set only the properties defined there.
- Prevention: Consult the Aspose.Words.Fields.BarcodeParameters documentation to use supported members only.

### Transient gateway timeout

- Symptom: 504 Gateway Time-out during code generation or execution.
- Root cause: Temporary network or service latency unrelated to the source code.
- Preferred fix: Retry the operation or allow the alternate engine result when the example otherwise passes verification.
- Prevention: Implement retry logic with exponential back-off for external service calls.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Keep the Aspose.Words, Aspose.BarCode, and Aspose.Drawing.Common package versions synchronized and update using directives when libraries are upgraded.
