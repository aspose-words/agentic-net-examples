# BarCode Image Examples for Aspose.Words for .NET

The BarCode Image category demonstrates how to generate and render barcodes in Aspose.Words documents using the typed FieldDisplayBarcode API and a custom IBarcodeGenerator implementation. Verified examples cover inserting barcode fields, configuring visual properties, handling orientation, caching images, and exporting to PDF with high-resolution output.

## Snapshot

- Category: BarCode Image
- Slug: barcode-image
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261102_031946_08fb5b
- Word Field Only examples: 5
- Custom Generator examples: 25

## Category rules that shaped these examples

- Never reference Font, Color, or BarcodeParameters without their full Aspose.Drawing or Aspose.Words.Fields qualification.
- Do not use builder.InsertField with a raw DISPLAYBARCODE string or any string-based field switches.
- Never rely on the DISPLAYBARCODE field rendering path when the target format is PDF or an image.
- Create a barcode field via DocumentBuilder.InsertField(FieldType.FieldDisplayBarcode, true) and cast to FieldDisplayBarcode.
- Assign barcode properties (BarcodeType, BarcodeValue, ForegroundColor, BackgroundColor, ScalingFactor, SymbolRotation) directly on the FieldDisplayBarcode instance.
- Register a custom IBarcodeGenerator on doc.FieldOptions.BarcodeGenerator for all non-DOCX rendered outputs.
- Always check that BarcodeValue and BarcodeType are non-empty before updating fields.
- Guard file-system paths with existence checks before loading templates or assets.
- Validate that enum values used with EncodeTypes or SymbologyEncodeType exist in the referenced Aspose.BarCode version.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.10.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
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
dotnet add package Aspose.Words --version 26.10.0

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
| 7 | `set-barcode-orientation-to-vertical-via-field-switches-and-verify-correct-rendering-in-pdf.cs` | Set barcode orientation to vertical via field switches and verify correct rendering in PDF output. | custom-generator | pdf | existing_repo |
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
| 24 | `validate-barcode-images-are-correctly-embedded-in-pdf-by-extracting-them-and-comparing-dim.cs` | Validate barcode images are correctly embedded in PDF by extracting them and comparing dimensions. | custom-generator | pdf | mcp |
| 25 | `validate-that-barcode-images-maintain-correct-aspect-ratio-after-converting-a-document-fro.cs` | Validate that barcode images maintain correct aspect ratio after converting a document from DOC to PDF. | custom-generator | pdf | mcp |
| 26 | `test-barcode-rendering-when-saving-a-document-to-pdf-a-format-ensuring-archival-compliance.cs` | Test barcode rendering when saving a document to PDF/A format, ensuring archival compliance. | custom-generator | pdf | mcp |
| 27 | `add-a-logging-mechanism-to-record-each-barcode-generation-event-with-field-name-and-image.cs` | Add a logging mechanism to record each barcode generation event with field name and image size. | custom-generator | pdf | mcp |
| 28 | `create-a-unit-test-that-loads-a-doc-file-renders-barcodes-and-asserts-pdf-output-contains.cs` | Create a unit test that loads a DOC file, renders barcodes, and asserts PDF output contains expected images. | custom-generator | pdf | mcp |
| 29 | `generate-barcodes-with-variable-widths-based-on-input-string-length-by-programmatically-ad.cs` | Generate barcodes with variable widths based on input string length by programmatically adjusting field switch parameters. | custom-generator | pdf | mcp |
| 30 | `implement-a-feature-to-disable-barcode-rendering-for-specific-fields-during-pdf-export-whi.cs` | Implement a feature to disable barcode rendering for specific fields during PDF export while retaining them in DOCX. | custom-generator | docx, pdf | mcp |

## Common failure patterns seen during generation and how they were corrected

### Ambiguous type references

- Symptom: Compiler reports an ambiguous reference for Font or BarcodeParameters.
- Root cause: Both Aspose.Drawing and Aspose.Words (or Aspose.BarCode) define types with the same name, and the code uses an unqualified identifier.
- Preferred fix: Qualify the type with its full namespace (Aspose.Drawing.Font, Aspose.Words.Fields.BarcodeParameters) everywhere barcode code is written.
- Prevention: Avoid adding broad using directives for conflicting namespaces; always use fully qualified names for Font, Color, and BarcodeParameters.

### Use of non-existent FieldDisplayBarcode members

- Symptom: CS1061 errors such as "FieldDisplayBarcode does not contain a definition for 'Height'" or "'Switches'".
- Root cause: Legacy examples rely on field-code string switches that are not exposed as properties on the typed FieldDisplayBarcode class.
- Preferred fix: Set only the supported typed properties (BarcodeType, BarcodeValue, ForegroundColor, BackgroundColor, ScalingFactor, SymbolRotation) and let the custom generator handle size and orientation.
- Prevention: Consult the Aspose.Words API reference for FieldDisplayBarcode and remove any code that accesses undefined members.

### Incorrect namespace usage

- Symptom: CS0234 errors indicating that namespaces such as Aspose.Words.Building, Aspose.Words.Nodes, or Aspose.Words.Builder do not exist.
- Root cause: Typographical errors or outdated documentation introduced invalid namespace names.
- Preferred fix: Replace the incorrect namespaces with the correct ones (e.g., Aspose.Words.Loading, Aspose.Words.Saving, Aspose.Words.Replacing).
- Prevention: Reference the official Aspose.Words documentation for the exact namespace hierarchy before adding using directives.

### Invalid EncodeTypes enum members

- Symptom: CS0117 or CS0722 errors stating that EncodeTypes does not contain definitions such as Code39Standard or that EncodeTypes is used as a static type.
- Root cause: The code uses obsolete or non-existent enum members from an older version of Aspose.BarCode.
- Preferred fix: Use the current enum members (e.g., EncodeTypes.Code39, EncodeTypes.QR, SymbologyEncodeType) that exist in version 26.10.0.
- Prevention: Verify enum members against the installed Aspose.BarCode package version and avoid hard-coding values that are not present.

### Missing input assets at runtime

- Symptom: FileNotFoundException or DirectoryNotFoundException when loading template.docx, input.docx, or processing a folder of documents.
- Root cause: Sample code assumes external files exist in the execution directory without creating or checking them.
- Preferred fix: Create temporary sample files programmatically before loading, or add existence checks and fallback sample content.
- Prevention: Always bootstrap required assets in examples or guard file-system operations with File.Exists / Directory.Exists checks.

### Transient MCP gateway timeout

- Symptom: 504 Gateway Time-out errors during generation.
- Root cause: Intermittent network or service availability issues unrelated to the source code.
- Preferred fix: Retry the generation or allow the alternative engine to produce the output when the primary engine fails.
- Prevention: Implement retry logic with exponential back-off for network-bound operations; this does not require code changes to the barcode logic.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Update the fully qualified type references if future Aspose releases introduce new conflicting namespaces.
- Periodically verify that the custom IBarcodeGenerator implementation still matches the latest Aspose.BarCode API signatures.
