---
name: barcode-image
description: C# examples for barcode-image using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - barcode-image

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **barcode-image** category.
This folder contains standalone C# examples for barcode-image operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **barcode-image**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (21/21 files) ← category-specific
- `using Aspose.Words.Fields;` (19/21 files)
- `using System;` (18/21 files)
- `using System.IO;` (9/21 files)
- `using Aspose.Words.Saving;` (7/21 files)
- `using System.Collections.Generic;` (4/21 files)
- `using System.Threading.Tasks;` (1/21 files)
- `using System.Diagnostics;` (1/21 files)
- `using Aspose.Words.Drawing;` (1/21 files)
- `using Aspose.Words.Loading;` (1/21 files)

## Common Code Pattern

Most files follow this pattern:

```csharp
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
// ... operations ...
doc.Save("output.docx");
```

## Files in this folder

| File | Key APIs | Description |
|------|----------|-------------|
| [add-logging-mechanism-record-each-barcode-generation-ev...](./add-logging-mechanism-record-each-barcode-generation-event-field-name-image-size.cs) | `Document`, `DocumentBuilder`, `FieldOptions` | Add logging mechanism record each barcode generation event field name image size |
| [apply-different-barcode-types-separate-displaybarcode-f...](./apply-different-barcode-types-separate-displaybarcode-fields-same-document-verify.cs) | `FieldType`, `Document`, `DocumentBuilder` | Apply different barcode types separate displaybarcode fields same document ve... |
| [barcodes-variable-widths-based-input-string-length-prog...](./barcodes-variable-widths-based-input-string-length-programmatically-adjusting-field.cs) | `Document`, `DocumentBuilder`, `FieldType` | Barcodes variable widths based input string length programmatically adjusting... |
| [batch-process-folder-doc-files-render-barcodes-each-doc...](./batch-process-folder-doc-files-render-barcodes-each-document-as-pdf.cs) | `Document`, `AppContext` | Batch process folder doc files render barcodes each document as pdf |
| [configure-barcode-height-width-via-field-switches-displ...](./configure-barcode-height-width-via-field-switches-displaybarcode-field-definition.cs) | `Document`, `DocumentBuilder`, `FieldType` | Configure barcode height width via field switches displaybarcode field defini... |
| [configure-custom-barcode-generator-cache-images-repeate...](./configure-custom-barcode-generator-cache-images-repeated-field-values-improving.cs) | `Document`, `DocumentBuilder`, `FieldType` | Configure custom barcode generator cache images repeated field values improving |
| [console-application-that-accepts-directory-path-process...](./console-application-that-accepts-directory-path-processes-supported-files-generates.cs) | `Document`, `DocumentBuilder`, `PdfSaveOptions` | Console application that accepts directory path processes supported files gen... |
| [customize-barcode-color-background-via-additional-field...](./customize-barcode-color-background-via-additional-field-switches-verify-visual.cs) | `Document`, `DocumentBuilder`, `PdfSaveOptions` | Customize barcode color background via additional field switches verify visual |
| [document-displaybarcode-fields-as-rtf-ensuring-barcodes...](./document-displaybarcode-fields-as-rtf-ensuring-barcodes-render-as-images-output.cs) | `Document`, `DocumentBuilder`, `RtfSaveOptions` | Document displaybarcode fields as rtf ensuring barcodes render as images output |
| [documentbuilder-insert-displaybarcode-field-datamatrix-...](./documentbuilder-insert-displaybarcode-field-datamatrix-barcode-type-switch.cs) | `Document`, `DocumentBuilder`, `FieldType` | Documentbuilder insert displaybarcode field datamatrix barcode type switch |
| [existing-docx-displaybarcode-fields-assign-custom-gener...](./existing-docx-displaybarcode-fields-assign-custom-generator-export-pdf.cs) | `Document`, `DocumentBuilder`, `CustomBarcodeGenerator` | Existing docx displaybarcode fields assign custom generator export pdf |
| [implement-error-handling-missing-barcode-data-displayba...](./implement-error-handling-missing-barcode-data-displaybarcode-fields-avoid-document.cs) | `Document`, `DocumentBuilder`, `Input` | Implement error handling missing barcode data displaybarcode fields avoid doc... |
| [implement-feature-disable-barcode-rendering-specific-fi...](./implement-feature-disable-barcode-rendering-specific-fields-during-pdf-export-while.cs) | `Document`, `DocumentBuilder`, `DOCX` | Implement feature disable barcode rendering specific fields during pdf export... |
| [macro-insert-displaybarcode-fields-predefined-switches-...](./macro-insert-displaybarcode-fields-predefined-switches-various-barcode-types.cs) | `Document`, `DocumentBuilder`, `DisplayBarcodes` | Macro insert displaybarcode fields predefined switches various barcode types |
| [new-document-insert-displaybarcode-field-then-document-...](./new-document-insert-displaybarcode-field-then-document-as-docx.cs) | `Document`, `DocumentBuilder`, `FieldType` | New document insert displaybarcode field then document as docx |
| [process-multiple-docx-files-parallel-each-its-own-barco...](./process-multiple-docx-files-parallel-each-its-own-barcode-generator-output-pdfs.cs) | `Document`, `DocumentBuilder`, `Doc1` | Process multiple docx files parallel each its own barcode generator output pdfs |
| [replace-placeholder-text-displaybarcode-field-dynamic-v...](./replace-placeholder-text-displaybarcode-field-dynamic-values-prior-barcode-generation.cs) | `Document`, `DocumentBuilder`, `Collections` | Replace placeholder text displaybarcode field dynamic values prior barcode ge... |
| [reusable-method-insert-displaybarcode-field-customizabl...](./reusable-method-insert-displaybarcode-field-customizable-height-width-type-switches.cs) | `Document`, `DocumentBuilder`, `BarcodeHelper` | Reusable method insert displaybarcode field customizable height width type sw... |
| [set-barcode-orientation-vertical-via-field-switches-ver...](./set-barcode-orientation-vertical-via-field-switches-verify-correct-rendering-pdf-output.cs) | `Document`, `DocumentBuilder`, `FieldType` | Set barcode orientation vertical via field switches verify correct rendering... |
| [test-barcode-rendering-when-document-pdf-format-ensurin...](./test-barcode-rendering-when-document-pdf-format-ensuring-archival-compliance.cs) | `Document`, `DocumentBuilder`, `PdfSaveOptions` | Test barcode rendering when document pdf format ensuring archival compliance |
| [validate-that-barcode-images-maintain-correct-aspect-ra...](./validate-that-barcode-images-maintain-correct-aspect-ratio-after-converting-document.cs) | `Document`, `Barcodes`, `DocumentBuilder` | Validate that barcode images maintain correct aspect ratio after converting d... |

## Category Statistics
- Total examples: 21

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for barcode-image patterns.


## Command Reference

### Build and Run

Files in this folder are standalone `.cs` examples. Run one example at a time by copying it into a temporary console project as `Program.cs`.

```bash
# Create a temporary console project from the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\barcode-image\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

### Notes for Agents and Developers

- Treat every `.cs` file in `barcode-image/` as a full console program, not a snippet.
- Run one file at a time by copying it to `Program.cs`.
- If a sample needs input documents, images, fonts, or data files, place them in the temporary project directory before running.
- See the root [AGENTS.md](../AGENTS.md) for repository-wide prerequisites, project file template, and testing guidance.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
