---
name: ole-objects
description: C# examples for ole-objects using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - ole-objects

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **ole-objects** category.
This folder contains standalone C# examples for ole-objects operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **ole-objects**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using System;` (27/27 files) ← category-specific
- `using Aspose.Words;` (27/27 files)
- `using System.IO;` (24/27 files)
- `using Aspose.Words.Drawing;` (23/27 files)
- `using System.Linq;` (3/27 files)
- `using System.Text;` (2/27 files)
- `using System.Collections.Generic;` (1/27 files)
- `using System.Runtime.InteropServices;` (1/27 files)
- `using Microsoft.Win32;` (1/27 files)
- `using System.IO.Compression;` (1/27 files)
- `using System.Drawing;` (1/27 files)
- `using Aspose.Words.Rendering;` (1/27 files)

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
| [adjust-ole-object-icon-dimensions-maintain-aspect-ratio...](./adjust-ole-object-icon-dimensions-maintain-aspect-ratio-after-resizing-document-layout.cs) | `Document`, `DocumentBuilder`, `Sample` | Adjust ole object icon dimensions maintain aspect ratio after resizing docume... |
| [batch-insert-identical-excel-ole-object-multiple-word-f...](./batch-insert-identical-excel-ole-object-multiple-word-files-foreach-loop.cs) | `Document`, `DocumentBuilder`, `Collections` | Batch insert identical excel ole object multiple word files foreach loop |
| [batch-insert-same-excel-ole-object-multiple-word-docume...](./batch-insert-same-excel-ole-object-multiple-word-documents-within-folder-loop.cs) | `Document`, `DocumentBuilder`, `Sample` | Batch insert same excel ole object multiple word documents within folder loop |
| [batch-process-collection-word-files-embed-common-online...](./batch-process-collection-word-files-embed-common-online-video-each-document.cs) | `Document`, `DocumentBuilder`, `AppContext` | Batch process collection word files embed common online video each document |
| [clone-ole-object-one-document-insert-it-another-inserto...](./clone-ole-object-one-document-insert-it-another-insertoleobject-extracted-data.cs) | `Document`, `DocumentBuilder`, `InvalidOperationException` | Clone ole object one document insert it another insertoleobject extracted data |
| [documentbuilder-insert-ole-object-custom-width-height-p...](./documentbuilder-insert-ole-object-custom-width-height-parameters-precise-layout-control.cs) | `Document`, `DocumentBuilder`, `SampleData` | Documentbuilder insert ole object custom width height parameters precise layo... |
| [embed-legacy-ole-package-when-handler-is-unknown-packag...](./embed-legacy-ole-package-when-handler-is-unknown-packager-approach.cs) | `Document`, `DocumentBuilder`, `OleFormat` | Embed legacy ole package when handler is unknown packager approach |
| [export-all-ole-objects-document-separate-files-preservi...](./export-all-ole-objects-document-separate-files-preserving-original-extensions.cs) | `Document`, `AppContext`, `InputDocument` | Export all ole objects document separate files preserving original extensions |
| [extract-ole-object-metadata-such-as-source-file-name-si...](./extract-ole-object-metadata-such-as-source-file-name-size-write-csv-report.cs) | `Document`, `StringBuilder`, `AppContext` | Extract ole object metadata such as source file name size write csv report |
| [extract-ole-object-stream-document-it-binary-file](./extract-ole-object-stream-document-it-binary-file.cs) | `Document`, `AppContext`, `Sample` | Extract ole object stream document it binary file |
| [implement-error-handling-insertoleobject-when-specified...](./implement-error-handling-insertoleobject-when-specified-progid-is-not-registered-system.cs) | `ArgumentNullException`, `Document`, `DocumentBuilder` | Implement error handling insertoleobject when specified progid is not registe... |
| [insert-ole-object-as-icon-default-system-icon-without-s...](./insert-ole-object-as-icon-default-system-icon-without-specifying-custom-image.cs) | `Document`, `DocumentBuilder`, `Sample` | Insert ole object as icon default system icon without specifying custom image |
| [insert-online-video-specified-width-height-parameters-p...](./insert-online-video-specified-width-height-parameters-precise-placement.cs) | `Document`, `DocumentBuilder`, `WrapType` | Insert online video specified width height parameters precise placement |
| [insert-online-youtube-video-word-document-documentbuild...](./insert-online-youtube-video-word-document-documentbuilder-insertonlinevideo.cs) | `Document`, `DocumentBuilder`, `API` | Insert online youtube video word document documentbuilder insertonlinevideo |
| [insert-spreadsheet-ole-object-docx-documentbuilder-inse...](./insert-spreadsheet-ole-object-docx-documentbuilder-insertoleobject-its-progid.cs) | `Document`, `DocumentBuilder`, `FileMode` | Insert spreadsheet ole object docx documentbuilder insertoleobject its progid |
| [iterate-through-all-ole-objects-docx-file-log-their-pro...](./iterate-through-all-ole-objects-docx-file-log-their-progids-display-sizes.cs) | `Document`, `Input`, `NodeType` | Iterate through all ole objects docx file log their progids display sizes |
| [olepackage-class-enumerate-all-parts-legacy-ole-package...](./olepackage-class-enumerate-all-parts-legacy-ole-package-inspection.cs) | `Document`, `SampleWithOlePackage`, `NodeType` | Olepackage class enumerate all parts legacy ole package inspection |
| [raw-binary-data-ole-object-temporary-file-external-anal...](./raw-binary-data-ole-object-temporary-file-external-analysis.cs) | `Document`, `InputDocument`, `NodeType` | Raw binary data ole object temporary file external analysis |
| [read-file-name-property-ole-package-compare-it-original...](./read-file-name-property-ole-package-compare-it-original-source-file.cs) | `Document`, `Package`, `DocumentWithOlePackage` | Read file name property ole package compare it original source file |
| [replace-existing-ole-object-doc-file-new-image-ole-obje...](./replace-existing-ole-object-doc-file-new-image-ole-object-insertoleobject.cs) | `DocumentBuilder`, `Document`, `Original` | Replace existing ole object doc file new image ole object insertoleobject |
| [retrieve-display-width-height-ole-object-adjust-its-siz...](./retrieve-display-width-height-ole-object-adjust-its-size-after-insertion.cs) | `Document`, `DocumentBuilder`, `OleFormat` | Retrieve display width height ole object adjust its size after insertion |
| [retrieve-ole-object-display-width-height-after-insertio...](./retrieve-ole-object-display-width-height-after-insertion-store-dimensions-layout.cs) | `ShapeRenderer`, `Document`, `DocumentBuilder` | Retrieve ole object display width height after insertion store dimensions layout |
| [retrieve-progid-inserted-ole-object-log-it-diagnostic-p...](./retrieve-progid-inserted-ole-object-log-it-diagnostic-purposes.cs) | `Document`, `DocumentBuilder`, `OleFormat` | Retrieve progid inserted ole object log it diagnostic purposes |
| [retrieve-raw-binary-data-ole-object-via-its-data-proper...](./retrieve-raw-binary-data-ole-object-via-its-data-property-custom-processing.cs) | `Document`, `OleObjects`, `NodeType` | Retrieve raw binary data ole object via its data property custom processing |
| [set-file-name-extension-inserted-ole-object-preserve-or...](./set-file-name-extension-inserted-ole-object-preserve-original-metadata.cs) | `OlePackage`, `Document`, `DocumentBuilder` | Set file name extension inserted ole object preserve original metadata |
| [validate-progid-ole-object-before-insertion-avoid-runti...](./validate-progid-ole-object-before-insertion-avoid-runtime-errors.cs) | `Document`, `DocumentBuilder`, `ArgumentException` | Validate progid ole object before insertion avoid runtime errors |
| [verify-successful-ole-object-insertion-checking-returne...](./verify-successful-ole-object-insertion-checking-returned-object-reference-is-not-null.cs) | `OleFormat`, `Document`, `DocumentBuilder` | Verify successful ole object insertion checking returned object reference is... |

## Category Statistics
- Total examples: 27

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for ole-objects patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
