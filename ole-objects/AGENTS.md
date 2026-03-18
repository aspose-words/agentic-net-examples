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

- `using Aspose.Words;` (29/29 files) ← category-specific
- `using System;` (28/29 files)
- `using Aspose.Words.Drawing;` (25/29 files)
- `using System.IO;` (23/29 files)
- `using System.Collections.Generic;` (2/29 files)
- `using System.Linq;` (2/29 files)
- `using System.Text;` (1/29 files)
- `using System.Runtime.InteropServices;` (1/29 files)
- `using Microsoft.Win32;` (1/29 files)
- `using Aspose.Words.Drawing.Ole;` (1/29 files)
- `using System.Drawing;` (1/29 files)
- `using Aspose.Words.Rendering;` (1/29 files)

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
| [adjust-ole-object-icon-dimensions-maintain-aspect-ratio...](./adjust-ole-object-icon-dimensions-maintain-aspect-ratio-after-resizing-document-layout.cs) | `Document`, `DocumentBuilder`, `Aspose` | Adjust ole object icon dimensions maintain aspect ratio after resizing docume... |
| [batch-insert-identical-excel-ole-object-multiple-word-f...](./batch-insert-identical-excel-ole-object-multiple-word-files-foreach-loop.cs) | `Document`, `DocumentBuilder`, `System` | Batch insert identical excel ole object multiple word files foreach loop |
| [batch-insert-same-excel-ole-object-multiple-word-docume...](./batch-insert-same-excel-ole-object-multiple-word-documents-within-folder-loop.cs) | `Document`, `DocumentBuilder`, `System` | Batch insert same excel ole object multiple word documents within folder loop |
| [batch-process-collection-word-files-embed-common-online...](./batch-process-collection-word-files-embed-common-online-video-each-document.cs) | `Document`, `DocumentBuilder`, `System` | Batch process collection word files embed common online video each document |
| [clone-ole-object-one-document-insert-it-another-inserto...](./clone-ole-object-one-document-insert-it-another-insertoleobject-extracted-data.cs) | `Document`, `Aspose`, `DocumentBuilder` | Clone ole object one document insert it another insertoleobject extracted data |
| [documentbuilder-insert-ole-object-custom-width-height-p...](./documentbuilder-insert-ole-object-custom-width-height-parameters-precise-layout-control.cs) | `Document`, `DocumentBuilder`, `Aspose` | Documentbuilder insert ole object custom width height parameters precise layo... |
| [embed-legacy-ole-package-when-handler-is-unknown-packag...](./embed-legacy-ole-package-when-handler-is-unknown-packager-approach.cs) | `Aspose`, `Document`, `DocumentBuilder` | Embed legacy ole package when handler is unknown packager approach |
| [export-all-ole-objects-document-separate-files-preservi...](./export-all-ole-objects-document-separate-files-preserving-original-extensions.cs) | `Document`, `Aspose`, `System` | Export all ole objects document separate files preserving original extensions |
| [extract-ole-object-metadata-such-as-source-file-name-si...](./extract-ole-object-metadata-such-as-source-file-name-size-write-csv-report.cs) | `Document`, `StringBuilder`, `System` | Extract ole object metadata such as source file name size write csv report |
| [extract-ole-object-stream-document-it-binary-file](./extract-ole-object-stream-document-it-binary-file.cs) | `Document`, `Aspose`, `System` | Extract ole object stream document it binary file |
| [implement-error-handling-insertoleobject-when-specified...](./implement-error-handling-insertoleobject-when-specified-progid-is-not-registered-system.cs) | `Aspose`, `ArgumentNullException`, `Document` | Implement error handling insertoleobject when specified progid is not registe... |
| [insert-ole-object-as-icon-default-system-icon-without-s...](./insert-ole-object-as-icon-default-system-icon-without-specifying-custom-image.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert ole object as icon default system icon without specifying custom image |
| [insert-online-video-specified-width-height-parameters-p...](./insert-online-video-specified-width-height-parameters-precise-placement.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert online video specified width height parameters precise placement |
| [insert-online-youtube-video-word-document-documentbuild...](./insert-online-youtube-video-word-document-documentbuilder-insertonlinevideo.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert online youtube video word document documentbuilder insertonlinevideo |
| [insert-pdf-file-as-ole-icon-specifying-custom-icon-imag...](./insert-pdf-file-as-ole-icon-specifying-custom-icon-image-display-size.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert pdf file as ole icon specifying custom icon image display size |
| [insert-spreadsheet-ole-object-docx-documentbuilder-inse...](./insert-spreadsheet-ole-object-docx-documentbuilder-insertoleobject-its-progid.cs) | `Aspose`, `Document`, `DocumentBuilder` | Insert spreadsheet ole object docx documentbuilder insertoleobject its progid |
| [iterate-through-all-ole-objects-docx-file-log-their-pro...](./iterate-through-all-ole-objects-docx-file-log-their-progids-display-sizes.cs) | `Document`, `Aspose`, `System` | Iterate through all ole objects docx file log their progids display sizes |
| [olepackage-class-enumerate-all-parts-legacy-ole-package...](./olepackage-class-enumerate-all-parts-legacy-ole-package-inspection.cs) | `Document`, `Aspose`, `System` | Olepackage class enumerate all parts legacy ole package inspection |
| [olepackage-class-read-modify-properties-legacy-ole-pack...](./olepackage-class-read-modify-properties-legacy-ole-package-net.cs) | `Document`, `Package`, `Aspose` | Olepackage class read modify properties legacy ole package net |
| [raw-binary-data-ole-object-temporary-file-external-anal...](./raw-binary-data-ole-object-temporary-file-external-analysis.cs) | `Document`, `Aspose`, `System` | Raw binary data ole object temporary file external analysis |
| [read-file-name-property-ole-package-compare-it-original...](./read-file-name-property-ole-package-compare-it-original-source-file.cs) | `Document`, `System`, `Aspose` | Read file name property ole package compare it original source file |
| [replace-existing-ole-object-doc-file-new-image-ole-obje...](./replace-existing-ole-object-doc-file-new-image-ole-object-insertoleobject.cs) | `Document`, `DocumentBuilder`, `Aspose` | Replace existing ole object doc file new image ole object insertoleobject |
| [retrieve-display-width-height-ole-object-adjust-its-siz...](./retrieve-display-width-height-ole-object-adjust-its-size-after-insertion.cs) | `Aspose`, `Document`, `DocumentBuilder` | Retrieve display width height ole object adjust its size after insertion |
| [retrieve-ole-object-display-width-height-after-insertio...](./retrieve-ole-object-display-width-height-after-insertion-store-dimensions-layout.cs) | `ShapeRenderer`, `Aspose`, `Document` | Retrieve ole object display width height after insertion store dimensions layout |
| [retrieve-progid-inserted-ole-object-log-it-diagnostic-p...](./retrieve-progid-inserted-ole-object-log-it-diagnostic-purposes.cs) | `Document`, `DocumentBuilder`, `Aspose` | Retrieve progid inserted ole object log it diagnostic purposes |
| [retrieve-raw-binary-data-ole-object-via-its-data-proper...](./retrieve-raw-binary-data-ole-object-via-its-data-property-custom-processing.cs) | `Document`, `Aspose`, `System` | Retrieve raw binary data ole object via its data property custom processing |
| [set-file-name-extension-inserted-ole-object-preserve-or...](./set-file-name-extension-inserted-ole-object-preserve-original-metadata.cs) | `OlePackage`, `Document`, `DocumentBuilder` | Set file name extension inserted ole object preserve original metadata |
| [validate-progid-ole-object-before-insertion-avoid-runti...](./validate-progid-ole-object-before-insertion-avoid-runtime-errors.cs) | `Document`, `DocumentBuilder`, `Aspose` | Validate progid ole object before insertion avoid runtime errors |
| [verify-successful-ole-object-insertion-checking-returne...](./verify-successful-ole-object-insertion-checking-returned-object-reference-is-not-null.cs) | `OleFormat`, `Document`, `DocumentBuilder` | Verify successful ole object insertion checking returned object reference is... |

## Category Statistics
- Total examples: 29

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for ole-objects patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082519`
<!-- AUTOGENERATED:END -->
