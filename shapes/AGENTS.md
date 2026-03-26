---
name: shapes
description: C# examples for shapes using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - shapes

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **shapes** category.
This folder contains standalone C# examples for shapes operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **shapes**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (35/35 files) ← category-specific
- `using Aspose.Words.Drawing;` (35/35 files)
- `using System;` (30/35 files)
- `using System.IO;` (15/35 files)
- `using System.Drawing;` (13/35 files)
- `using Aspose.Words.Tables;` (5/35 files)
- `using Aspose.Words.Saving;` (5/35 files)
- `using System.Linq;` (5/35 files)
- `using Aspose.Words.Rendering;` (2/35 files)
- `using System.Text;` (1/35 files)
- `using Aspose.Words.Loading;` (1/35 files)

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
| [append-new-shape-existing-groupshape-update-group-bound...](./append-new-shape-existing-groupshape-update-group-bounds-accordingly.cs) | `DocumentBuilder`, `ShapeType`, `Color` | Append new shape existing groupshape update group bounds accordingly |
| [apply-hyperlink-shape-that-points-external-website-test...](./apply-hyperlink-shape-that-points-external-website-test-navigation-functionality.cs) | `InvalidOperationException`, `Document`, `DocumentBuilder` | Apply hyperlink shape that points external website test navigation functionality |
| [apply-uniform-fill-color-every-shape-document-consisten...](./apply-uniform-fill-color-every-shape-document-consistent-visual-branding.cs) | `Document`, `Input`, `NodeType` | Apply uniform fill color every shape document consistent visual branding |
| [batch-process-folder-docx-files-inserting-semi-transpar...](./batch-process-folder-docx-files-inserting-semi-transparent-watermark-shape-each.cs) | `Document`, `Color`, `TextWatermarkOptions` | Batch process folder docx files inserting semi transparent watermark shape each |
| [change-picture-shape-autoshape-programmatically-setting...](./change-picture-shape-autoshape-programmatically-setting-shapetype-while-preserving-size.cs) | `Document`, `DocumentBuilder`, `Shape` | Change picture shape autoshape programmatically setting shapetype while prese... |
| [clone-existing-shape-modify-its-fill-color-insert-clone...](./clone-existing-shape-modify-its-fill-color-insert-clone-at-different-location.cs) | `Document`, `DocumentBuilder`, `Color` | Clone existing shape modify its fill color insert clone at different location |
| [configure-text-wrapping-around-shape-both-sides-page-sh...](./configure-text-wrapping-around-shape-both-sides-page-shape-wraptype-property.cs) | `Document`, `DocumentBuilder`, `ShapeType` | Configure text wrapping around shape both sides page shape wraptype property |
| [convert-floating-shape-inline-shape-later-revert-it-bac...](./convert-floating-shape-inline-shape-later-revert-it-back-floating-properties.cs) | `WrapType`, `Document`, `DocumentBuilder` | Convert floating shape inline shape later revert it back floating properties |
| [custom-snip-corner-rectangle-shape-defined-corner-radiu...](./custom-snip-corner-rectangle-shape-defined-corner-radius-light-gray-fill.cs) | `Document`, `DocumentBuilder`, `OoxmlSaveOptions` | Custom snip corner rectangle shape defined corner radius light gray fill |
| [detect-smartart-shapes-shape-issmartart-property-replac...](./detect-smartart-shapes-shape-issmartart-property-replace-them-alternative-diagrams.cs) | `Document`, `Shape`, `DocumentBuilder` | Detect smartart shapes shape issmartart property replace them alternative dia... |
| [document-containing-shapes-as-pdf-while-preserving-exac...](./document-containing-shapes-as-pdf-while-preserving-exact-layout-visual-fidelity.cs) | `Document`, `DocumentBuilder`, `PdfSaveOptions` | Document containing shapes as pdf while preserving exact layout visual fidelity |
| [docx-template-insert-required-shapes-modified-document-...](./docx-template-insert-required-shapes-modified-document-as-new-docx-file.cs) | `Drawing`, `Color`, `Document` | Docx template insert required shapes modified document as new docx file |
| [export-each-shape-s-image-separate-file-naming-files-ba...](./export-each-shape-s-image-separate-file-naming-files-based-shape-index-type.cs) | `Document`, `ImageData`, `NodeType` | Export each shape s image separate file naming files based shape index type |
| [group-multiple-shapes-together-apply-collective-rotatio...](./group-multiple-shapes-together-apply-collective-rotation-30-degrees-entire-group.cs) | `Color`, `Document`, `DocumentBuilder` | Group multiple shapes together apply collective rotation 30 degrees entire group |
| [groupshape-instance-add-picture-textbox-autoshape-then-...](./groupshape-instance-add-picture-textbox-autoshape-then-arrange-them.cs) | `Drawing`, `Color`, `Document` | Groupshape instance add picture textbox autoshape then arrange them |
| [import-mathml-content-as-shapes-position-them-inline-ad...](./import-mathml-content-as-shapes-position-them-inline-adjust-baseline-alignment.cs) | `Document`, `HtmlLoadOptions`, `MathML` | Import mathml content as shapes position them inline adjust baseline alignment |
| [insert-horizontal-rule-shape-custom-width-thickness-col...](./insert-horizontal-rule-shape-custom-width-thickness-color-separate-sections.cs) | `Document`, `DocumentBuilder`, `HorizontalRuleAlignment` | Insert horizontal rule shape custom width thickness color separate sections |
| [insert-image-shape-documentbuilder-insertimage-specifie...](./insert-image-shape-documentbuilder-insertimage-specified-size-wrap-type-positioning.cs) | `Document`, `DocumentBuilder`, `Guid` | Insert image shape documentbuilder insertimage specified size wrap type posit... |
| [insert-ole-object-shape-lock-its-aspect-ratio-preserve-...](./insert-ole-object-shape-lock-its-aspect-ratio-preserve-original-proportions.cs) | `Document`, `DocumentBuilder`, `OleObjectLocked` | Insert ole object shape lock its aspect ratio preserve original proportions |
| [insert-picture-shape-table-cell-enable-islayoutincell-p...](./insert-picture-shape-table-cell-enable-islayoutincell-proper-layout.cs) | `Document`, `DocumentBuilder`, `NodeType` | Insert picture shape table cell enable islayoutincell proper layout |
| [insert-rectangle-autoshape-set-its-fill-color-blue-defi...](./insert-rectangle-autoshape-set-its-fill-color-blue-define-line-weight.cs) | `Document`, `DocumentBuilder`, `WrapType` | Insert rectangle autoshape set its fill color blue define line weight |
| [insert-shape-inside-table-cell-then-adjust-cell-s-left-...](./insert-shape-inside-table-cell-then-adjust-cell-s-left-right-margins-proper-spacing.cs) | `Document`, `DocumentBuilder`, `CellFormat` | Insert shape inside table cell then adjust cell s left right margins proper s... |
| [insert-shape-relative-positioning-preceding-paragraph-m...](./insert-shape-relative-positioning-preceding-paragraph-maintain-flow-within-text.cs) | `Document`, `DocumentBuilder`, `Color` | Insert shape relative positioning preceding paragraph maintain flow within text |
| [insert-textbox-shape-defined-dimensions-border-style-in...](./insert-textbox-shape-defined-dimensions-border-style-interior-formatting-applied.cs) | `Document`, `DocumentBuilder`, `Color` | Insert textbox shape defined dimensions border style interior formatting applied |
| [iterate-through-all-shapes-document-output-each-shape-s...](./iterate-through-all-shapes-document-output-each-shape-s-type-shapetype-enumeration.cs) | `Document`, `Input`, `NodeType` | Iterate through all shapes document output each shape s type shapetype enumer... |
| [lock-aspect-ratio-shapes-setting-shape-aspectratiolocke...](./lock-aspect-ratio-shapes-setting-shape-aspectratiolocked-property-true.cs) | `Document`, `DocumentBuilder`, `Shape` | Lock aspect ratio shapes setting shape aspectratiolocked property true |
| [move-documentbuilder-cursor-bookmark-before-inserting-s...](./move-documentbuilder-cursor-bookmark-before-inserting-shape-at-bookmarked-location.cs) | `Document`, `DocumentBuilder`, `ShapeType` | Move documentbuilder cursor bookmark before inserting shape at bookmarked loc... |
| [retrieve-actual-bounds-shape-shape-getactualbounds-log-...](./retrieve-actual-bounds-shape-shape-getactualbounds-log-coordinate-points.cs) | `Document`, `DocumentBuilder`, `ShapeType` | Retrieve actual bounds shape shape getactualbounds log coordinate points |
| [retrieve-shape-s-z-order-index-bring-shape-front-docume...](./retrieve-shape-s-z-order-index-bring-shape-front-document-layering.cs) | `Document`, `NodeType` | Retrieve shape s z order index bring shape front document layering |
| [send-shape-back-document-s-layering-order-ensure-underl...](./send-shape-back-document-s-layering-order-ensure-underlying-content-remains-visible.cs) | `ShapeType`, `RelativeHorizontalPosition`, `RelativeVerticalPosition` | Send shape back document s layering order ensure underlying content remains v... |
| ... | | *and 5 more files* |

## Category Statistics
- Total examples: 35

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for shapes patterns.


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
Copy-Item ..\shapes\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

### Notes for Agents and Developers

- Treat every `.cs` file in `shapes/` as a full console program, not a snippet.
- Run one file at a time by copying it to `Program.cs`.
- If a sample needs input documents, images, fonts, or data files, place them in the temporary project directory before running.
- See the root [AGENTS.md](../AGENTS.md) for repository-wide prerequisites, project file template, and testing guidance.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
