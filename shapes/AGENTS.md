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
- `using System;` (29/35 files)
- `using System.Drawing;` (14/35 files)
- `using System.IO;` (6/35 files)
- `using Aspose.Words.Saving;` (6/35 files)
- `using System.Linq;` (5/35 files)
- `using Aspose.Words.Tables;` (4/35 files)
- `using Aspose.Words.Rendering;` (2/35 files)
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
| [append-new-shape-existing-groupshape-update-group-bound...](./append-new-shape-existing-groupshape-update-group-bounds-accordingly.cs) | `Shape`, `ShapeType`, `Color` | Append new shape existing groupshape update group bounds accordingly |
| [apply-hyperlink-shape-that-points-external-website-test...](./apply-hyperlink-shape-that-points-external-website-test-navigation-functionality.cs) | `InvalidOperationException`, `Document`, `DocumentBuilder` | Apply hyperlink shape that points external website test navigation functionality |
| [apply-uniform-fill-color-every-shape-document-consisten...](./apply-uniform-fill-color-every-shape-document-consistent-visual-branding.cs) | `Aspose`, `Document`, `Words` | Apply uniform fill color every shape document consistent visual branding |
| [batch-process-folder-docx-files-inserting-semi-transpar...](./batch-process-folder-docx-files-inserting-semi-transparent-watermark-shape-each.cs) | `Document`, `System`, `Aspose` | Batch process folder docx files inserting semi transparent watermark shape each |
| [change-picture-shape-autoshape-programmatically-setting...](./change-picture-shape-autoshape-programmatically-setting-shapetype-while-preserving-size.cs) | `Document`, `Shape`, `Aspose` | Change picture shape autoshape programmatically setting shapetype while prese... |
| [clone-existing-shape-modify-its-fill-color-insert-clone...](./clone-existing-shape-modify-its-fill-color-insert-clone-at-different-location.cs) | `Document`, `DocumentBuilder`, `Aspose` | Clone existing shape modify its fill color insert clone at different location |
| [configure-text-wrapping-around-shape-both-sides-page-sh...](./configure-text-wrapping-around-shape-both-sides-page-shape-wraptype-property.cs) | `Document`, `DocumentBuilder`, `Aspose` | Configure text wrapping around shape both sides page shape wraptype property |
| [convert-floating-shape-inline-shape-later-revert-it-bac...](./convert-floating-shape-inline-shape-later-revert-it-back-floating-properties.cs) | `WrapType`, `Document`, `DocumentBuilder` | Convert floating shape inline shape later revert it back floating properties |
| [custom-snip-corner-rectangle-shape-defined-corner-radiu...](./custom-snip-corner-rectangle-shape-defined-corner-radius-light-gray-fill.cs) | `Aspose`, `Document`, `DocumentBuilder` | Custom snip corner rectangle shape defined corner radius light gray fill |
| [detect-smartart-shapes-shape-issmartart-property-replac...](./detect-smartart-shapes-shape-issmartart-property-replace-them-alternative-diagrams.cs) | `System`, `Document`, `Shape` | Detect smartart shapes shape issmartart property replace them alternative dia... |
| [document-containing-shapes-as-pdf-while-preserving-exac...](./document-containing-shapes-as-pdf-while-preserving-exact-layout-visual-fidelity.cs) | `Aspose`, `Document`, `DocumentBuilder` | Document containing shapes as pdf while preserving exact layout visual fidelity |
| [docx-template-insert-required-shapes-modified-document-...](./docx-template-insert-required-shapes-modified-document-as-new-docx-file.cs) | `System`, `Drawing`, `Color` | Docx template insert required shapes modified document as new docx file |
| [export-each-shape-s-image-separate-file-naming-files-ba...](./export-each-shape-s-image-separate-file-naming-files-based-shape-index-type.cs) | `Aspose`, `Document`, `System` | Export each shape s image separate file naming files based shape index type |
| [group-multiple-shapes-together-apply-collective-rotatio...](./group-multiple-shapes-together-apply-collective-rotation-30-degrees-entire-group.cs) | `Color`, `Document`, `DocumentBuilder` | Group multiple shapes together apply collective rotation 30 degrees entire group |
| [groupshape-instance-add-picture-textbox-autoshape-then-...](./groupshape-instance-add-picture-textbox-autoshape-then-arrange-them.cs) | `Document`, `DocumentBuilder`, `Shape` | Groupshape instance add picture textbox autoshape then arrange them |
| [import-mathml-content-as-shapes-position-them-inline-ad...](./import-mathml-content-as-shapes-position-them-inline-adjust-baseline-alignment.cs) | `Aspose`, `Words`, `HtmlLoadOptions` | Import mathml content as shapes position them inline adjust baseline alignment |
| [insert-horizontal-rule-shape-custom-width-thickness-col...](./insert-horizontal-rule-shape-custom-width-thickness-color-separate-sections.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert horizontal rule shape custom width thickness color separate sections |
| [insert-image-shape-documentbuilder-insertimage-specifie...](./insert-image-shape-documentbuilder-insertimage-specified-size-wrap-type-positioning.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert image shape documentbuilder insertimage specified size wrap type posit... |
| [insert-ole-object-shape-lock-its-aspect-ratio-preserve-...](./insert-ole-object-shape-lock-its-aspect-ratio-preserve-original-proportions.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert ole object shape lock its aspect ratio preserve original proportions |
| [insert-picture-shape-table-cell-enable-islayoutincell-p...](./insert-picture-shape-table-cell-enable-islayoutincell-proper-layout.cs) | `Aspose`, `Document`, `DocumentBuilder` | Insert picture shape table cell enable islayoutincell proper layout |
| [insert-rectangle-autoshape-set-its-fill-color-blue-defi...](./insert-rectangle-autoshape-set-its-fill-color-blue-define-line-weight.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert rectangle autoshape set its fill color blue define line weight |
| [insert-shape-inside-table-cell-then-adjust-cell-s-left-...](./insert-shape-inside-table-cell-then-adjust-cell-s-left-right-margins-proper-spacing.cs) | `Aspose`, `Document`, `DocumentBuilder` | Insert shape inside table cell then adjust cell s left right margins proper s... |
| [insert-shape-relative-positioning-preceding-paragraph-m...](./insert-shape-relative-positioning-preceding-paragraph-maintain-flow-within-text.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert shape relative positioning preceding paragraph maintain flow within text |
| [insert-textbox-shape-defined-dimensions-border-style-in...](./insert-textbox-shape-defined-dimensions-border-style-interior-formatting-applied.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert textbox shape defined dimensions border style interior formatting applied |
| [iterate-through-all-shapes-document-output-each-shape-s...](./iterate-through-all-shapes-document-output-each-shape-s-type-shapetype-enumeration.cs) | `Document`, `Aspose`, `Words` | Iterate through all shapes document output each shape s type shapetype enumer... |
| [lock-aspect-ratio-shapes-setting-shape-aspectratiolocke...](./lock-aspect-ratio-shapes-setting-shape-aspectratiolocked-property-true.cs) | `Document`, `DocumentBuilder`, `Aspose` | Lock aspect ratio shapes setting shape aspectratiolocked property true |
| [move-documentbuilder-cursor-bookmark-before-inserting-s...](./move-documentbuilder-cursor-bookmark-before-inserting-shape-at-bookmarked-location.cs) | `Document`, `DocumentBuilder`, `Aspose` | Move documentbuilder cursor bookmark before inserting shape at bookmarked loc... |
| [retrieve-actual-bounds-shape-shape-getactualbounds-log-...](./retrieve-actual-bounds-shape-shape-getactualbounds-log-coordinate-points.cs) | `Aspose`, `Document`, `DocumentBuilder` | Retrieve actual bounds shape shape getactualbounds log coordinate points |
| [retrieve-shape-s-z-order-index-bring-shape-front-docume...](./retrieve-shape-s-z-order-index-bring-shape-front-document-layering.cs) | `Document`, `Aspose`, `System` | Retrieve shape s z order index bring shape front document layering |
| [send-shape-back-document-s-layering-order-ensure-underl...](./send-shape-back-document-s-layering-order-ensure-underlying-content-remains-visible.cs) | `ShapeType`, `RelativeHorizontalPosition`, `RelativeVerticalPosition` | Send shape back document s layering order ensure underlying content remains v... |
| ... | | *and 5 more files* |

## Category Statistics
- Total examples: 35

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for shapes patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082610`
<!-- AUTOGENERATED:END -->
