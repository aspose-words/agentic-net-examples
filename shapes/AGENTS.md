---
name: shapes
description: Verified C# examples for shapes scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Shapes

## Purpose

Provide an AI coding assistant with concise, reliable guidance for creating, configuring, and manipulating shapes in Word documents using Aspose.Words 26.10.0 while avoiding common pitfalls and unsupported APIs.

## Non-negotiable conventions

- Never reference System.Drawing types such as Bitmap, Graphics, or Color unless the System.Drawing.Common package is explicitly added and supported on the target platform.
- Do not use Shape members that are not present in Aspose.Words 26.10.0, such as LockAspectRatio, GetActualBounds, CornerRadius, UpdateGroupShapeBounds, BringToFront, SendToBack, Hyperlink, or TextBoxContent.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Shape Workflow workflow**: 35 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. After each operation, verify that the document file exists and that any shape-specific property (e.g., IsLayoutInCell, AspectRatioLocked) has the expected value.

## File-to-task reference

- `insert-an-image-shape-using-documentbuilder-insertimage-with-specified-size-wrap-type-and.cs`
  - Task: Insert an image shape using DocumentBuilder.InsertImage with specified size, wrap type, and positioning.
  - Workflow: shape-workflow
  - Outputs: doc
  - Selected engine: mcp
- `insert-a-textbox-shape-with-defined-dimensions-border-style-and-interior-formatting-applie.cs`
  - Task: Insert a textbox shape with defined dimensions, border style, and interior formatting applied.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-rectangle-autoshape-set-its-fill-color-to-blue-and-define-line-weight.cs`
  - Task: Insert a rectangle AutoShape, set its fill color to blue, and define line weight.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-an-ole-object-shape-and-lock-its-aspect-ratio-to-preserve-original-proportions.cs`
  - Task: Insert an OLE object shape and lock its aspect ratio to preserve original proportions.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: llm
- `insert-a-picture-shape-into-a-table-cell-and-enable-islayoutincell-for-proper-layout.cs`
  - Task: Insert a picture shape into a table cell and enable IsLayoutInCell for proper layout.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `lock-aspect-ratio-for-shapes-by-setting-shape-aspectratiolocked-property-to-true.cs`
  - Task: Lock aspect ratio for shapes by setting Shape.AspectRatioLocked property to true.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `set-shape-layout-inside-table-cells-by-configuring-islayoutincell-property-to-true.cs`
  - Task: Set shape layout inside table cells by configuring IsLayoutInCell property to true.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-custom-snip-corner-rectangle-shape-with-defined-corner-radius-and-light-gray-fill.cs`
  - Task: Create a custom snip corner rectangle shape with defined corner radius and light gray fill.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-actual-bounds-of-a-shape-using-shape-getactualbounds-and-log-coordinate-points.cs`
  - Task: Retrieve actual bounds of a shape using Shape.GetActualBounds and log coordinate points.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: llm
- `set-vertical-anchor-of-a-shape-to-top-middle-or-bottom-based-on-design-requirements.cs`
  - Task: Set vertical anchor of a shape to top, middle, or bottom based on design requirements.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `detect-smartart-shapes-using-shape-issmartart-property-and-replace-them-with-alternative-d.cs`
  - Task: Detect SmartArt shapes using Shape.IsSmartArt property and replace them with alternative diagrams.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-horizontal-rule-shape-with-custom-width-thickness-and-color-to-separate-sections.cs`
  - Task: Insert a horizontal rule shape with custom width, thickness, and color to separate sections.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `import-mathml-content-as-shapes-position-them-inline-and-adjust-baseline-alignment.cs`
  - Task: Import MathML content as shapes, position them inline, and adjust baseline alignment.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-groupshape-instance-add-a-picture-textbox-and-autoshape-then-arrange-them.cs`
  - Task: Create a GroupShape instance, add a picture, textbox, and AutoShape, then arrange them.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: llm
- `append-a-new-shape-to-an-existing-groupshape-and-update-group-bounds-accordingly.cs`
  - Task: Append a new shape to an existing GroupShape and update group bounds accordingly.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `iterate-through-all-shapes-in-a-document-and-output-each-shape-s-type-using-shapetype-enum.cs`
  - Task: Iterate through all shapes in a document and output each shape's type using ShapeType enumeration.
  - Workflow: shape-workflow
  - Outputs: doc
  - Selected engine: mcp
- `change-a-picture-shape-to-an-autoshape-programmatically-by-setting-shapetype-while-preserv.cs`
  - Task: Change a picture shape to an AutoShape programmatically by setting ShapeType while preserving size.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: llm
- `apply-a-uniform-fill-color-to-every-shape-in-a-document-for-consistent-visual-branding.cs`
  - Task: Apply a uniform fill color to every shape in a document for consistent visual branding.
  - Workflow: shape-workflow
  - Outputs: doc
  - Selected engine: mcp
- `export-each-shape-s-image-to-a-separate-file-naming-files-based-on-shape-index-and-type.cs`
  - Task: Export each shape's image to a separate file, naming files based on shape index and type.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `load-a-docx-template-insert-required-shapes-and-save-the-modified-document-as-a-new-docx-f.cs`
  - Task: Load a DOCX template, insert required shapes, and save the modified document as a new DOCX file.
  - Workflow: shape-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `save-a-document-containing-shapes-as-pdf-while-preserving-exact-layout-and-visual-fidelity.cs`
  - Task: Save a document containing shapes as PDF while preserving exact layout and visual fidelity.
  - Workflow: shape-workflow
  - Outputs: doc, pdf
  - Selected engine: mcp
- `batch-process-a-folder-of-docx-files-inserting-a-semi-transparent-watermark-shape-into-eac.cs`
  - Task: Batch process a folder of DOCX files, inserting a semi-transparent watermark shape into each document.
  - Workflow: shape-workflow
  - Outputs: docx, doc
  - Selected engine: llm
- `move-documentbuilder-cursor-to-a-bookmark-before-inserting-a-shape-at-the-bookmarked-locat.cs`
  - Task: Move DocumentBuilder cursor to a bookmark before inserting a shape at the bookmarked location.
  - Workflow: shape-workflow
  - Outputs: doc
  - Selected engine: mcp
- `insert-a-shape-with-relative-positioning-to-the-preceding-paragraph-to-maintain-flow-withi.cs`
  - Task: Insert a shape with relative positioning to the preceding paragraph to maintain flow within the text.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `set-a-shape-s-rotation-angle-to-45-degrees-and-verify-visual-result-in-rendered-document.cs`
  - Task: Set a shape's rotation angle to 45 degrees and verify visual result in rendered document.
  - Workflow: shape-workflow
  - Outputs: doc
  - Selected engine: mcp
- `configure-text-wrapping-around-a-shape-on-both-sides-of-the-page-using-shape-wraptype-prop.cs`
  - Task: Configure text wrapping around a shape on both sides of the page using Shape.WrapType property.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-shape-inside-a-table-cell-then-adjust-the-cell-s-left-and-right-margins-for-prope.cs`
  - Task: Insert a shape inside a table cell, then adjust the cell's left and right margins for proper spacing.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `clone-an-existing-shape-modify-its-fill-color-and-insert-the-clone-at-a-different-location.cs`
  - Task: Clone an existing shape, modify its fill color, and insert the clone at a different location.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `convert-a-floating-shape-to-an-inline-shape-and-later-revert-it-back-to-floating-using-pro.cs`
  - Task: Convert a floating shape to an inline shape and later revert it back to floating using properties.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-hyperlink-to-a-shape-that-points-to-an-external-website-and-test-navigation-functi.cs`
  - Task: Apply a hyperlink to a shape that points to an external website and test navigation functionality.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: existing_repo
- `set-alternative-text-for-a-shape-to-improve-accessibility-for-screen-readers-and-assistive.cs`
  - Task: Set alternative text for a shape to improve accessibility for screen readers and assistive technologies.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp
- `store-custom-metadata-in-a-shape-s-title-property-for-later-retrieval-during-document-proc.cs`
  - Task: Store custom metadata in a shape's Title property for later retrieval during document processing.
  - Workflow: shape-workflow
  - Outputs: doc
  - Selected engine: mcp
- `retrieve-a-shape-s-z-order-index-and-bring-the-shape-to-the-front-of-the-document-layering.cs`
  - Task: Retrieve a shape's Z-order index and bring the shape to the front of the document layering.
  - Workflow: shape-workflow
  - Outputs: doc
  - Selected engine: mcp
- `send-a-shape-to-the-back-of-the-document-s-layering-order-to-ensure-underlying-content-rem.cs`
  - Task: Send a shape to the back of the document's layering order to ensure underlying content remains visible.
  - Workflow: shape-workflow
  - Outputs: doc
  - Selected engine: mcp
- `group-multiple-shapes-together-and-apply-a-collective-rotation-of-30-degrees-to-the-entire.cs`
  - Task: Group multiple shapes together and apply a collective rotation of 30 degrees to the entire group.
  - Workflow: shape-workflow
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Missing System.Drawing reference**
  - Symptom: Compilation fails with CS1069 errors for Bitmap, Graphics, or Color types.
  - Root cause: The code imports System.Drawing types that are not available without the System.Drawing.Common package, which is unsupported on non-Windows runtimes.
  - Preferred fix: Replace System.Drawing usage with raw image byte arrays or use Aspose.Words image insertion APIs that accept a file path or stream.
  - Prevention: Never add System.Drawing namespaces in examples; generate sample images as byte arrays or embed pre-created PNG data.

- **Use of non-existent Shape members**
  - Symptom: Compilation fails with CS1061 errors for members such as LockAspectRatio, GetActualBounds, CornerRadius, UpdateGroupShapeBounds, BringToFront, SendToBack, Hyperlink, or TextBoxContent.
  - Root cause: The code references APIs that are not part of Aspose.Words 26.10.0.
  - Preferred fix: Replace invented members with documented alternatives (e.g., AspectRatioLocked, Bounds, FillColor, StrokeColor, HyperlinkInfo via Shape.HyperlinkInfo property if needed).
  - Prevention: Consult the official Aspose.Words API reference for the exact version before using any Shape property or method.

- **Invalid DocumentBuilder table methods**
  - Symptom: Compilation fails with CS1061 errors for EndCell, EndRow, or similar methods.
  - Root cause: Older examples use DocumentBuilder.EndCell/EndRow which were removed in recent versions.
  - Preferred fix: Construct tables using the Table, Row, and Cell classes and append them to the document body; use builder.MoveTo(cell.FirstParagraph) to position the cursor.
  - Prevention: Follow the current table-building pattern documented for Aspose.Words 26.x.

- **Incorrect InsertOleObject overload**
  - Symptom: Compilation fails with CS1501 or CS1503 errors because the arguments do not match any overload.
  - Root cause: The code supplies a file path string or mismatched parameter types to InsertOleObject.
  - Preferred fix: Use the overload that takes (Stream oleStream, string progId, bool isIconic, Stream? iconStream) and pass a FileStream for the OLE content.
  - Prevention: Check the method signature in the API docs and always provide a Stream for the OLE data.

- **Unsupported positioning enums**
  - Symptom: Compilation fails with CS0246 or CS0234 errors for VerticalAnchor, RelativeVerticalPosition, etc.
  - Root cause: The code references enums that are not defined in Aspose.Words.Drawing for the current version.
  - Preferred fix: Use RelativeHorizontalPosition and RelativeVerticalPosition enums together with Left and Top properties to control shape placement.
  - Prevention: Reference only the positioning enums listed in the official documentation.

- **Invalid node insertion location**
  - Symptom: Runtime throws System.ArgumentException: Cannot insert a node of this type at this location.
  - Root cause: The builder attempts to insert a shape when the cursor is not positioned inside a valid paragraph or cell.
  - Preferred fix: Move the DocumentBuilder to a paragraph within the target cell (builder.MoveTo(cell.FirstParagraph)) before calling InsertShape or InsertImage.
  - Prevention: Always ensure the insertion point is a CompositeNode that accepts Shape children.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.10.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
```

```powershell
Copy-Item ..\shapes\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Periodically review Aspose.Words release notes to update the list of supported Shape members and adjust examples accordingly.
- Run the verification suite after each library upgrade to catch newly introduced API changes.
