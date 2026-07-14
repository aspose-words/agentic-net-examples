---
name: shapes
description: Verified C# examples for Shapes scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Shapes

## Purpose

This folder is a live, curated example set for Shapes scenarios. Treat every `.cs` file as a standalone console application. The goal is correct, warning-free examples that use documented Aspose APIs and match the original task intent.

## Non-negotiable conventions

- Always use documented Aspose.Words APIs directly.
- Always create local sample source documents or images when a task refers to an existing file, folder, stream, template, or input asset.
- Prefer `DocumentBuilder.InsertShape`, `InsertImage`, and documented `Shape` workflows.
- Keep validation narrow and task-specific.
- Do not invent shape, grouping, or rendering helper APIs.

## Recommended workflow selection

- Shape Workflow workflow: 35 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. API usage must be supported by the configured package versions.
3. Exported outputs must actually be written by the example.
4. Validation scenarios must inspect only the behavior requested by the task.
5. Examples that depend on files, folders, streams, images, or data should bootstrap those inputs locally during the example run.

## File-to-task reference

- `insert-an-image-shape-using-documentbuilder-insertimage-with-specified-size-wrap-type-and.cs`
  - Task: Insert an image shape using DocumentBuilder.InsertImage with specified size, wrap type, and positioning.
  - Workflow: Shape Workflow
  - Outputs: doc
  - Selected engine: mcp
- `insert-a-textbox-shape-with-defined-dimensions-border-style-and-interior-formatting-applie.cs`
  - Task: Insert a textbox shape with defined dimensions, border style, and interior formatting applied.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-rectangle-autoshape-set-its-fill-color-to-blue-and-define-line-weight.cs`
  - Task: Insert a rectangle AutoShape, set its fill color to blue, and define line weight.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-an-ole-object-shape-and-lock-its-aspect-ratio-to-preserve-original-proportions.cs`
  - Task: Insert an OLE object shape and lock its aspect ratio to preserve original proportions.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-picture-shape-into-a-table-cell-and-enable-islayoutincell-for-proper-layout.cs`
  - Task: Insert a picture shape into a table cell and enable IsLayoutInCell for proper layout.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `lock-aspect-ratio-for-shapes-by-setting-shape-aspectratiolocked-property-to-true.cs`
  - Task: Lock aspect ratio for shapes by setting Shape.AspectRatioLocked property to true.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `set-shape-layout-inside-table-cells-by-configuring-islayoutincell-property-to-true.cs`
  - Task: Set shape layout inside table cells by configuring IsLayoutInCell property to true.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-custom-snip-corner-rectangle-shape-with-defined-corner-radius-and-light-gray-fill.cs`
  - Task: Create a custom snip corner rectangle shape with defined corner radius and light gray fill.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: existing_repo
- `retrieve-actual-bounds-of-a-shape-using-shape-getactualbounds-and-log-coordinate-points.cs`
  - Task: Retrieve actual bounds of a shape using Shape.GetActualBounds and log coordinate points.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `set-vertical-anchor-of-a-shape-to-top-middle-or-bottom-based-on-design-requirements.cs`
  - Task: Set vertical anchor of a shape to top, middle, or bottom based on design requirements.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `detect-smartart-shapes-using-shape-issmartart-property-and-replace-them-with-alternative-d.cs`
  - Task: Detect SmartArt shapes using Shape.IsSmartArt property and replace them with alternative diagrams.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: llm
- `insert-a-horizontal-rule-shape-with-custom-width-thickness-and-color-to-separate-sections.cs`
  - Task: Insert a horizontal rule shape with custom width, thickness, and color to separate sections.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `import-mathml-content-as-shapes-position-them-inline-and-adjust-baseline-alignment.cs`
  - Task: Import MathML content as shapes, position them inline, and adjust baseline alignment.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-groupshape-instance-add-a-picture-textbox-and-autoshape-then-arrange-them.cs`
  - Task: Create a GroupShape instance, add a picture, textbox, and AutoShape, then arrange them.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `append-a-new-shape-to-an-existing-groupshape-and-update-group-bounds-accordingly.cs`
  - Task: Append a new shape to an existing GroupShape and update group bounds accordingly.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `iterate-through-all-shapes-in-a-document-and-output-each-shape-s-type-using-shapetype-enum.cs`
  - Task: Iterate through all shapes in a document and output each shape's type using ShapeType enumeration.
  - Workflow: Shape Workflow
  - Outputs: doc
  - Selected engine: mcp
- `change-a-picture-shape-to-an-autoshape-programmatically-by-setting-shapetype-while-preserv.cs`
  - Task: Change a picture shape to an AutoShape programmatically by setting ShapeType while preserving size.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-uniform-fill-color-to-every-shape-in-a-document-for-consistent-visual-branding.cs`
  - Task: Apply a uniform fill color to every shape in a document for consistent visual branding.
  - Workflow: Shape Workflow
  - Outputs: doc
  - Selected engine: mcp
- `export-each-shape-s-image-to-a-separate-file-naming-files-based-on-shape-index-and-type.cs`
  - Task: Export each shape's image to a separate file, naming files based on shape index and type.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `load-a-docx-template-insert-required-shapes-and-save-the-modified-document-as-a-new-docx-f.cs`
  - Task: Load a DOCX template, insert required shapes, and save the modified document as a new DOCX file.
  - Workflow: Shape Workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `save-a-document-containing-shapes-as-pdf-while-preserving-exact-layout-and-visual-fidelity.cs`
  - Task: Save a document containing shapes as PDF while preserving exact layout and visual fidelity.
  - Workflow: Shape Workflow
  - Outputs: doc, pdf
  - Selected engine: mcp
- `batch-process-a-folder-of-docx-files-inserting-a-semi-transparent-watermark-shape-into-eac.cs`
  - Task: Batch process a folder of DOCX files, inserting a semi-transparent watermark shape into each document.
  - Workflow: Shape Workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `move-documentbuilder-cursor-to-a-bookmark-before-inserting-a-shape-at-the-bookmarked-locat.cs`
  - Task: Move DocumentBuilder cursor to a bookmark before inserting a shape at the bookmarked location.
  - Workflow: Shape Workflow
  - Outputs: doc
  - Selected engine: mcp
- `insert-a-shape-with-relative-positioning-to-the-preceding-paragraph-to-maintain-flow-withi.cs`
  - Task: Insert a shape with relative positioning to the preceding paragraph to maintain flow within the text.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `set-a-shape-s-rotation-angle-to-45-degrees-and-verify-visual-result-in-rendered-document.cs`
  - Task: Set a shape's rotation angle to 45 degrees and verify visual result in rendered document.
  - Workflow: Shape Workflow
  - Outputs: doc
  - Selected engine: mcp
- `configure-text-wrapping-around-a-shape-on-both-sides-of-the-page-using-shape-wraptype-prop.cs`
  - Task: Configure text wrapping around a shape on both sides of the page using Shape.WrapType property.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-shape-inside-a-table-cell-then-adjust-the-cell-s-left-and-right-margins-for-prope.cs`
  - Task: Insert a shape inside a table cell, then adjust the cell's left and right margins for proper spacing.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `clone-an-existing-shape-modify-its-fill-color-and-insert-the-clone-at-a-different-location.cs`
  - Task: Clone an existing shape, modify its fill color, and insert the clone at a different location.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `convert-a-floating-shape-to-an-inline-shape-and-later-revert-it-back-to-floating-using-pro.cs`
  - Task: Convert a floating shape to an inline shape and later revert it back to floating using properties.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-hyperlink-to-a-shape-that-points-to-an-external-website-and-test-navigation-functi.cs`
  - Task: Apply a hyperlink to a shape that points to an external website and test navigation functionality.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `set-alternative-text-for-a-shape-to-improve-accessibility-for-screen-readers-and-assistive.cs`
  - Task: Set alternative text for a shape to improve accessibility for screen readers and assistive technologies.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp
- `store-custom-metadata-in-a-shape-s-title-property-for-later-retrieval-during-document-proc.cs`
  - Task: Store custom metadata in a shape's Title property for later retrieval during document processing.
  - Workflow: Shape Workflow
  - Outputs: doc
  - Selected engine: mcp
- `retrieve-a-shape-s-z-order-index-and-bring-the-shape-to-the-front-of-the-document-layering.cs`
  - Task: Retrieve a shape's Z-order index and bring the shape to the front of the document layering.
  - Workflow: Shape Workflow
  - Outputs: doc
  - Selected engine: mcp
- `send-a-shape-to-the-back-of-the-document-s-layering-order-to-ensure-underlying-content-rem.cs`
  - Task: Send a shape to the back of the document's layering order to ensure underlying content remains visible.
  - Workflow: Shape Workflow
  - Outputs: doc
  - Selected engine: mcp
- `group-multiple-shapes-together-and-apply-a-collective-rotation-of-30-degrees-to-the-entire.cs`
  - Task: Group multiple shapes together and apply a collective rotation of 30 degrees to the entire group.
  - Workflow: Shape Workflow
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- Unsupported API invention
  - Symptom: Generated code references members that do not exist in the selected package version.
  - Preferred fix: Replace invented members with documented Aspose.Words APIs already proven in this category.

- Missing local bootstrap inputs
  - Symptom: The example assumes source files, folders, images, or data already exist.
  - Preferred fix: Create deterministic local inputs before loading, processing, or validating them.

- Over-broad validation
  - Symptom: The example fails at runtime while checking unrelated document internals.
  - Preferred fix: Validate only the requested behavior and the existence of expected outputs.

## Build and run contract

- Target framework: `net8.0`
- Package: `Aspose.Words` `26.5.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.5.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\shapes\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer documented Aspose APIs over speculative shortcuts.
