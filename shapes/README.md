# Shapes Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **Shapes** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category guidance.

## Snapshot

- Category: **Shapes**
- Slug: **shapes**
- Total examples: **35**
- Workflow examples: **35 / 35** use the standard shape workflow

## Category rules that shaped these examples

- Use native Aspose.Words APIs directly.
- Create local sample documents and images when a task needs them.
- Do not assume external files or images already exist.
- Prefer documented `Shape`, `GroupShape`, `DocumentBuilder.InsertShape`, and `InsertImage` workflows.
- Keep validation narrow and task-specific.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
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

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# PowerShell example
Copy-Item ..\shapes\insert-an-image-shape-using-documentbuilder-insertimage-with-specified-size-wrap-type-and.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `insert-an-image-shape-using-documentbuilder-insertimage-with-specified-size-wrap-type-and.cs` | Insert an image shape using DocumentBuilder.InsertImage with specified size, wrap type, and positioning. | shape-workflow | doc | mcp |
| 2 | `insert-a-textbox-shape-with-defined-dimensions-border-style-and-interior-formatting-applie.cs` | Insert a textbox shape with defined dimensions, border style, and interior formatting applied. | shape-workflow | docx | mcp |
| 3 | `insert-a-rectangle-autoshape-set-its-fill-color-to-blue-and-define-line-weight.cs` | Insert a rectangle AutoShape, set its fill color to blue, and define line weight. | shape-workflow | docx | mcp |
| 4 | `insert-an-ole-object-shape-and-lock-its-aspect-ratio-to-preserve-original-proportions.cs` | Insert an OLE object shape and lock its aspect ratio to preserve original proportions. | shape-workflow | docx | mcp |
| 5 | `insert-a-picture-shape-into-a-table-cell-and-enable-islayoutincell-for-proper-layout.cs` | Insert a picture shape into a table cell and enable IsLayoutInCell for proper layout. | shape-workflow | docx | mcp |
| 6 | `lock-aspect-ratio-for-shapes-by-setting-shape-aspectratiolocked-property-to-true.cs` | Lock aspect ratio for shapes by setting Shape.AspectRatioLocked property to true. | shape-workflow | docx | mcp |
| 7 | `set-shape-layout-inside-table-cells-by-configuring-islayoutincell-property-to-true.cs` | Set shape layout inside table cells by configuring IsLayoutInCell property to true. | shape-workflow | docx | mcp |
| 8 | `create-a-custom-snip-corner-rectangle-shape-with-defined-corner-radius-and-light-gray-fill.cs` | Create a custom snip corner rectangle shape with defined corner radius and light gray fill. | shape-workflow | docx | mcp |
| 9 | `retrieve-actual-bounds-of-a-shape-using-shape-getactualbounds-and-log-coordinate-points.cs` | Retrieve actual bounds of a shape using Shape.GetActualBounds and log coordinate points. | shape-workflow | docx | mcp |
| 10 | `set-vertical-anchor-of-a-shape-to-top-middle-or-bottom-based-on-design-requirements.cs` | Set vertical anchor of a shape to top, middle, or bottom based on design requirements. | shape-workflow | docx | mcp |
| 11 | `detect-smartart-shapes-using-shape-issmartart-property-and-replace-them-with-alternative-d.cs` | Detect SmartArt shapes using Shape.IsSmartArt property and replace them with alternative diagrams. | shape-workflow | docx | mcp |
| 12 | `insert-a-horizontal-rule-shape-with-custom-width-thickness-and-color-to-separate-sections.cs` | Insert a horizontal rule shape with custom width, thickness, and color to separate sections. | shape-workflow | docx | mcp |
| 13 | `import-mathml-content-as-shapes-position-them-inline-and-adjust-baseline-alignment.cs` | Import MathML content as shapes, position them inline, and adjust baseline alignment. | shape-workflow | docx | mcp |
| 14 | `create-a-groupshape-instance-add-a-picture-textbox-and-autoshape-then-arrange-them.cs` | Create a GroupShape instance, add a picture, textbox, and AutoShape, then arrange them. | shape-workflow | docx | mcp |
| 15 | `append-a-new-shape-to-an-existing-groupshape-and-update-group-bounds-accordingly.cs` | Append a new shape to an existing GroupShape and update group bounds accordingly. | shape-workflow | docx | mcp |
| 16 | `iterate-through-all-shapes-in-a-document-and-output-each-shape-s-type-using-shapetype-enum.cs` | Iterate through all shapes in a document and output each shape's type using ShapeType enumeration. | shape-workflow | doc | mcp |
| 17 | `change-a-picture-shape-to-an-autoshape-programmatically-by-setting-shapetype-while-preserv.cs` | Change a picture shape to an AutoShape programmatically by setting ShapeType while preserving size. | shape-workflow | docx | mcp |
| 18 | `apply-a-uniform-fill-color-to-every-shape-in-a-document-for-consistent-visual-branding.cs` | Apply a uniform fill color to every shape in a document for consistent visual branding. | shape-workflow | doc | mcp |
| 19 | `export-each-shape-s-image-to-a-separate-file-naming-files-based-on-shape-index-and-type.cs` | Export each shape's image to a separate file, naming files based on shape index and type. | shape-workflow | docx | mcp |
| 20 | `load-a-docx-template-insert-required-shapes-and-save-the-modified-document-as-a-new-docx-f.cs` | Load a DOCX template, insert required shapes, and save the modified document as a new DOCX file. | shape-workflow | docx, doc | mcp |
| 21 | `save-a-document-containing-shapes-as-pdf-while-preserving-exact-layout-and-visual-fidelity.cs` | Save a document containing shapes as PDF while preserving exact layout and visual fidelity. | shape-workflow | doc, pdf | mcp |
| 22 | `batch-process-a-folder-of-docx-files-inserting-a-semi-transparent-watermark-shape-into-eac.cs` | Batch process a folder of DOCX files, inserting a semi‑transparent watermark shape into each document. | shape-workflow | docx, doc | mcp |
| 23 | `move-documentbuilder-cursor-to-a-bookmark-before-inserting-a-shape-at-the-bookmarked-locat.cs` | Move DocumentBuilder cursor to a bookmark before inserting a shape at the bookmarked location. | shape-workflow | doc | mcp |
| 24 | `insert-a-shape-with-relative-positioning-to-the-preceding-paragraph-to-maintain-flow-withi.cs` | Insert a shape with relative positioning to the preceding paragraph to maintain flow within the text. | shape-workflow | docx | mcp |
| 25 | `set-a-shape-s-rotation-angle-to-45-degrees-and-verify-visual-result-in-rendered-document.cs` | Set a shape's rotation angle to 45 degrees and verify visual result in rendered document. | shape-workflow | doc | mcp |
| 26 | `configure-text-wrapping-around-a-shape-on-both-sides-of-the-page-using-shape-wraptype-prop.cs` | Configure text wrapping around a shape on both sides of the page using Shape.WrapType property. | shape-workflow | docx | mcp |
| 27 | `insert-a-shape-inside-a-table-cell-then-adjust-the-cell-s-left-and-right-margins-for-prope.cs` | Insert a shape inside a table cell, then adjust the cell's left and right margins for proper spacing. | shape-workflow | docx | mcp |
| 28 | `clone-an-existing-shape-modify-its-fill-color-and-insert-the-clone-at-a-different-location.cs` | Clone an existing shape, modify its fill color, and insert the clone at a different location. | shape-workflow | docx | mcp |
| 29 | `convert-a-floating-shape-to-an-inline-shape-and-later-revert-it-back-to-floating-using-pro.cs` | Convert a floating shape to an inline shape and later revert it back to floating using properties. | shape-workflow | docx | mcp |
| 30 | `apply-a-hyperlink-to-a-shape-that-points-to-an-external-website-and-test-navigation-functi.cs` | Apply a hyperlink to a shape that points to an external website and test navigation functionality. | shape-workflow | docx | mcp |
| 31 | `set-alternative-text-for-a-shape-to-improve-accessibility-for-screen-readers-and-assistive.cs` | Set alternative text for a shape to improve accessibility for screen readers and assistive technologies. | shape-workflow | docx | mcp |
| 32 | `store-custom-metadata-in-a-shape-s-title-property-for-later-retrieval-during-document-proc.cs` | Store custom metadata in a shape's Title property for later retrieval during document processing. | shape-workflow | doc | mcp |
| 33 | `retrieve-a-shape-s-z-order-index-and-bring-the-shape-to-the-front-of-the-document-layering.cs` | Retrieve a shape's Z‑order index and bring the shape to the front of the document layering. | shape-workflow | doc | mcp |
| 34 | `send-a-shape-to-the-back-of-the-document-s-layering-order-to-ensure-underlying-content-rem.cs` | Send a shape to the back of the document's layering order to ensure underlying content remains visible. | shape-workflow | doc | mcp |
| 35 | `group-multiple-shapes-together-and-apply-a-collective-rotation-of-30-degrees-to-the-entire.cs` | Group multiple shapes together and apply a collective rotation of 30 degrees to the entire group. | shape-workflow | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Creating standalone shapes without inserting them into the document tree

- Symptom: The example compiles, but the created shape never appears in the saved document.
- Fix: Insert shapes through `DocumentBuilder.InsertShape`, `InsertImage`, or append them to a valid parent node.

### Assuming external images already exist

- Symptom: Runtime failures when image insertion code points to a missing file.
- Fix: Create a local placeholder or sample image inside the example before inserting it.

### Overcomplicating group or inline positioning logic

- Symptom: Working shapes fail final validation because of unnecessary structural or layout assumptions.
- Fix: Keep validation focused only on the exact requested size, position, grouping, wrapping, or saved output.

## Notes for maintainers

- The selected file for each task is the verified winner recorded in the batch run.
- This category achieved full success with no primary rules and only narrow corrective guidance when needed.
- Preserve exact file-to-task traceability when updating the category.
- Bootstrap all sample input files locally inside the example when the task refers to an existing asset.
