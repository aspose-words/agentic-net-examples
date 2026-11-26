# Shapes Examples for Aspose.Words for .NET

The Shapes category demonstrates how to create, configure, and manipulate drawing objects in Word documents using Aspose.Words 26.11.0. Verified examples cover image insertion, textbox creation, OLE embedding, table-cell layout, custom rectangles, bounds retrieval, vertical anchoring, SmartArt replacement, horizontal rules, and MathML rendering.

## Snapshot

- Category: Shapes
- Slug: shapes
- Total examples: 35
- Publish-ready successful examples: 35 / 35
- Source run: 20261124_190740_1c00f5
- Shape Workflow examples: 35

## Category rules that shaped these examples

- Never call members that are not defined in Aspose.Words 26.11.0, such as Shape.StrokeDashStyle, Shape.GetActualBounds, or ShapeType.SnipCornerRectangle.
- Do not reference System.Drawing types without adding the System.Drawing.Common package.
- Insert shapes with DocumentBuilder.InsertShape or InsertImage and then adjust properties like Width, Height, WrapType, RelativeHorizontalPosition, and RelativeVerticalPosition.
- When a shape must appear inside a table cell, set IsLayoutInCell = true after insertion.
- Validate only the existence of the saved document and that the inserted shape's ShapeType matches the expected value.
- Avoid inspecting unrelated node collections or document internals that are not part of the requested behavior.

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

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\shapes\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `shapes/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0

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
| 4 | `insert-an-ole-object-shape-and-lock-its-aspect-ratio-to-preserve-original-proportions.cs` | Insert an OLE object shape and lock its aspect ratio to preserve original proportions. | shape-workflow | docx | existing_repo |
| 5 | `insert-a-picture-shape-into-a-table-cell-and-enable-islayoutincell-for-proper-layout.cs` | Insert a picture shape into a table cell and enable IsLayoutInCell for proper layout. | shape-workflow | docx | mcp |
| 6 | `lock-aspect-ratio-for-shapes-by-setting-shape-aspectratiolocked-property-to-true.cs` | Lock aspect ratio for shapes by setting Shape.AspectRatioLocked property to true. | shape-workflow | docx | mcp |
| 7 | `set-shape-layout-inside-table-cells-by-configuring-islayoutincell-property-to-true.cs` | Set shape layout inside table cells by configuring IsLayoutInCell property to true. | shape-workflow | docx | mcp |
| 8 | `create-a-custom-snip-corner-rectangle-shape-with-defined-corner-radius-and-light-gray-fill.cs` | Create a custom snip corner rectangle shape with defined corner radius and light gray fill. | shape-workflow | docx | existing_repo |
| 9 | `retrieve-actual-bounds-of-a-shape-using-shape-getactualbounds-and-log-coordinate-points.cs` | Retrieve actual bounds of a shape using Shape.GetActualBounds and log coordinate points. | shape-workflow | docx | llm |
| 10 | `set-vertical-anchor-of-a-shape-to-top-middle-or-bottom-based-on-design-requirements.cs` | Set vertical anchor of a shape to top, middle, or bottom based on design requirements. | shape-workflow | docx | mcp |
| 11 | `detect-smartart-shapes-using-shape-issmartart-property-and-replace-them-with-alternative-d.cs` | Detect SmartArt shapes using Shape.IsSmartArt property and replace them with alternative diagrams. | shape-workflow | docx | llm |
| 12 | `insert-a-horizontal-rule-shape-with-custom-width-thickness-and-color-to-separate-sections.cs` | Insert a horizontal rule shape with custom width, thickness, and color to separate sections. | shape-workflow | docx | mcp |
| 13 | `import-mathml-content-as-shapes-position-them-inline-and-adjust-baseline-alignment.cs` | Import MathML content as shapes, position them inline, and adjust baseline alignment. | shape-workflow | docx | llm |
| 14 | `create-a-groupshape-instance-add-a-picture-textbox-and-autoshape-then-arrange-them.cs` | Create a GroupShape instance, add a picture, textbox, and AutoShape, then arrange them. | shape-workflow | docx | llm |
| 15 | `append-a-new-shape-to-an-existing-groupshape-and-update-group-bounds-accordingly.cs` | Append a new shape to an existing GroupShape and update group bounds accordingly. | shape-workflow | docx | mcp |
| 16 | `iterate-through-all-shapes-in-a-document-and-output-each-shape-s-type-using-shapetype-enum.cs` | Iterate through all shapes in a document and output each shape's type using ShapeType enumeration. | shape-workflow | doc | mcp |
| 17 | `change-a-picture-shape-to-an-autoshape-programmatically-by-setting-shapetype-while-preserv.cs` | Change a picture shape to an AutoShape programmatically by setting ShapeType while preserving size. | shape-workflow | docx | existing_repo |
| 18 | `apply-a-uniform-fill-color-to-every-shape-in-a-document-for-consistent-visual-branding.cs` | Apply a uniform fill color to every shape in a document for consistent visual branding. | shape-workflow | doc | mcp |
| 19 | `export-each-shape-s-image-to-a-separate-file-naming-files-based-on-shape-index-and-type.cs` | Export each shape's image to a separate file, naming files based on shape index and type. | shape-workflow | docx | mcp |
| 20 | `load-a-docx-template-insert-required-shapes-and-save-the-modified-document-as-a-new-docx-f.cs` | Load a DOCX template, insert required shapes, and save the modified document as a new DOCX file. | shape-workflow | docx, doc | mcp |
| 21 | `save-a-document-containing-shapes-as-pdf-while-preserving-exact-layout-and-visual-fidelity.cs` | Save a document containing shapes as PDF while preserving exact layout and visual fidelity. | shape-workflow | doc, pdf | mcp |
| 22 | `batch-process-a-folder-of-docx-files-inserting-a-semi-transparent-watermark-shape-into-eac.cs` | Batch process a folder of DOCX files, inserting a semi-transparent watermark shape into each document. | shape-workflow | docx, doc | existing_repo |
| 23 | `move-documentbuilder-cursor-to-a-bookmark-before-inserting-a-shape-at-the-bookmarked-locat.cs` | Move DocumentBuilder cursor to a bookmark before inserting a shape at the bookmarked location. | shape-workflow | doc | mcp |
| 24 | `insert-a-shape-with-relative-positioning-to-the-preceding-paragraph-to-maintain-flow-withi.cs` | Insert a shape with relative positioning to the preceding paragraph to maintain flow within the text. | shape-workflow | docx | mcp |
| 25 | `set-a-shape-s-rotation-angle-to-45-degrees-and-verify-visual-result-in-rendered-document.cs` | Set a shape's rotation angle to 45 degrees and verify visual result in rendered document. | shape-workflow | doc | mcp |
| 26 | `configure-text-wrapping-around-a-shape-on-both-sides-of-the-page-using-shape-wraptype-prop.cs` | Configure text wrapping around a shape on both sides of the page using Shape.WrapType property. | shape-workflow | docx | mcp |
| 27 | `insert-a-shape-inside-a-table-cell-then-adjust-the-cell-s-left-and-right-margins-for-prope.cs` | Insert a shape inside a table cell, then adjust the cell's left and right margins for proper spacing. | shape-workflow | docx | mcp |
| 28 | `clone-an-existing-shape-modify-its-fill-color-and-insert-the-clone-at-a-different-location.cs` | Clone an existing shape, modify its fill color, and insert the clone at a different location. | shape-workflow | docx | mcp |
| 29 | `convert-a-floating-shape-to-an-inline-shape-and-later-revert-it-back-to-floating-using-pro.cs` | Convert a floating shape to an inline shape and later revert it back to floating using properties. | shape-workflow | docx | mcp |
| 30 | `apply-a-hyperlink-to-a-shape-that-points-to-an-external-website-and-test-navigation-functi.cs` | Apply a hyperlink to a shape that points to an external website and test navigation functionality. | shape-workflow | docx | llm |
| 31 | `set-alternative-text-for-a-shape-to-improve-accessibility-for-screen-readers-and-assistive.cs` | Set alternative text for a shape to improve accessibility for screen readers and assistive technologies. | shape-workflow | docx | mcp |
| 32 | `store-custom-metadata-in-a-shape-s-title-property-for-later-retrieval-during-document-proc.cs` | Store custom metadata in a shape's Title property for later retrieval during document processing. | shape-workflow | doc | mcp |
| 33 | `retrieve-a-shape-s-z-order-index-and-bring-the-shape-to-the-front-of-the-document-layering.cs` | Retrieve a shape's Z-order index and bring the shape to the front of the document layering. | shape-workflow | doc | mcp |
| 34 | `send-a-shape-to-the-back-of-the-document-s-layering-order-to-ensure-underlying-content-rem.cs` | Send a shape to the back of the document's layering order to ensure underlying content remains visible. | shape-workflow | doc | mcp |
| 35 | `group-multiple-shapes-together-and-apply-a-collective-rotation-of-30-degrees-to-the-entire.cs` | Group multiple shapes together and apply a collective rotation of 30 degrees to the entire group. | shape-workflow | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Unsupported API invention

- Symptom: The compiler reports missing members or enumerations such as Shape.StrokeDashStyle or ShapeType.SnipCornerRectangle.
- Root cause: Code references APIs that were never implemented in the selected Aspose.Words version.
- Preferred fix: Replace the invented members with documented alternatives like Shape.StrokeColor, Shape.StrokeWeight, or use ShapeType.Rectangle for custom rectangles.
- Prevention: Consult the official Aspose.Words API reference for the exact version before using any member.

### Missing System.Drawing reference

- Symptom: Compilation fails with CS1069 indicating that Bitmap or Graphics cannot be found.
- Root cause: The project does not reference the System.Drawing.Common assembly required for image creation.
- Preferred fix: Add a reference to System.Drawing.Common via NuGet and use the provided types only for temporary image generation.
- Prevention: Include System.Drawing.Common in the project file whenever the sample creates bitmap files.

### Incorrect method overload

- Symptom: CS1501 errors show that InsertOleObject is called with an unsupported number or type of arguments.
- Root cause: The code supplies parameters that do not match any overload defined in Aspose.Words.
- Preferred fix: Call InsertOleObject with the signature (string fileName, bool isLinked, bool asIcon, string iconFileName) and omit unsupported arguments.
- Prevention: Verify overload signatures in the API documentation before passing arguments.

### Read-only property misuse

- Symptom: CS0200 errors indicate attempts to assign to ShapeBase.ShapeType or Shape.AspectRatioLocked.
- Root cause: The properties are read-only; shape type cannot be changed after creation, and aspect ratio is controlled via Shape.AspectRatioLocked (available) rather than a non-existent setter.
- Preferred fix: Create the shape with the desired ShapeType initially and set AspectRatioLocked = true after insertion.
- Prevention: Set shape characteristics only through mutable properties; do not attempt to reassign read-only members.

### Invalid node insertion

- Symptom: Runtime ArgumentException states that a node of this type cannot be inserted at the current location.
- Root cause: A shape or GroupShape is appended to a node that does not accept child nodes of that type (e.g., inserting a GroupShape directly into the document body).
- Preferred fix: Insert shapes using DocumentBuilder methods or append them to a Paragraph node, which is a valid container for Shape nodes.
- Prevention: Always use DocumentBuilder or Paragraph.AppendChild to add Shape objects.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Periodically review the Aspose.Words release notes to update the list of supported shape properties and enumerations.
