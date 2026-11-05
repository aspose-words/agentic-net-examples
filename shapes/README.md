# Shapes Examples for Aspose.Words for .NET

The Shapes category demonstrates how to work with Aspose.Words drawing objects such as images, AutoShapes, OLE objects, and GroupShapes. Verified examples illustrate correct image insertion, OLE embedding with aspect-ratio locking, table-cell layout, custom rectangle creation, shape bounds retrieval, vertical positioning, SmartArt replacement, and horizontal rule insertion.

## Snapshot

- Category: Shapes
- Slug: shapes
- Total examples: 35
- Publish-ready successful examples: 35 / 35
- Source run: 20261102_031946_08fb5b
- Shape Workflow examples: 35

## Category rules that shaped these examples

- Never reference System.Drawing types such as Bitmap, Graphics, or Color unless the System.Drawing.Common package is explicitly added and supported on the target platform.
- Do not use Shape members that are not present in Aspose.Words 26.10.0, such as LockAspectRatio, GetActualBounds, CornerRadius, UpdateGroupShapeBounds, BringToFront, SendToBack, Hyperlink, or TextBoxContent.
- Insert images by providing a file path or a byte array and then adjust WrapType, RelativeHorizontalPosition, RelativeVerticalPosition, Left, and Top as needed.
- Embed OLE objects using the InsertOleObject overload that accepts a Stream, a progId string, a Boolean for icon display, and an optional icon Stream.
- Create tables with the Table, Row, and Cell classes; populate cells with paragraphs before inserting shapes, and set IsLayoutInCell on the shape to control layout.
- Validate that every temporary file or directory exists before it is read, and confirm that the output document file is created after saving.
- Check that a shape was successfully inserted before accessing its properties, throwing a clear exception if the shape reference is null.

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
dotnet add package Aspose.Words --version 26.10.0

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
| 4 | `insert-an-ole-object-shape-and-lock-its-aspect-ratio-to-preserve-original-proportions.cs` | Insert an OLE object shape and lock its aspect ratio to preserve original proportions. | shape-workflow | docx | llm |
| 5 | `insert-a-picture-shape-into-a-table-cell-and-enable-islayoutincell-for-proper-layout.cs` | Insert a picture shape into a table cell and enable IsLayoutInCell for proper layout. | shape-workflow | docx | mcp |
| 6 | `lock-aspect-ratio-for-shapes-by-setting-shape-aspectratiolocked-property-to-true.cs` | Lock aspect ratio for shapes by setting Shape.AspectRatioLocked property to true. | shape-workflow | docx | mcp |
| 7 | `set-shape-layout-inside-table-cells-by-configuring-islayoutincell-property-to-true.cs` | Set shape layout inside table cells by configuring IsLayoutInCell property to true. | shape-workflow | docx | mcp |
| 8 | `create-a-custom-snip-corner-rectangle-shape-with-defined-corner-radius-and-light-gray-fill.cs` | Create a custom snip corner rectangle shape with defined corner radius and light gray fill. | shape-workflow | docx | mcp |
| 9 | `retrieve-actual-bounds-of-a-shape-using-shape-getactualbounds-and-log-coordinate-points.cs` | Retrieve actual bounds of a shape using Shape.GetActualBounds and log coordinate points. | shape-workflow | docx | llm |
| 10 | `set-vertical-anchor-of-a-shape-to-top-middle-or-bottom-based-on-design-requirements.cs` | Set vertical anchor of a shape to top, middle, or bottom based on design requirements. | shape-workflow | docx | mcp |
| 11 | `detect-smartart-shapes-using-shape-issmartart-property-and-replace-them-with-alternative-d.cs` | Detect SmartArt shapes using Shape.IsSmartArt property and replace them with alternative diagrams. | shape-workflow | docx | mcp |
| 12 | `insert-a-horizontal-rule-shape-with-custom-width-thickness-and-color-to-separate-sections.cs` | Insert a horizontal rule shape with custom width, thickness, and color to separate sections. | shape-workflow | docx | mcp |
| 13 | `import-mathml-content-as-shapes-position-them-inline-and-adjust-baseline-alignment.cs` | Import MathML content as shapes, position them inline, and adjust baseline alignment. | shape-workflow | docx | mcp |
| 14 | `create-a-groupshape-instance-add-a-picture-textbox-and-autoshape-then-arrange-them.cs` | Create a GroupShape instance, add a picture, textbox, and AutoShape, then arrange them. | shape-workflow | docx | llm |
| 15 | `append-a-new-shape-to-an-existing-groupshape-and-update-group-bounds-accordingly.cs` | Append a new shape to an existing GroupShape and update group bounds accordingly. | shape-workflow | docx | mcp |
| 16 | `iterate-through-all-shapes-in-a-document-and-output-each-shape-s-type-using-shapetype-enum.cs` | Iterate through all shapes in a document and output each shape's type using ShapeType enumeration. | shape-workflow | doc | mcp |
| 17 | `change-a-picture-shape-to-an-autoshape-programmatically-by-setting-shapetype-while-preserv.cs` | Change a picture shape to an AutoShape programmatically by setting ShapeType while preserving size. | shape-workflow | docx | llm |
| 18 | `apply-a-uniform-fill-color-to-every-shape-in-a-document-for-consistent-visual-branding.cs` | Apply a uniform fill color to every shape in a document for consistent visual branding. | shape-workflow | doc | mcp |
| 19 | `export-each-shape-s-image-to-a-separate-file-naming-files-based-on-shape-index-and-type.cs` | Export each shape's image to a separate file, naming files based on shape index and type. | shape-workflow | docx | mcp |
| 20 | `load-a-docx-template-insert-required-shapes-and-save-the-modified-document-as-a-new-docx-f.cs` | Load a DOCX template, insert required shapes, and save the modified document as a new DOCX file. | shape-workflow | docx, doc | mcp |
| 21 | `save-a-document-containing-shapes-as-pdf-while-preserving-exact-layout-and-visual-fidelity.cs` | Save a document containing shapes as PDF while preserving exact layout and visual fidelity. | shape-workflow | doc, pdf | mcp |
| 22 | `batch-process-a-folder-of-docx-files-inserting-a-semi-transparent-watermark-shape-into-eac.cs` | Batch process a folder of DOCX files, inserting a semi-transparent watermark shape into each document. | shape-workflow | docx, doc | llm |
| 23 | `move-documentbuilder-cursor-to-a-bookmark-before-inserting-a-shape-at-the-bookmarked-locat.cs` | Move DocumentBuilder cursor to a bookmark before inserting a shape at the bookmarked location. | shape-workflow | doc | mcp |
| 24 | `insert-a-shape-with-relative-positioning-to-the-preceding-paragraph-to-maintain-flow-withi.cs` | Insert a shape with relative positioning to the preceding paragraph to maintain flow within the text. | shape-workflow | docx | mcp |
| 25 | `set-a-shape-s-rotation-angle-to-45-degrees-and-verify-visual-result-in-rendered-document.cs` | Set a shape's rotation angle to 45 degrees and verify visual result in rendered document. | shape-workflow | doc | mcp |
| 26 | `configure-text-wrapping-around-a-shape-on-both-sides-of-the-page-using-shape-wraptype-prop.cs` | Configure text wrapping around a shape on both sides of the page using Shape.WrapType property. | shape-workflow | docx | mcp |
| 27 | `insert-a-shape-inside-a-table-cell-then-adjust-the-cell-s-left-and-right-margins-for-prope.cs` | Insert a shape inside a table cell, then adjust the cell's left and right margins for proper spacing. | shape-workflow | docx | mcp |
| 28 | `clone-an-existing-shape-modify-its-fill-color-and-insert-the-clone-at-a-different-location.cs` | Clone an existing shape, modify its fill color, and insert the clone at a different location. | shape-workflow | docx | mcp |
| 29 | `convert-a-floating-shape-to-an-inline-shape-and-later-revert-it-back-to-floating-using-pro.cs` | Convert a floating shape to an inline shape and later revert it back to floating using properties. | shape-workflow | docx | mcp |
| 30 | `apply-a-hyperlink-to-a-shape-that-points-to-an-external-website-and-test-navigation-functi.cs` | Apply a hyperlink to a shape that points to an external website and test navigation functionality. | shape-workflow | docx | existing_repo |
| 31 | `set-alternative-text-for-a-shape-to-improve-accessibility-for-screen-readers-and-assistive.cs` | Set alternative text for a shape to improve accessibility for screen readers and assistive technologies. | shape-workflow | docx | mcp |
| 32 | `store-custom-metadata-in-a-shape-s-title-property-for-later-retrieval-during-document-proc.cs` | Store custom metadata in a shape's Title property for later retrieval during document processing. | shape-workflow | doc | mcp |
| 33 | `retrieve-a-shape-s-z-order-index-and-bring-the-shape-to-the-front-of-the-document-layering.cs` | Retrieve a shape's Z-order index and bring the shape to the front of the document layering. | shape-workflow | doc | mcp |
| 34 | `send-a-shape-to-the-back-of-the-document-s-layering-order-to-ensure-underlying-content-rem.cs` | Send a shape to the back of the document's layering order to ensure underlying content remains visible. | shape-workflow | doc | mcp |
| 35 | `group-multiple-shapes-together-and-apply-a-collective-rotation-of-30-degrees-to-the-entire.cs` | Group multiple shapes together and apply a collective rotation of 30 degrees to the entire group. | shape-workflow | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Missing System.Drawing reference

- Symptom: Compilation fails with CS1069 errors for Bitmap, Graphics, or Color types.
- Root cause: The code imports System.Drawing types that are not available without the System.Drawing.Common package, which is unsupported on non-Windows runtimes.
- Preferred fix: Replace System.Drawing usage with raw image byte arrays or use Aspose.Words image insertion APIs that accept a file path or stream.
- Prevention: Never add System.Drawing namespaces in examples; generate sample images as byte arrays or embed pre-created PNG data.

### Use of non-existent Shape members

- Symptom: Compilation fails with CS1061 errors for members such as LockAspectRatio, GetActualBounds, CornerRadius, UpdateGroupShapeBounds, BringToFront, SendToBack, Hyperlink, or TextBoxContent.
- Root cause: The code references APIs that are not part of Aspose.Words 26.10.0.
- Preferred fix: Replace invented members with documented alternatives (e.g., AspectRatioLocked, Bounds, FillColor, StrokeColor, HyperlinkInfo via Shape.HyperlinkInfo property if needed).
- Prevention: Consult the official Aspose.Words API reference for the exact version before using any Shape property or method.

### Invalid DocumentBuilder table methods

- Symptom: Compilation fails with CS1061 errors for EndCell, EndRow, or similar methods.
- Root cause: Older examples use DocumentBuilder.EndCell/EndRow which were removed in recent versions.
- Preferred fix: Construct tables using the Table, Row, and Cell classes and append them to the document body; use builder.MoveTo(cell.FirstParagraph) to position the cursor.
- Prevention: Follow the current table-building pattern documented for Aspose.Words 26.x.

### Incorrect InsertOleObject overload

- Symptom: Compilation fails with CS1501 or CS1503 errors because the arguments do not match any overload.
- Root cause: The code supplies a file path string or mismatched parameter types to InsertOleObject.
- Preferred fix: Use the overload that takes (Stream oleStream, string progId, bool isIconic, Stream? iconStream) and pass a FileStream for the OLE content.
- Prevention: Check the method signature in the API docs and always provide a Stream for the OLE data.

### Unsupported positioning enums

- Symptom: Compilation fails with CS0246 or CS0234 errors for VerticalAnchor, RelativeVerticalPosition, etc.
- Root cause: The code references enums that are not defined in Aspose.Words.Drawing for the current version.
- Preferred fix: Use RelativeHorizontalPosition and RelativeVerticalPosition enums together with Left and Top properties to control shape placement.
- Prevention: Reference only the positioning enums listed in the official documentation.

### Invalid node insertion location

- Symptom: Runtime throws System.ArgumentException: Cannot insert a node of this type at this location.
- Root cause: The builder attempts to insert a shape when the cursor is not positioned inside a valid paragraph or cell.
- Preferred fix: Move the DocumentBuilder to a paragraph within the target cell (builder.MoveTo(cell.FirstParagraph)) before calling InsertShape or InsertImage.
- Prevention: Always ensure the insertion point is a CompositeNode that accepts Shape children.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Periodically review Aspose.Words release notes to update the list of supported Shape members and adjust examples accordingly.
- Run the verification suite after each library upgrade to catch newly introduced API changes.
