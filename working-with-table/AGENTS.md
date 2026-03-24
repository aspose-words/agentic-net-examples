---
name: working-with-table
description: C# examples for working-with-table using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - working-with-table

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **working-with-table** category.
This folder contains standalone C# examples for working-with-table operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **working-with-table**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (189/189 files) ← category-specific
- `using Aspose.Words.Tables;` (186/189 files)
- `using System;` (159/189 files)
- `using System.Drawing;` (43/189 files)
- `using System.IO;` (20/189 files)
- `using Aspose.Words.Drawing;` (18/189 files)
- `using Aspose.Words.Saving;` (10/189 files)
- `using System.Collections.Generic;` (6/189 files)
- `using Aspose.Words.Fields;` (5/189 files)
- `using System.Text.Json;` (3/189 files)
- `using System.Data;` (2/189 files)
- `using System.Xml;` (2/189 files)

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
| [add-comment-specific-table-cell-locating-cell-documentb...](./add-comment-specific-table-cell-locating-cell-documentbuilder-insertcomment.cs) | `Document`, `DocumentBuilder`, `Comment` | Add comment specific table cell locating cell documentbuilder insertcomment |
| [add-conditional-row-that-appears-only-when-data-value-e...](./add-conditional-row-that-appears-only-when-data-value-exceeds-threshold-if-fields.cs) | `Document`, `DocumentBuilder`, `DataTable` | Add conditional row that appears only when data value exceeds threshold if fi... |
| [add-double-border-around-entire-table-configuring-table...](./add-double-border-around-entire-table-configuring-table-border-each-side.cs) | `DocumentBuilder`, `BorderType`, `LineStyle` | Add double border around entire table configuring table border each side |
| [add-footer-row-totals-calculating-column-sums-inserting...](./add-footer-row-totals-calculating-column-sums-inserting-new-formatted-row.cs) | `RowFormat`, `Document`, `DocumentBuilder` | Add footer row totals calculating column sums inserting new formatted row |
| [add-new-row-existing-table-calling-table-rows-add-then-...](./add-new-row-existing-table-calling-table-rows-add-then-inserting-cells-row-cells-add.cs) | `Document`, `DocumentBuilder`, `Row` | Add new row existing table calling table rows add then inserting cells row ce... |
| [add-row-cells-containing-rich-text-formatting-inserting...](./add-row-cells-containing-rich-text-formatting-inserting-formatted-runs-each-cell-s.cs) | `Run`, `FirstParagraph`, `Font` | Add row cells containing rich text formatting inserting formatted runs each c... |
| [add-row-cells-varying-column-spans-merging-cells-horizo...](./add-row-cells-varying-column-spans-merging-cells-horizontally-across-multiple-columns.cs) | `CellFormat`, `CellMerge`, `Document` | Add row cells varying column spans merging cells horizontally across multiple... |
| [add-table-footer-row-that-calculates-totals-field-codes...](./add-table-footer-row-that-calculates-totals-field-codes-column-sums.cs) | `Document`, `DocumentBuilder`, `AutoFitBehavior` | Add table footer row that calculates totals field codes column sums |
| [add-toc-entry-table-inserting-bookmark-referencing-it-t...](./add-toc-entry-table-inserting-bookmark-referencing-it-toc-field.cs) | `Document`, `DocumentBuilder`, `BreakType` | Add toc entry table inserting bookmark referencing it toc field |
| [adjust-cell-padding-three-points-all-sides-setting-cell...](./adjust-cell-padding-three-points-all-sides-setting-cellformat-paddingtop-bottom-left.cs) | `Document`, `DocumentBuilder`, `CellFormat` | Adjust cell padding three points all sides setting cellformat paddingtop bott... |
| [adjust-row-spacing-configuring-row-rowformat-spaceafter...](./adjust-row-spacing-configuring-row-rowformat-spaceafter-spacebefore-properties-each-row.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Adjust row spacing configuring row rowformat spaceafter spacebefore propertie... |
| [align-table-center-page-setting-table-alignment-tableal...](./align-table-center-page-setting-table-alignment-tablealignment-center.cs) | `Document`, `DocumentBuilder`, `Input` | Align table center page setting table alignment tablealignment center |
| [apply-built-grid-table-5-dark-style-table-after-inserti...](./apply-built-grid-table-5-dark-style-table-after-inserting-its-rows.cs) | `TableStyleOptions`, `Document`, `DocumentBuilder` | Apply built grid table 5 dark style table after inserting its rows |
| [apply-conditional-cell-shading-based-numeric-values-ite...](./apply-conditional-cell-shading-based-numeric-values-iterating-cells-setting.cs) | `Document`, `DocumentBuilder`, `CellFormat` | Apply conditional cell shading based numeric values iterating cells setting |
| [apply-custom-border-color-first-column-cells-setting-ce...](./apply-custom-border-color-first-column-cells-setting-cellformat-border-color-property.cs) | `Document`, `DocumentBuilder`, `FirstSection` | Apply custom border color first column cells setting cellformat border color... |
| [apply-custom-border-color-top-edge-table-setting-table-...](./apply-custom-border-color-top-edge-table-setting-table-borders-top-color-property.cs) | `Document`, `DocumentBuilder`, `BorderType` | Apply custom border color top edge table setting table borders top color prop... |
| [apply-custom-border-thickness-two-points-all-sides-tabl...](./apply-custom-border-thickness-two-points-all-sides-table-setting-table-borders.cs) | `Document`, `DocumentBuilder`, `Output` | Apply custom border thickness two points all sides table setting table borders |
| [apply-custom-cell-border-style-different-line-widths-ea...](./apply-custom-cell-border-style-different-line-widths-each-side-setting-cellformat.cs) | `CellFormat`, `BorderType`, `LineStyle` | Apply custom cell border style different line widths each side setting cellfo... |
| [apply-custom-cell-margins-setting-cell-cellformat-margi...](./apply-custom-cell-margins-setting-cell-cellformat-margintop-marginbottom-marginleft.cs) | `CellFormat`, `Document`, `DocumentBuilder` | Apply custom cell margins setting cell cellformat margintop marginbottom marg... |
| [apply-custom-margin-around-entire-table-setting-table-l...](./apply-custom-margin-around-entire-table-setting-table-leftindent-table-rightindent.cs) | `Document`, `DocumentBuilder`, `PreferredWidth` | Apply custom margin around entire table setting table leftindent table righti... |
| [apply-custom-outer-border-style-while-disabling-inner-b...](./apply-custom-outer-border-style-while-disabling-inner-borders-configuring-table.cs) | `BorderType`, `LineStyle`, `Color` | Apply custom outer border style while disabling inner borders configuring table |
| [apply-custom-table-style-that-includes-alternating-row-...](./apply-custom-table-style-that-includes-alternating-row-shading-setting-tablestyle.cs) | `Document`, `DocumentBuilder`, `ConditionalStyleType` | Apply custom table style that includes alternating row shading setting tables... |
| [apply-custom-table-style-that-includes-double-line-bord...](./apply-custom-table-style-that-includes-double-line-border-around-table-single-lines.cs) | `BorderType`, `LineStyle`, `Color` | Apply custom table style that includes double line border around table single... |
| [apply-diagonal-line-across-cell-setting-cellformat-diag...](./apply-diagonal-line-across-cell-setting-cellformat-diagonaldown-diagonalup-properties.cs) | `CellFormat`, `BorderType`, `DocumentBuilder` | Apply diagonal line across cell setting cellformat diagonaldown diagonalup pr... |
| [apply-different-cell-background-colors-based-column-ind...](./apply-different-cell-background-colors-based-column-index-iterating-cells-setting.cs) | `Document`, `DocumentBuilder`, `CellFormat` | Apply different cell background colors based column index iterating cells set... |
| [apply-distinct-border-styles-first-row-last-row-inner-c...](./apply-distinct-border-styles-first-row-last-row-inner-cells-via-rowformat-cellformat.cs) | `Borders`, `RowFormat`, `LineStyle` | Apply distinct border styles first row last row inner cells via rowformat cel... |
| [apply-gradient-shading-cells-setting-cellformat-shading...](./apply-gradient-shading-cells-setting-cellformat-shading-texture.cs) | `CellFormat`, `Shading`, `Document` | Apply gradient shading cells setting cellformat shading texture |
| [apply-light-gray-shading-header-row-cells-via-cell-cell...](./apply-light-gray-shading-header-row-cells-via-cell-cellformat-shading.cs) | `CellFormat`, `Shading`, `Document` | Apply light gray shading header row cells via cell cellformat shading |
| [apply-predefined-built-tablestyle-setting-table-style-t...](./apply-predefined-built-tablestyle-setting-table-style-tablegrid.cs) | `Document`, `DocumentBuilder`, `StyleIdentifier` | Apply predefined built tablestyle setting table style tablegrid |
| [apply-table-background-shading-entire-table-setting-tab...](./apply-table-background-shading-entire-table-setting-table-shading.cs) | `Document`, `DocumentBuilder`, `Color` | Apply table background shading entire table setting table shading |
| ... | | *and 159 more files* |

## Category Statistics
- Total examples: 189

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for working-with-table patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
