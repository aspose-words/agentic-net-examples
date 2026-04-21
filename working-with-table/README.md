# Working with Table Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **Working with Table** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Working with Table**
- Slug: **working-with-table**
- Total task entries: **198**
- Workflow examples: **198 / 198** use the standard table workflow
- Duplicate published filename mappings: **1**

## Category rules that shaped these examples

- Use native Aspose.Words APIs directly.
- Create local sample documents when a task refers to an existing file, stream, template, or input asset.
- Do not assume external files or folders already exist.
- Prefer `DocumentBuilder.StartTable`, `InsertCell`, `EndRow`, and `EndTable` for builder-based creation.
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
Copy-Item ..\working-with-table\<example-file>.cs .\Program.cs

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
Copy-Item ..\working-with-table\create-a-table-with-three-rows-and-four-columns-using-documentbuilder-in-a-new-document.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `create-a-table-with-three-rows-and-four-columns-using-documentbuilder-in-a-new-document.cs` | Create a table with three rows and four columns using DocumentBuilder in a new document. | table-workflow | doc | mcp |
| 2 | `create-a-table-using-the-dom-api-by-instantiating-table-row-and-cell-nodes-and-appending-t.cs` | Create a table using the DOM API by instantiating Table, Row, and Cell nodes and appending them. | table-workflow | docx | mcp |
| 3 | `add-a-new-row-to-an-existing-table-by-calling-table-rows-add-and-then-inserting-cells-with.cs` | Add a new row to an existing table by calling Table.Rows.Add and then inserting cells with Row.Cells.Add. | table-workflow | docx | mcp |
| 4 | `insert-a-table-at-a-bookmark-location-by-moving-the-documentbuilder-cursor-to-the-bookmark.cs` | Insert a table at a bookmark location by moving the DocumentBuilder cursor to the bookmark and calling InsertTable. | table-workflow | doc | mcp |
| 5 | `insert-a-table-into-a-header-by-moving-documentbuilder-to-the-header-and-inserting-the-tab.cs` | Insert a table into a header by moving DocumentBuilder to the header and inserting the table. | table-workflow | doc | mcp |
| 6 | `insert-a-table-into-a-header-footer-by-moving-documentbuilder-to-the-respective-section-an.cs` | Insert a table into a header/footer by moving DocumentBuilder to the respective section and calling InsertTable. | table-workflow | doc | mcp |
| 7 | `insert-a-table-into-a-document-end-by-moving-documentbuilder-to-the-end-and-calling-insert.cs` | Insert a table into a document end by moving DocumentBuilder to the end and calling InsertTable. | table-workflow | doc | mcp |
| 8 | `insert-a-table-after-a-specific-paragraph-by-locating-the-paragraph-node-and-using-inserta.cs` | Insert a table after a specific paragraph by locating the paragraph node and using InsertAfter. | table-workflow | docx | mcp |
| 9 | `insert-a-table-after-a-heading-paragraph-by-locating-the-heading-node-and-using-insertafte.cs` | Insert a table after a heading paragraph by locating the heading node and using InsertAfter with the table. | table-workflow | docx | mcp |
| 10 | `insert-a-table-after-each-heading-paragraph-by-iterating-headings-and-using-insertafter-fo.cs` | Insert a table after each heading paragraph by iterating headings and using InsertAfter for each table. | table-workflow | docx | mcp |
| 11 | `insert-a-table-after-a-paragraph-containing-a-specific-keyword-by-searching-nodes-and-usin.cs` | Insert a table after a paragraph containing a specific keyword by searching nodes and using InsertAfter. | table-workflow | docx | mcp |
| 12 | `insert-an-empty-paragraph-after-a-table-to-prevent-automatic-merging-with-the-following-ta.cs` | Insert an empty paragraph after a table to prevent automatic merging with the following table. | table-workflow | docx | mcp |
| 13 | `clone-an-existing-table-node-and-insert-the-copy-after-a-specific-paragraph-using-node-ins.cs` | Clone an existing Table node and insert the copy after a specific paragraph using Node.InsertAfter. | table-workflow | docx | mcp |
| 14 | `clone-a-table-from-a-template-document-and-replace-placeholder-text-in-each-cell-using-fin.cs` | Clone a table from a template document and replace placeholder text in each cell using FindReplaceOptions. | table-workflow | doc | mcp |
| 15 | `load-an-html-file-with-a-complex-table-and-convert-it-to-a-word-document-preserving-merged.cs` | Load an HTML file with a complex table and convert it to a Word document preserving merged cells. | table-workflow | doc, html | mcp |
| 16 | `insert-a-table-from-an-html-string-using-documentbuilder-inserthtml-while-preserving-cell.cs` | Insert a table from an HTML string using DocumentBuilder.InsertHtml while preserving cell formatting. | table-workflow | doc, html | mcp |
| 17 | `set-table-autofit-to-autofitbehavior-autofittowindow-to-automatically-adjust-column-widths.cs` | Set Table.AutoFit to AutoFitBehavior.AutoFitToWindow to automatically adjust column widths to page margins. | table-workflow | docx | mcp |
| 18 | `enable-automatic-resizing-of-table-columns-to-fit-content-by-setting-table-autofit-to-auto.cs` | Enable automatic resizing of table columns to fit content by setting Table.AutoFit to AutoFitBehavior.AutoFitToContents. | table-workflow | docx | mcp |
| 19 | `prevent-automatic-resizing-of-a-table-after-content-changes-by-disabling-table-autofit-and.cs` | Prevent automatic resizing of a table after content changes by disabling Table.AutoFit and fixing column widths. | table-workflow | docx | mcp |
| 20 | `set-a-table-s-left-indent-to-two-centimeters-using-the-table-leftindent-property-for-preci.cs` | Set a table's left indent to two centimeters using the Table.LeftIndent property for precise positioning. | table-workflow | docx | mcp |
| 21 | `set-the-table-s-left-and-right-margins-to-one-centimeter-each-by-configuring-table-leftind.cs` | Set the table's left and right margins to one centimeter each by configuring Table.LeftIndent and Table.RightIndent. | table-workflow | docx | mcp |
| 22 | `define-a-table-s-preferred-width-as-100-percent-of-the-page-width-by-assigning-preferredwi.cs` | Define a table's preferred width as 100 percent of the page width by assigning PreferredWidth type and value. | table-workflow | docx | mcp |
| 23 | `create-a-table-with-a-fixed-width-of-fifteen-centimeters-by-setting-preferredwidth-type-an.cs` | Create a table with a fixed width of fifteen centimeters by setting PreferredWidth type and value accordingly. | table-workflow | docx | mcp |
| 24 | `create-a-table-with-a-fixed-width-and-centered-alignment-by-setting-preferredwidth-and-tab.cs` | Create a table with a fixed width and centered alignment by setting PreferredWidth and Table.Alignment properties. | table-workflow | docx | mcp |
| 25 | `set-the-table-s-vertical-alignment-within-the-page-by-adjusting-table-verticalalignment-pr.cs` | Set the table's vertical alignment within the page by adjusting Table.VerticalAlignment property to Middle. | table-workflow | docx | mcp |
| 26 | `align-a-table-to-the-center-of-the-page-by-setting-table-alignment-to-tablealignment-cente.cs` | Align a table to the center of the page by setting Table.Alignment to TableAlignment.Center. | table-workflow | docx | mcp |
| 27 | `apply-the-built-in-grid-table-5-dark-style-to-a-table-after-inserting-its-rows.cs` | Apply the built‑in "Grid Table 5 Dark" style to a table after inserting its rows. | table-workflow | docx | mcp |
| 28 | `iterate-all-tables-in-a-document-and-set-each-style-to-table-grid-using-a-foreach-loop.cs` | Iterate all tables in a document and set each style to "Table Grid" using a foreach loop. | table-workflow | doc | mcp |
| 29 | `apply-a-table-style-that-includes-header-row-formatting-by-enabling-table-styleoptions-fir.cs` | Apply a table style that includes header row formatting by enabling Table.StyleOptions.FirstRow. | table-workflow | docx | mcp |
| 30 | `apply-a-table-style-that-includes-bold-text-for-the-first-row-by-enabling-table-styleoptio.cs` | Apply a table style that includes bold text for the first row by enabling Table.StyleOptions.FirstRowBold. | table-workflow | docx | mcp |
| 31 | `apply-a-table-style-that-includes-bold-formatting-for-the-first-column-by-enabling-table-s.cs` | Apply a table style that includes bold formatting for the first column by enabling Table.StyleOptions.FirstColumnBold. | table-workflow | docx | mcp |
| 32 | `apply-a-table-style-that-includes-alternating-row-shading-by-enabling-table-styleoptions-r.cs` | Apply a table style that includes alternating row shading by enabling Table.StyleOptions.RowBanding. | table-workflow | docx | mcp |
| 33 | `apply-a-table-style-that-adds-a-thick-top-border-and-thin-bottom-border-by-configuring-tab.cs` | Apply a table style that adds a thick top border and thin bottom border by configuring Table.Borders. | table-workflow | docx | mcp |
| 34 | `apply-a-table-style-that-hides-borders-by-setting-table-styleoptions-to-tablestyleoptions.cs` | Apply a table style that hides borders by setting Table.StyleOptions to TableStyleOptions.NoBorders. | table-workflow | docx | mcp |
| 35 | `apply-a-table-style-that-disables-all-cell-shading-by-setting-table-styleoptions-to-tables.cs` | Apply a table style that disables all cell shading by setting Table.StyleOptions to TableStyleOptions.NoShading. | table-workflow | docx | mcp |
| 36 | `apply-a-table-style-that-disables-row-banding-by-setting-table-styleoptions-to-tablestyleo.cs` | Apply a table style that disables row banding by setting Table.StyleOptions to TableStyleOptions.NoRowBanding. | table-workflow | docx | mcp |
| 37 | `apply-a-table-style-that-enforces-fixed-column-widths-by-disabling-autofit-and-defining-pr.cs` | Apply a table style that enforces fixed column widths by disabling AutoFit and defining PreferredWidth for each column. | table-workflow | docx | mcp |
| 38 | `apply-a-custom-border-thickness-of-two-points-to-all-sides-of-a-table-by-setting-table-bor.cs` | Apply a custom border thickness of two points to all sides of a table by setting Table.Borders.LineWidth. | table-workflow | docx | mcp |
| 39 | `apply-a-custom-border-color-to-the-top-edge-of-a-table-by-setting-table-borders-top-color.cs` | Apply a custom border color to the top edge of a table by setting Table.Borders.Top.Color property. | table-workflow | docx | mcp |
| 40 | `remove-all-borders-from-a-table-by-setting-each-table-borders-linestyle-to-linestyle-none.cs` | Remove all borders from a table by setting each Table.Borders.LineStyle to LineStyle.None. | table-workflow | docx | mcp |
| 41 | `set-cell-background-shading-to-light-gray-by-assigning-a-shading-pattern-and-color-to-cell.cs` | Set cell background shading to light gray by assigning a shading pattern and color to CellFormat.Shading. | table-workflow | docx | mcp |
| 42 | `set-the-table-s-background-color-to-light-blue-by-assigning-a-shading-pattern-and-color-to.cs` | Set the table's background color to light blue by assigning a shading pattern and color to Table.Shading. | table-workflow | docx | mcp |
| 43 | `set-cell-text-alignment-to-center-horizontally-and-vertically-by-configuring-cellformat-ho.cs` | Set cell text alignment to center horizontally and vertically by configuring CellFormat.HorizontalAlignment and VerticalAlignment. | table-workflow | docx | mcp |
| 44 | `set-cell-vertical-alignment-to-bottom-by-configuring-cellformat-verticalalignment-property.cs` | Set cell vertical alignment to bottom by configuring CellFormat.VerticalAlignment property for each cell. | table-workflow | docx | mcp |
| 45 | `set-cell-text-direction-to-right-to-left-for-arabic-content-by-configuring-cellformat-text.cs` | Set cell text direction to right‑to‑left for Arabic content by configuring CellFormat.TextDirection property. | table-workflow | docx | mcp |
| 46 | `set-the-table-s-text-direction-to-right-to-left-for-languages-requiring-bidirectional-layo.cs` | Set the table's text direction to right‑to‑left for languages requiring bidirectional layout. | table-workflow | docx | mcp |
| 47 | `set-the-table-s-text-direction-to-top-to-bottom-for-vertical-asian-scripts-by-configuring.cs` | Set the table's text direction to top‑to‑bottom for vertical Asian scripts by configuring Table.TextDirection property. | table-workflow | docx | mcp |
| 48 | `set-individual-cell-margins-to-five-points-on-all-sides-by-adjusting-cellformat-margintop.cs` | Set individual cell margins to five points on all sides by adjusting CellFormat.MarginTop, Bottom, Left, and Right. | table-workflow | docx | mcp |
| 49 | `adjust-cell-padding-to-three-points-on-all-sides-by-setting-cellformat-paddingtop-bottom-l.cs` | Adjust cell padding to three points on all sides by setting CellFormat.PaddingTop, Bottom, Left, and Right. | table-workflow | docx | mcp |
| 50 | `set-the-table-s-default-cell-margins-to-two-points-by-configuring-table-defaultcellmargin.cs` | Set the table's default cell margins to two points by configuring Table.DefaultCellMargin globally. | table-workflow | docx | mcp |
| 51 | `set-row-height-to-exactly-twenty-points-by-assigning-row-height-and-row-heightrule-propert.cs` | Set row height to exactly twenty points by assigning Row.Height and Row.HeightRule properties appropriately. | table-workflow | docx | mcp |
| 52 | `set-the-table-s-row-height-to-auto-by-assigning-row-heightrule-to-heightrule-auto-without.cs` | Set the table's row height to auto by assigning Row.HeightRule to HeightRule.Auto without specifying Height. | table-workflow | docx | mcp |
| 53 | `set-the-table-s-row-spacing-to-ten-points-by-adjusting-row-height-and-row-heightrule-to-ex.cs` | Set the table's row spacing to ten points by adjusting Row.Height and Row.HeightRule to Exactly. | table-workflow | docx | mcp |
| 54 | `set-the-table-s-row-spacing-to-double-line-spacing-by-adjusting-row-heightrule-to-auto-and.cs` | Set the table's row spacing to double line spacing by adjusting Row.HeightRule to Auto and setting Height accordingly. | table-workflow | docx | mcp |
| 55 | `set-the-table-s-bottom-margin-to-five-points-by-configuring-table-bottommargin-property-fo.cs` | Set the table's bottom margin to five points by configuring Table.BottomMargin property for precise layout control. | table-workflow | docx | mcp |
| 56 | `set-the-table-s-top-margin-to-three-points-by-configuring-table-topmargin-property-for-pre.cs` | Set the table's top margin to three points by configuring Table.TopMargin property for precise vertical placement. | table-workflow | docx | mcp |
| 57 | `set-the-table-s-row-height-rule-to-atleast-and-specify-a-minimum-height-of-fifteen-points.cs` | Set the table's row height rule to AtLeast and specify a minimum height of fifteen points. | table-workflow | docx | mcp |
| 58 | `disable-row-breaking-across-pages-by-setting-row-allowbreakacrosspages-to-false-for-each-r.cs` | Disable row breaking across pages by setting Row.AllowBreakAcrossPages to false for each row. | table-workflow | docx | mcp |
| 59 | `configure-a-table-to-avoid-breaking-across-pages-by-setting-table-allowbreakacrosspages-to.cs` | Configure a table to avoid breaking across pages by setting Table.AllowBreakAcrossPages to false. | table-workflow | docx | mcp |
| 60 | `prevent-a-table-from-being-split-across-pages-by-setting-the-first-row-s-allowbreakacrossp.cs` | Prevent a table from being split across pages by setting the first row's AllowBreakAcrossPages to false. | table-workflow | docx | mcp |
| 61 | `prevent-a-table-from-being-split-across-columns-in-a-multi-column-layout-by-setting-allowb.cs` | Prevent a table from being split across columns in a multi‑column layout by setting AllowBreakAcrossColumns to false. | table-workflow | docx | mcp |
| 62 | `merge-two-adjacent-cells-horizontally-by-setting-the-first-cell-s-mergehorizontal-property.cs` | Merge two adjacent cells horizontally by setting the first cell's MergeHorizontal property to true. | table-workflow | docx | mcp |
| 63 | `merge-a-group-of-cells-vertically-by-setting-the-top-cell-s-mergevertical-property-to-true.cs` | Merge a group of cells vertically by setting the top cell's MergeVertical property to true. | table-workflow | docx | mcp |
| 64 | `merge-cells-vertically-across-three-rows-by-setting-the-top-cell-s-mergevertical-property.cs` | Merge cells vertically across three rows by setting the top cell's MergeVertical property to true. | table-workflow | docx | mcp |
| 65 | `add-a-row-of-cells-with-varying-column-spans-by-merging-cells-horizontally-across-multiple.cs` | Add a row of cells with varying column spans by merging cells horizontally across multiple columns. | table-workflow | docx | mcp |
| 66 | `insert-a-table-with-merged-header-cells-spanning-three-columns-by-setting-mergehorizontal.cs` | Insert a table with merged header cells spanning three columns by setting MergeHorizontal on the first row's cells. | table-workflow | docx | mcp |
| 67 | `create-a-table-with-one-cell-spanning-full-width-by-merging-all-first-row-cells-horizontal.cs` | Create a table with one cell spanning full width by merging all first‑row cells horizontally. | table-workflow | docx | mcp |
| 68 | `create-a-table-with-a-merged-cell-spanning-the-entire-first-row-by-merging-all-first-row-c.cs` | Create a table with a merged cell spanning the entire first row by merging all first‑row cells horizontally. | table-workflow | docx | mcp |
| 69 | `insert-a-table-with-a-cell-spanning-two-rows-and-three-columns-by-setting-mergehorizontal.cs` | Insert a table with a cell spanning two rows and three columns by setting MergeHorizontal and MergeVertical. | table-workflow | docx | mcp |
| 70 | `split-a-merged-cell-back-into-individual-cells-by-resetting-its-mergehorizontal-and-mergev.cs` | Split a merged cell back into individual cells by resetting its MergeHorizontal and MergeVertical properties to false. | table-workflow | docx | mcp |
| 71 | `split-a-table-into-two-separate-tables-at-the-third-row-using-the-table-split-method.cs` | Split a table into two separate tables at the third row using the Table.Split method. | table-workflow | docx | mcp |
| 72 | `split-a-table-into-three-parts-by-calling-table-split-twice-at-specified-row-indices-for-c.cs` | Split a table into three parts by calling Table.Split twice at specified row indices for custom pagination. | table-workflow | docx | mcp |
| 73 | `join-adjacent-tables-by-inserting-an-empty-paragraph-between-them.cs` | Join adjacent tables by inserting an empty paragraph between them. | table-workflow | docx | mcp |
| 74 | `set-the-table-s-text-wrapping-mode-to-around-surrounding-text-by-configuring-table-wraparo.cs` | Set the table's text wrapping mode to around surrounding text by configuring Table.WrapAround property. | table-workflow | docx | mcp |
| 75 | `add-a-row-with-cells-containing-rich-text-formatting-by-inserting-formatted-runs-into-each.cs` | Add a row with cells containing rich text formatting by inserting formatted runs into each cell's paragraphs. | table-workflow | docx | mcp |
| 76 | `create-a-table-with-a-header-row-that-repeats-on-each-page-by-setting-the-first-row-s-head.cs` | Create a table with a header row that repeats on each page by setting the first row's HeadingFormat to true. | table-workflow | docx | mcp |
| 77 | `create-a-table-with-alternating-header-and-data-row-colors-by-applying-shading-based-on-ro.cs` | Create a table with alternating header and data row colors by applying shading based on row index parity. | table-workflow | docx | mcp |
| 78 | `create-a-table-with-alternating-row-colors-by-applying-a-shading-pattern-to-every-second-r.cs` | Create a table with alternating row colors by applying a shading pattern to every second row in a loop. | table-workflow | docx | mcp |
| 79 | `create-a-table-with-alternating-column-shading-by-applying-a-shading-pattern-to-every-seco.cs` | Create a table with alternating column shading by applying a shading pattern to every second column. | table-workflow | docx | mcp |
| 80 | `create-a-table-with-a-single-column-and-multiple-rows-then-merge-every-second-cell-vertica.cs` | Create a table with a single column and multiple rows, then merge every second cell vertically for grouping. | table-workflow | docx | mcp |
| 81 | `create-a-table-with-a-fixed-number-of-columns-and-dynamically-add-rows-from-a-data-source.cs` | Create a table with a fixed number of columns and dynamically add rows from a data source collection. | table-workflow | docx | mcp |
| 82 | `insert-a-table-using-documentbuilder-specify-column-count-and-define-initial-rows.cs` | Insert a table using DocumentBuilder, specify column count, and define initial rows. | table-workflow | doc | mcp |
| 83 | `build-a-table-from-a-datatable-by-iterating-rows-and-inserting-cells-with-documentbuilder.cs` | Build a table from a DataTable by iterating rows and inserting cells with DocumentBuilder. | table-workflow | doc | mcp |
| 84 | `insert-a-table-with-a-caption-that-automatically-updates-numbering-when-new-tables-are-add.cs` | Insert a table with a caption that automatically updates numbering when new tables are added. | table-workflow | docx | mcp |
| 85 | `insert-a-table-at-a-specific-bookmark-location-by-moving-documentbuilder-to-the-bookmark-b.cs` | Insert a table at a specific bookmark location by moving DocumentBuilder to the bookmark before building. | table-workflow | doc | mcp |
| 86 | `insert-a-table-into-the-header-section-by-moving-documentbuilder-to-the-header-and-buildin.cs` | Insert a table into the header section by moving DocumentBuilder to the header and building the table. | table-workflow | doc | mcp |
| 87 | `insert-a-nested-table-inside-a-cell-by-creating-a-table-and-adding-it-to-the-parent-cell.cs` | Insert a nested table inside a cell by creating a Table and adding it to the parent cell. | table-workflow | docx | mcp |
| 88 | `insert-a-table-into-a-floating-textbox-by-moving-documentbuilder-to-the-textbox-and-buildi.cs` | Insert a table into a floating textbox by moving DocumentBuilder to the textbox and building the table. | table-workflow | doc | mcp |
| 89 | `insert-a-table-into-a-styled-paragraph-by-moving-documentbuilder-to-the-paragraph-and-buil.cs` | Insert a table into a styled paragraph by moving DocumentBuilder to the paragraph and building the table. | table-workflow | doc | mcp |
| 90 | `insert-a-table-into-a-footnote-by-moving-documentbuilder-to-the-footnote-and-building-the.cs` | Insert a table into a footnote by moving DocumentBuilder to the footnote and building the table. | table-workflow | doc | mcp |
| 91 | `insert-a-table-into-a-header-footer-by-moving-documentbuilder-to-the-respective-section-an.cs` | Insert a table into a header/footer by moving DocumentBuilder to the respective section and building the table. | table-workflow | doc | mcp |
| 92 | `set-table-preferredwidth-to-half-page-and-enable-auto-fit-to-adjust-columns-dynamically.cs` | Set table PreferredWidth to half page and enable auto‑fit to adjust columns dynamically. | table-workflow | docx | mcp |
| 93 | `disable-table-auto-fit-by-setting-allowautofit-property-to-false-keeping-column-widths-fix.cs` | Disable table auto‑fit by setting AllowAutoFit property to false, keeping column widths fixed. | table-workflow | docx | mcp |
| 94 | `configure-row-height-to-exactly-20-points-using-row-rowformat-height-and-heightrule-exact.cs` | Configure row height to exactly 20 points using Row.RowFormat.Height and HeightRule Exact. | table-workflow | docx | mcp |
| 95 | `apply-light-gray-shading-to-header-row-cells-via-cell-cellformat-shading-backgroundpattern.cs` | Apply light gray shading to header row cells via Cell.CellFormat.Shading.BackgroundPatternColor. | table-workflow | docx | mcp |
| 96 | `set-cell-padding-to-five-points-on-all-sides-by-modifying-cell-cellformat-padding-values.cs` | Set cell padding to five points on all sides by modifying Cell.CellFormat.Padding values. | table-workflow | docx | mcp |
| 97 | `add-a-double-border-around-the-entire-table-by-configuring-table-border-for-each-side.cs` | Add a double border around the entire table by configuring Table.Border for each side. | table-workflow | docx | mcp |
| 98 | `create-a-custom-tablestyle-named-customstyle-define-shading-and-borders-and-add-to-documen.cs` | Create a custom TableStyle named CustomStyle, define shading and borders, and add to document styles. | table-workflow | doc | mcp |
| 99 | `copy-an-existing-tablestyle-from-a-source-document-using-document-styles-addcopy-method.cs` | Copy an existing TableStyle from a source document using document.Styles.AddCopy method. | table-workflow | doc | mcp |
| 100 | `apply-a-tablestyle-to-a-table-by-setting-the-table-style-property-to-the-style-name.cs` | Apply a TableStyle to a table by setting the Table.Style property to the style name. | table-workflow | docx | mcp |
| 101 | `translate-a-tablestyle-into-direct-formatting-by-reading-its-properties-and-applying-them.cs` | Translate a TableStyle into direct formatting by reading its properties and applying them to Table, Row, and Cell. | table-workflow | docx | mcp |
| 102 | `join-two-adjacent-tables-into-a-single-continuous-table-using-the-table-join-method.cs` | Join two adjacent tables into a single continuous table using the Table.Join method. | table-workflow | docx | mcp |
| 103 | `split-a-large-table-at-a-specific-row-index-into-smaller-tables-using-table-split.cs` | Split a large table at a specific row index into smaller tables using Table.Split. | table-workflow | docx | mcp |
| 104 | `apply-distinct-border-styles-to-first-row-last-row-and-inner-cells-via-rowformat-and-cellf.cs` | Apply distinct border styles to first row, last row, and inner cells via RowFormat and CellFormat. | table-workflow | docx | mcp |
| 105 | `center-the-table-on-the-page-by-setting-table-alignment-property-to-center.cs` | Center the table on the page by setting Table.Alignment property to Center. | table-workflow | docx | mcp |
| 106 | `enable-table-text-wrapping-around-images-by-setting-table-allowoverlap-to-true-and-adjusti.cs` | Enable table text wrapping around images by setting Table.AllowOverlap to true and adjusting layout. | table-workflow | docx | mcp |
| 107 | `prevent-row-deletion-by-setting-row-rowformat-allowbreakacrosspages-to-false-in-the-docume.cs` | Prevent row deletion by setting Row.RowFormat.AllowBreakAcrossPages to false in the document. | table-workflow | doc | mcp |
| 108 | `apply-conditional-cell-shading-based-on-numeric-values-by-iterating-cells-and-setting-cell.cs` | Apply conditional cell shading based on numeric values by iterating cells and setting CellFormat.Shading. | table-workflow | docx | mcp |
| 109 | `merge-cells-horizontally-across-multiple-columns-using-cell-merge-on-adjacent-cells.cs` | Merge cells horizontally across multiple columns using Cell.Merge on adjacent cells. | table-workflow | docx | mcp |
| 110 | `merge-cells-vertically-within-a-column-by-calling-cell-merge-on-consecutive-rows.cs` | Merge cells vertically within a column by calling Cell.Merge on consecutive rows. | table-workflow | docx | mcp |
| 111 | `set-table-direction-to-right-to-left-for-arabic-text-by-enabling-table-righttoleft-propert.cs` | Set table direction to right‑to‑left for Arabic text by enabling Table.RightToLeft property. | table-workflow | docx | mcp |
| 112 | `preserve-original-table-formatting-when-loading-a-document-by-disabling-loadoptions-preser.cs` | Preserve original table formatting when loading a document by disabling LoadOptions.PreserveFormatting. | table-workflow | doc | mcp |
| 113 | `optimize-table-rendering-performance-for-large-documents-by-turning-off-layout-updates-dur.cs` | Optimize table rendering performance for large documents by turning off layout updates during batch modifications. | table-workflow | doc | mcp |
| 114 | `generate-a-pdf-from-a-docx-containing-tables-while-maintaining-all-table-styles-and-format.cs` | Generate a PDF from a DOCX containing tables while maintaining all table styles and formatting. | table-workflow | docx, doc, pdf | mcp |
| 115 | `convert-a-docx-with-complex-tables-to-html-ensuring-borders-and-cell-padding-are-retained.cs` | Convert a DOCX with complex tables to HTML, ensuring borders and cell padding are retained. | table-workflow | docx, doc, html | mcp |
| 116 | `serialize-a-table-s-structure-and-content-to-json-by-iterating-rows-and-cells-and-capturin.cs` | Serialize a table's structure and content to JSON by iterating rows and cells and capturing formatting. | table-workflow | json | mcp |
| 117 | `deserialize-json-table-data-back-into-a-word-document-by-creating-a-table-and-applying-sto.cs` | Deserialize JSON table data back into a Word document by creating a table and applying stored formatting. | table-workflow | doc, json | mcp |
| 118 | `apply-gradient-shading-to-cells-by-setting-cellformat-shading-texture-and-backgroundpatter.cs` | Apply gradient shading to cells by setting CellFormat.Shading.Texture and BackgroundPatternColor properties. | table-workflow | docx | mcp |
| 119 | `set-uniform-column-widths-for-all-cells-by-iterating-columns-and-assigning-cellformat-widt.cs` | Set uniform column widths for all cells by iterating columns and assigning CellFormat.Width values. | table-workflow | docx | mcp |
| 120 | `adjust-row-spacing-by-configuring-row-rowformat-spaceafter-and-spacebefore-properties-for.cs` | Adjust row spacing by configuring Row.RowFormat.SpaceAfter and SpaceBefore properties for each row. | table-workflow | docx | mcp |
| 121 | `insert-a-caption-above-a-table-using-documentbuilder-insertcaption-with-appropriate-label.cs` | Insert a caption above a table using DocumentBuilder.InsertCaption with appropriate label and numbering. | table-workflow | doc | mcp |
| 122 | `update-all-table-captions-after-adding-new-tables-by-iterating-tables-and-refreshing-capti.cs` | Update all table captions after adding new tables by iterating tables and refreshing caption numbers. | table-workflow | docx | mcp |
| 123 | `protect-a-table-from-editing-by-applying-a-read-only-restriction-using-document-protect-wi.cs` | Protect a table from editing by applying a read‑only restriction using Document.Protect with editing type. | table-workflow | doc | mcp |
| 124 | `export-a-table-as-an-image-by-rendering-the-table-node-to-a-bitmap-and-saving-as-png.cs` | Export a table as an image by rendering the table node to a bitmap and saving as PNG. | table-workflow | docx | mcp |
| 125 | `replace-an-existing-table-with-a-template-by-locating-the-original-table-node-and-insertin.cs` | Replace an existing table with a template by locating the original Table node and inserting the new one. | table-workflow | docx | mcp |
| 126 | `apply-different-cell-background-colors-based-on-column-index-by-iterating-cells-and-settin.cs` | Apply different cell background colors based on column index by iterating cells and setting CellFormat.Shading. | table-workflow | docx | mcp |
| 127 | `set-table-to-repeat-header-row-on-each-page-by-enabling-row-isheader-property-for-the-firs.cs` | Set table to repeat header row on each page by enabling Row.IsHeader property for the first row. | table-workflow | docx | mcp |
| 128 | `remove-all-empty-rows-from-a-table-by-checking-row-rowformat-height-and-deleting-rows-with.cs` | Remove all empty rows from a table by checking Row.RowFormat.Height and deleting rows with zero height. | table-workflow | docx | mcp |
| 129 | `convert-a-table-with-merged-cells-into-a-flat-structure-by-unmerging-cells-and-redistribut.cs` | Convert a table with merged cells into a flat structure by unmerging cells and redistributing content appropriately. | table-workflow | docx | mcp |
| 130 | `add-a-footer-row-with-totals-by-calculating-column-sums-and-inserting-a-new-formatted-row.cs` | Add a footer row with totals by calculating column sums and inserting a new formatted row. | table-workflow | docx | mcp |
| 131 | `implement-pagination-for-a-large-table-by-setting-table-allowbreakacrosspages-to-true-and.cs` | Implement pagination for a large table by setting Table.AllowBreakAcrossPages to true and adjusting row heights. | table-workflow | docx | mcp |
| 132 | `create-a-table-with-alternating-row-colors-for-readability-by-applying-shading-based-on-ro.cs` | Create a table with alternating row colors for readability by applying shading based on row index parity. | table-workflow | docx | mcp |
| 133 | `export-table-layout-to-xml-by-iterating-rows-and-cells-and-writing-attributes-for-widths-a.cs` | Export table layout to XML by iterating rows and cells and writing attributes for widths and styles. | table-workflow | xml | mcp |
| 134 | `import-a-table-definition-from-xml-by-parsing-elements-and-constructing-a-table-with-colum.cs` | Import a table definition from XML by parsing elements and constructing a table with column widths and styles. | table-workflow | xml | mcp |
| 135 | `apply-a-custom-border-color-to-the-first-column-cells-by-setting-cellformat-border-color-p.cs` | Apply a custom border color to the first column cells by setting CellFormat.Border.Color property. | table-workflow | docx | mcp |
| 136 | `set-table-to-automatically-resize-columns-to-fit-content-by-enabling-table-autofit-behavio.cs` | Set table to automatically resize columns to fit content by enabling Table.AutoFit behavior. | table-workflow | docx | mcp |
| 137 | `disable-automatic-column-resizing-while-preserving-existing-widths-by-setting-table-autofi.cs` | Disable automatic column resizing while preserving existing widths by setting Table.AutoFit property to false. | table-workflow | docx | mcp |
| 138 | `add-a-toc-entry-for-a-table-by-inserting-a-bookmark-and-referencing-it-in-the-toc-field.cs` | Add a TOC entry for a table by inserting a bookmark and referencing it in the TOC field. | table-workflow | docx | mcp |
| 139 | `update-table-style-definitions-programmatically-by-iterating-document-styles-and-modifying.cs` | Update table style definitions programmatically by iterating document.Styles and modifying TableStyle properties. | table-workflow | doc | mcp |
| 140 | `clone-an-existing-table-modify-its-content-and-insert-it-at-a-different-location-in-the-sa.cs` | Clone an existing table, modify its content, and insert it at a different location in the same document. | table-workflow | doc | mcp |
| 141 | `apply-a-diagonal-line-across-a-cell-by-setting-cellformat-diagonaldown-and-diagonalup-prop.cs` | Apply a diagonal line across a cell by setting CellFormat.DiagonalDown and DiagonalUp properties. | table-workflow | docx | mcp |
| 142 | `set-cell-text-direction-to-vertical-for-east-asian-languages-by-configuring-cellformat-tex.cs` | Set cell text direction to vertical for East Asian languages by configuring CellFormat.TextDirection property. | table-workflow | docx | mcp |
| 143 | `apply-a-custom-margin-around-the-entire-table-by-setting-table-leftindent-and-table-righti.cs` | Apply a custom margin around the entire table by setting Table.LeftIndent and Table.RightIndent properties. | table-workflow | docx | mcp |
| 144 | `set-table-vertical-alignment-within-the-page-by-adjusting-table-verticalalignment-property.cs` | Set table vertical alignment within the page by adjusting Table.VerticalAlignment property to Center. | table-workflow | docx | mcp |
| 145 | `apply-a-table-style-that-includes-banded-rows-and-columns-by-configuring-tablestyle-banded.cs` | Apply a table style that includes banded rows and columns by configuring TableStyle.BandedRows and BandedColumns. | table-workflow | docx | mcp |
| 146 | `generate-a-report-with-multiple-tables-each-using-a-distinct-tablestyle-and-ensure-consist.cs` | Generate a report with multiple tables, each using a distinct TableStyle, and ensure consistent spacing between them. | table-workflow | docx | mcp |
| 147 | `convert-a-table-with-complex-merged-cells-to-plain-html-by-calculating-appropriate-colspan.cs` | Convert a table with complex merged cells to plain HTML by calculating appropriate colspan and rowspan attributes. | table-workflow | html | mcp |
| 148 | `apply-a-watermark-behind-a-table-by-inserting-a-shape-with-text-and-positioning-it-behind.cs` | Apply a watermark behind a table by inserting a Shape with text and positioning it behind table nodes. | table-workflow | docx | mcp |
| 149 | `set-table-text-wrapping-style-to-square-by-configuring-table-textwrapping-property-accordi.cs` | Set table text wrapping style to square by configuring Table.TextWrapping property accordingly. | table-workflow | docx | mcp |
| 150 | `enable-table-overlapping-with-other-floating-objects-by-setting-table-allowoverlap-to-true.cs` | Enable table overlapping with other floating objects by setting Table.AllowOverlap to true. | table-workflow | docx | mcp |
| 151 | `add-a-comment-to-a-specific-table-cell-by-locating-the-cell-and-using-documentbuilder-inse.cs` | Add a comment to a specific table cell by locating the cell and using DocumentBuilder.InsertComment. | table-workflow | doc | mcp |
| 152 | `retrieve-the-list-of-all-tables-in-a-document-by-iterating-nodes-of-type-nodetype-table.cs` | Retrieve the list of all tables in a document by iterating nodes of type NodeType.Table. | table-workflow | doc | mcp |
| 153 | `delete-a-table-that-contains-a-specific-keyword-by-searching-its-range-text-and-removing-t.cs` | Delete a table that contains a specific keyword by searching its Range.Text and removing the node. | table-workflow | docx | mcp |
| 154 | `apply-custom-cell-margins-by-setting-cell-cellformat-margintop-marginbottom-marginleft-and.cs` | Apply custom cell margins by setting Cell.CellFormat.MarginTop, MarginBottom, MarginLeft, and MarginRight properties. | table-workflow | docx | mcp |
| 155 | `set-table-background-shading-to-light-blue-by-configuring-table-shading-backgroundpatternc.cs` | Set table background shading to light blue by configuring Table.Shading.BackgroundPatternColor property. | table-workflow | docx | mcp |
| 156 | `export-each-table-in-a-document-to-separate-docx-files-while-preserving-their-individual-s.cs` | Export each table in a document to separate DOCX files while preserving their individual styles. | table-workflow | docx, doc | mcp |
| 157 | `import-multiple-tables-from-separate-docx-files-and-merge-them-into-a-single-document-main.cs` | Import multiple tables from separate DOCX files and merge them into a single document maintaining original formatting. | table-workflow | docx, doc | mcp |
| 158 | `apply-a-table-style-with-a-bold-header-row-and-italic-footer-row-by-customizing-style-prop.cs` | Apply a table style with a bold header row and italic footer row by customizing style properties. | table-workflow | docx | mcp |
| 159 | `set-table-left-and-right-margins-to-align-with-page-margins-by-calculating-page-width-and.cs` | Set table left and right margins to align with page margins by calculating page width and adjusting indents. | table-workflow | docx | mcp |
| 160 | `add-a-conditional-row-that-appears-only-when-a-data-value-exceeds-a-threshold-using-if-fie.cs` | Add a conditional row that appears only when a data value exceeds a threshold using IF fields. | table-workflow | docx | mcp |
| 161 | `apply-a-table-style-with-alternating-column-shading-by-setting-tablestyle-bandedcolumns-pr.cs` | Apply a table style with alternating column shading by setting TableStyle.BandedColumns property to true. | table-workflow | docx | mcp |
| 162 | `insert-a-table-with-a-fixed-number-of-rows-and-columns-then-populate-cells-from-a-json-arr.cs` | Insert a table with a fixed number of rows and columns, then populate cells from a JSON array. | table-workflow | json | mcp |
| 163 | `update-cell-formulas-after-inserting-new-rows-by-recalculating-fields-using-document-updat.cs` | Update cell formulas after inserting new rows by recalculating fields using Document.UpdateFields method. | table-workflow | doc | mcp |
| 164 | `apply-a-custom-outer-border-style-while-disabling-inner-borders-by-configuring-table-borde.cs` | Apply a custom outer border style while disabling inner borders by configuring Table.Border settings appropriately. | table-workflow | docx | mcp |
| 165 | `set-table-text-direction-to-right-to-left-for-the-entire-table-by-enabling-table-righttole.cs` | Set table text direction to right‑to‑left for the entire table by enabling Table.RightToLeft property. | table-workflow | docx | mcp |
| 166 | `apply-a-table-style-with-a-thick-outer-border-and-thin-inner-borders-by-configuring-style.cs` | Apply a table style with a thick outer border and thin inner borders by configuring style Border properties. | table-workflow | docx | mcp |
| 167 | `generate-a-table-of-contents-for-tables-only-by-creating-a-custom-toc-field-that-reference.cs` | Generate a table of contents for tables only by creating a custom TOC field that references table captions. | table-workflow | docx | mcp |
| 168 | `configure-rows-to-expand-automatically-with-content-by-setting-each-row-rowformat-heightru.cs` | Configure rows to expand automatically with content by setting each Row.RowFormat.HeightRule to Auto. | table-workflow | docx | mcp |
| 169 | `apply-a-custom-cell-border-style-with-different-line-widths-for-each-side-by-setting-cellf.cs` | Apply a custom cell border style with different line widths for each side by setting CellFormat.Border properties. | table-workflow | docx | mcp |
| 170 | `apply-a-table-style-with-a-header-row-using-a-larger-font-size-by-customizing-the-style-s.cs` | Apply a table style with a header row using a larger font size by customizing the style's FontSize. | table-workflow | docx | mcp |
| 171 | `set-table-column-widths-to-match-data-length-by-measuring-text-width-and-assigning-cellfor.cs` | Set table column widths to match data length by measuring text width and assigning CellFormat.Width accordingly. | table-workflow | docx | mcp |
| 172 | `add-a-table-with-a-footer-row-that-calculates-totals-using-field-codes-for-column-sums.cs` | Add a table with a footer row that calculates totals using field codes for column sums. | table-workflow | docx | mcp |
| 173 | `apply-a-table-style-with-a-light-gray-background-for-all-cells-by-setting-tablestyle-shadi.cs` | Apply a table style with a light gray background for all cells by setting TableStyle.Shading property. | table-workflow | docx | mcp |
| 174 | `insert-a-table-with-alternating-column-background-colors-by-iterating-columns-and-applying.cs` | Insert a table with alternating column background colors by iterating columns and applying CellFormat.Shading. | table-workflow | docx | mcp |
| 175 | `set-table-text-direction-to-vertical-for-all-cells-by-configuring-cellformat-textdirection.cs` | Set table text direction to vertical for all cells by configuring CellFormat.TextDirection globally. | table-workflow | docx | mcp |
| 176 | `apply-a-custom-table-style-that-includes-a-double-line-border-around-the-table-and-single.cs` | Apply a custom table style that includes a double line border around the table and single lines inside. | table-workflow | docx | mcp |
| 177 | `apply-a-table-style-with-bold-header-row-and-italic-footer-row-by-customizing-style-proper.cs` | Apply a table style with bold header row and italic footer row by customizing style properties. | table-workflow | docx | mcp |
| 178 | `set-cell-vertical-alignment-to-middle-using-cell-cellformat-verticalalignment-property-for.cs` | Set cell vertical alignment to middle using Cell.CellFormat.VerticalAlignment property for better presentation. | table-workflow | docx | mcp |
| 179 | `apply-text-alignment-within-cells-by-setting-cell-cellformat-alignment-to-center-for-unifo.cs` | Apply text alignment within cells by setting Cell.CellFormat.Alignment to Center for uniform appearance. | table-workflow | docx | mcp |
| 180 | `define-cell-margins-by-configuring-cell-cellformat-margintop-marginbottom-marginleft-and-m.cs` | Define cell margins by configuring Cell.CellFormat.MarginTop, MarginBottom, MarginLeft, and MarginRight. | table-workflow | docx | mcp |
| 181 | `set-table-top-and-bottom-spacing-using-table-spacebefore-and-table-spaceafter-properties.cs` | Set table top and bottom spacing using Table.SpaceBefore and Table.SpaceAfter properties. | table-workflow | docx | mcp |
| 182 | `apply-table-background-shading-to-the-entire-table-by-setting-table-shading-backgroundpatt.cs` | Apply table background shading to the entire table by setting Table.Shading.BackgroundPatternColor. | table-workflow | docx | mcp |
| 183 | `prevent-rows-from-breaking-across-pages-by-setting-row-rowformat-keeptogether-to-true.cs` | Prevent rows from breaking across pages by setting Row.RowFormat.KeepTogether to true. | table-workflow | docx | mcp |
| 184 | `set-table-to-use-fixed-layout-by-disabling-table-autofit-and-specifying-explicit-column-wi.cs` | Set table to use fixed layout by disabling Table.AutoFit and specifying explicit column widths. | table-workflow | docx | mcp |
| 185 | `apply-a-predefined-built-in-tablestyle-by-setting-table-style-to-tablegrid.cs` | Apply a predefined built‑in TableStyle by setting Table.Style to "TableGrid". | table-workflow | docx | mcp |
| 186 | `retrieve-the-number-of-rows-and-columns-in-a-table-using-table-rows-count-and-table-rows-0.cs` | Retrieve the number of rows and columns in a table using Table.Rows.Count and Table.Rows[0].Cells.Count. | table-workflow | docx | mcp |
| 187 | `insert-a-new-row-at-a-specific-index-using-table-insertrow-method.cs` | Insert a new row at a specific index using Table.InsertRow method. | table-workflow | docx | mcp |
| 188 | `delete-a-specific-row-from-a-table-using-table-rows-index-remove-method.cs` | Delete a specific row from a table using Table.Rows[index].Remove method. | table-workflow | docx | mcp |
| 189 | `insert-a-new-column-at-a-specific-position-using-table-insertcolumn-method.cs` | Insert a new column at a specific position using Table.InsertColumn method. | table-workflow | docx | mcp |
| 190 | `delete-a-specific-column-from-a-table-using-table-columns-index-remove-method.cs` | Delete a specific column from a table using Table.Columns[index].Remove method. | table-workflow | docx | mcp |
| 191 | `merge-cells-across-a-row-by-calling-cell-merge-on-adjacent-cells.cs` | Merge cells across a row by calling Cell.Merge on adjacent cells. | table-workflow | docx | mcp |
| 192 | `split-a-merged-cell-back-into-separate-cells-using-cell-split-method.cs` | Split a merged cell back into separate cells using Cell.Split method. | table-workflow | docx | mcp |
| 193 | `apply-a-table-style-that-includes-a-header-row-with-a-different-background-color-by-settin.cs` | Apply a table style that includes a header row with a different background color by setting TableStyle.HeaderRow properties. | table-workflow | docx | mcp |
| 194 | `set-table-left-and-right-indents-to-align-with-paragraph-indentation-using-table-leftinden.cs` | Set table left and right indents to align with paragraph indentation using Table.LeftIndent and Table.RightIndent. | table-workflow | docx | mcp |
| 195 | `keep-the-entire-table-on-a-single-page-by-setting-row-rowformat-keeptogether-to-true-for-a.cs` | Keep the entire table on a single page by setting Row.RowFormat.KeepTogether to true for all rows. | table-workflow | docx | mcp |
| 196 | `apply-a-custom-table-style-that-includes-alternating-row-shading-by-setting-tablestyle-ban.cs` | Apply a custom table style that includes alternating row shading by setting TableStyle.BandedRows to true. | table-workflow | docx | mcp |
| 197 | `enable-table-to-allow-overlapping-with-other-floating-objects-by-setting-table-allowoverla.cs` | Enable table to allow overlapping with other floating objects by setting Table.AllowOverlap to true. | table-workflow | docx | mcp |
| 198 | `update-table-fields-after-modifying-content-by-calling-document-updatefields-to-recalculat.cs` | Update table fields after modifying content by calling Document.UpdateFields to recalculate any embedded calculations. | table-workflow | doc | mcp |

## Common failure patterns seen during generation and how they were corrected

### Inventing unsupported builder or table APIs

- Symptom: Build failures caused by `DocumentBuilder.InsertTable`, unsupported table properties, or unsupported style-option enum members.
- Fix: Use `StartTable`, `InsertCell`, `EndRow`, `EndTable`, and only package-version-safe table members.

### Inserting tables into invalid DOM locations

- Symptom: Runtime failures when a `Table` node is inserted directly into a bookmark, header, or footer location that rejects the node.
- Fix: Move the builder to a valid paragraph or container and build the table there.

### Applying formatting on the wrong object scope

- Symptom: Build failures or wrong behavior from using cell-level properties for paragraph alignment or unsupported table-level border properties.
- Fix: Apply formatting on the correct scope such as paragraphs inside cells, or through documented border and shading APIs already proven in the environment.

## Notes for maintainers

- The selected file for each task is the verified winner recorded in the batch run.
- This category performed best with light primary rules plus a narrow patch for builder insertion, property correctness, and validation scope.
- One published file name is reused by two near-duplicate header/footer tasks because the final ZIP contains one shared implementation file for both task rows.
- Preserve exact file-to-task traceability when updating the category.
