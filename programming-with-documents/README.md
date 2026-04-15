# Programming with Documents Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **Programming with Documents** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Programming with Documents**
- Slug: **programming-with-documents**
- Total examples: **110**
- Verified winners: **both=0**, **mcp-only=110**, **llm-only=0**, **none=0**
- Workflow examples: **110 / 110** use the standard document workflow

## Category rules that shaped these examples

- Use native Aspose.Words APIs directly.
- Create local sample source documents when a task refers to an existing file, folder, stream, or template.
- Do not assume external files or folders already exist.
- Keep validation narrow and task-specific.
- Prefer documented `Document` and `DocumentBuilder` workflows over speculative helpers.

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
Copy-Item ..\programming-with-documents\<example-file>.cs .\Program.cs

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
Copy-Item ..\programming-with-documents\create-a-first-page-header-with-a-centered-logo-image-using-documentbuilder.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `create-a-first-page-header-with-a-centered-logo-image-using-documentbuilder.cs` | Create a first‑page header with a centered logo image using DocumentBuilder. | document-workflow | doc | mcp |
| 2 | `add-an-odd-page-header-containing-a-right-aligned-chapter-title-in-a-custom-font.cs` | Add an odd‑page header containing a right‑aligned chapter title in a custom font. | document-workflow | docx | mcp |
| 3 | `insert-an-even-page-footer-that-shows-page-numbers-in-roman-numerals-via-insertfield.cs` | Insert an even‑page footer that shows page numbers in Roman numerals via InsertField. | document-workflow | rtf | mcp |
| 4 | `configure-header-types-for-first-odd-and-even-pages-by-setting-headerfooter-type.cs` | Configure header types for first, odd, and even pages by setting HeaderFooter.Type. | document-workflow | docx | mcp |
| 5 | `place-an-absolutely-positioned-image-in-the-header-using-insertimage-with-left-and-top-off.cs` | Place an absolutely positioned image in the header using InsertImage with left and top offsets. | document-workflow | docx | mcp |
| 6 | `enable-unique-first-page-header-content-by-setting-differentfirstpageheaderfooter-to-true.cs` | Enable unique first‑page header content by setting DifferentFirstPageHeaderFooter to true. | document-workflow | docx | mcp |
| 7 | `activate-odd-and-even-page-headers-by-setting-section-pagesetup-differentoddandevenpages-p.cs` | Activate odd and even page headers by setting Section.PageSetup.DifferentOddAndEvenPages property. | document-workflow | docx | mcp |
| 8 | `copy-header-content-from-the-previous-section-into-the-current-section-using-headerfooterc.cs` | Copy header content from the previous section into the current section using HeaderFooterCollection.Add. | document-workflow | docx | mcp |
| 9 | `remove-all-footers-from-a-section-by-clearing-its-section-footers-collection.cs` | Remove all footers from a section by clearing its Section.Footers collection. | document-workflow | docx | mcp |
| 10 | `insert-a-footnote-at-the-current-cursor-position-with-custom-text-via-insertfootnote.cs` | Insert a footnote at the current cursor position with custom text via InsertFootnote. | document-workflow | rtf | mcp |
| 11 | `add-an-endnote-with-reference-text-using-insertfootnote-and-specify-endnote-type.cs` | Add an endnote with reference text using InsertFootnote and specify Endnote type. | document-workflow | rtf | mcp |
| 12 | `change-footnote-numbering-style-to-lower-case-roman-numerals-by-setting-document-footnoteo.cs` | Change footnote numbering style to lower‑case Roman numerals by setting Document.FootnoteOptions.NumberStyle. | document-workflow | doc | mcp |
| 13 | `configure-footnote-layout-to-use-three-columns-by-assigning-document-footnoteoptions-colum.cs` | Configure footnote layout to use three columns by assigning Document.FootnoteOptions.Columns. | document-workflow | doc | mcp |
| 14 | `position-footnotes-at-the-bottom-of-the-page-by-adjusting-footnoteoptions-position.cs` | Position footnotes at the bottom of the page by adjusting FootnoteOptions.Position. | document-workflow | docx | mcp |
| 15 | `update-all-footnote-and-endnote-references-after-modifications-by-calling-document-updatef.cs` | Update all footnote and endnote references after modifications by calling Document.UpdateFields. | document-workflow | doc | mcp |
| 16 | `insert-a-table-of-contents-at-the-document-start-using-inserttableofcontents-with-default.cs` | Insert a table of contents at the document start using InsertTableOfContents with default switches. | document-workflow | doc | mcp |
| 17 | `customize-toc-entry-levels-to-include-headings-1-3-by-specifying-the-o-1-3-switch.cs` | Customize TOC entry levels to include headings 1‑3 by specifying the \o "1-3" switch. | document-workflow | docx | mcp |
| 18 | `include-only-specific-styles-in-the-toc-by-using-the-t-switch-with-a-comma-separated-list.cs` | Include only specific styles in the TOC by using the \t switch with a comma‑separated list. | document-workflow | docx | mcp |
| 19 | `update-existing-toc-fields-after-adding-sections-by-invoking-document-updatefields.cs` | Update existing TOC fields after adding sections by invoking Document.UpdateFields. | document-workflow | doc | mcp |
| 20 | `rebuild-toc-page-numbers-by-calling-document-updatepagelayout-method-following-updatefield.cs` | Rebuild TOC page numbers by calling Document.UpdatePageLayout method following UpdateFields. | document-workflow | doc | mcp |
| 21 | `extract-all-toc-entries-by-iterating-over-document-range-fields-where-field-type-equals-fi.cs` | Extract all TOC entries by iterating over Document.Range.Fields where Field.Type equals FieldType.TOC. | document-workflow | doc | mcp |
| 22 | `remove-a-specific-toc-field-by-locating-it-and-calling-its-remove-method.cs` | Remove a specific TOC field by locating it and calling its Remove method. | document-workflow | docx | mcp |
| 23 | `apply-a-custom-paragraph-style-to-toc-entries-by-modifying-the-toc-field-s-style-property.cs` | Apply a custom paragraph style to TOC entries by modifying the TOC field’s style property. | document-workflow | docx | mcp |
| 24 | `insert-a-bookmark-named-chapterstart-at-the-beginning-of-each-chapter-using-documentbuilde.cs` | Insert a bookmark named "ChapterStart" at the beginning of each chapter using DocumentBuilder. | document-workflow | doc | mcp |
| 25 | `navigate-to-a-bookmark-called-conclusion-and-insert-a-summary-paragraph-with-documentbuild.cs` | Navigate to a bookmark called "Conclusion" and insert a summary paragraph with DocumentBuilder. | document-workflow | doc | mcp |
| 26 | `delete-a-bookmark-named-draft-and-its-content-by-using-document-range-bookmarks-remove.cs` | Delete a bookmark named "Draft" and its content by using Document.Range.Bookmarks.Remove. | document-workflow | doc | mcp |
| 27 | `list-all-bookmark-names-by-iterating-through-document-range-bookmarks-collection-in-the-do.cs` | List all bookmark names by iterating through Document.Range.Bookmarks collection in the document. | document-workflow | doc | mcp |
| 28 | `create-a-multi-section-document-where-each-section-has-distinct-header-and-footer-configur.cs` | Create a multi‑section document where each section has distinct header and footer configurations. | document-workflow | doc | mcp |
| 29 | `set-page-orientation-to-landscape-for-sections-containing-wide-tables-via-section-pagesetu.cs` | Set page orientation to landscape for sections containing wide tables via Section.PageSetup.Orientation. | document-workflow | docx | mcp |
| 30 | `adjust-the-left-margin-of-the-first-page-header-to-align-with-the-main-text-margin.cs` | Adjust the left margin of the first‑page header to align with the main text margin. | document-workflow | docx | mcp |
| 31 | `insert-a-page-number-field-formatted-as-page-x-of-y-into-the-footer-using-insertfield-with.cs` | Insert a page number field formatted as "Page X of Y" into the footer using InsertField with switches. | document-workflow | rtf | mcp |
| 32 | `apply-bold-and-underline-formatting-to-header-text-by-configuring-documentbuilder-font-bef.cs` | Apply bold and underline formatting to header text by configuring DocumentBuilder.Font before insertion. | document-workflow | doc | mcp |
| 33 | `create-an-even-page-footer-that-displays-the-document-title-aligned-left-using-documentbui.cs` | Create an even‑page footer that displays the document title aligned left using DocumentBuilder.InsertField. | document-workflow | doc, rtf | mcp |
| 34 | `copy-an-existing-header-from-one-document-to-another-by-exporting-headerfooter-as-a-node-a.cs` | Copy an existing header from one document to another by exporting HeaderFooter as a node and importing. | document-workflow | doc | mcp |
| 35 | `remove-all-footnotes-while-preserving-endnotes-by-clearing-document-footnotes-collection.cs` | Remove all footnotes while preserving endnotes by clearing Document.Footnotes collection. | document-workflow | doc | mcp |
| 36 | `create-a-new-word-document-and-insert-a-table-with-five-rows-and-three-columns.cs` | Create a new Word document and insert a table with five rows and three columns. | document-workflow | doc | mcp |
| 37 | `merge-the-first-two-cells-of-the-second-row-in-the-table-to-form-a-single-cell.cs` | Merge the first two cells of the second row in the table to form a single cell. | document-workflow | docx | mcp |
| 38 | `split-the-merged-cell-back-into-two-separate-cells-using-the-splitcell-method.cs` | Split the merged cell back into two separate cells using the SplitCell method. | document-workflow | docx | mcp |
| 39 | `apply-the-built-in-grid-table-5-dark-style-to-the-entire-table-for-consistent-formatting.cs` | Apply the built‑in 'Grid Table 5 Dark' style to the entire table for consistent formatting. | document-workflow | docx | mcp |
| 40 | `set-a-custom-background-color-for-the-header-row-of-the-table-using-tablestyle-properties.cs` | Set a custom background color for the header row of the table using TableStyle properties. | document-workflow | docx | mcp |
| 41 | `populate-a-table-directly-from-a-net-datatable-containing-employee-records-and-format-head.cs` | Populate a table directly from a .NET DataTable containing employee records and format header row. | document-workflow | docx | mcp |
| 42 | `adjust-column-widths-proportionally-to-fit-content-after-loading-data-into-the-table.cs` | Adjust column widths proportionally to fit content after loading data into the table. | document-workflow | docx | mcp |
| 43 | `insert-a-new-row-at-the-end-of-the-table-and-fill-each-cell-with-placeholder-text.cs` | Insert a new row at the end of the table and fill each cell with placeholder text. | document-workflow | docx | mcp |
| 44 | `delete-the-third-column-from-the-table-while-preserving-remaining-cell-data.cs` | Delete the third column from the table while preserving remaining cell data. | document-workflow | docx | mcp |
| 45 | `create-a-nested-table-inside-a-cell-of-the-outer-table-to-display-detailed-information.cs` | Create a nested table inside a cell of the outer table to display detailed information. | document-workflow | docx | mcp |
| 46 | `export-the-constructed-table-to-an-html-fragment-while-preserving-cell-borders.cs` | Export the constructed table to an HTML fragment while preserving cell borders. | document-workflow | html | mcp |
| 47 | `load-an-existing-docx-file-locate-the-first-table-and-change-its-border-thickness.cs` | Load an existing DOCX file, locate the first table, and change its border thickness. | document-workflow | docx, doc | mcp |
| 48 | `iterate-through-all-tables-in-a-document-and-apply-a-uniform-cell-padding-of-five-points.cs` | Iterate through all tables in a document and apply a uniform cell padding of five points. | document-workflow | doc | mcp |
| 49 | `convert-a-datatable-with-numeric-values-into-a-table-and-format-numeric-cells-as-currency.cs` | Convert a DataTable with numeric values into a table and format numeric cells as currency. | document-workflow | docx | mcp |
| 50 | `add-a-caption-below-the-table-using-a-paragraph-styled-as-caption-and-reference-it.cs` | Add a caption below the table using a paragraph styled as 'Caption' and reference it. | document-workflow | docx | mcp |
| 51 | `create-a-bulleted-list-with-three-items-and-customize-the-bullet-character-to-a-dash.cs` | Create a bulleted list with three items and customize the bullet character to a dash. | document-workflow | docx | mcp |
| 52 | `generate-a-numbered-list-with-roman-numeral-formatting-for-the-first-level-items.cs` | Generate a numbered list with Roman numeral formatting for the first level items. | document-workflow | docx | mcp |
| 53 | `define-a-second-list-level-with-alphabetic-lower-case-numbering-and-increased-indentation.cs` | Define a second list level with alphabetic lower‑case numbering and increased indentation. | document-workflow | docx | mcp |
| 54 | `restart-list-numbering-at-1-for-a-new-section-after-a-heading-paragraph.cs` | Restart list numbering at 1 for a new section after a heading paragraph. | document-workflow | docx | mcp |
| 55 | `limit-list-nesting-to-nine-levels-and-verify-that-deeper-levels-default-to-plain-text.cs` | Limit list nesting to nine levels and verify that deeper levels default to plain text. | document-workflow | docx | mcp |
| 56 | `apply-the-listformat-removenumbers-method-to-convert-a-numbered-list-back-to-plain-paragra.cs` | Apply the ListFormat.RemoveNumbers method to convert a numbered list back to plain paragraphs. | document-workflow | docx | mcp |
| 57 | `insert-a-multi-level-list-where-each-level-uses-a-different-list-style-defined-in-the-docu.cs` | Insert a multi‑level list where each level uses a different list style defined in the document. | document-workflow | doc | mcp |
| 58 | `set-the-document-s-ooxmlcompliance-property-to-ooxmlcompliance-strict-to-preserve-advanced.cs` | Set the document's OoxmlCompliance property to OoxmlCompliance.Strict to preserve advanced list settings. | document-workflow | doc, xml | mcp |
| 59 | `create-a-paragraph-write-a-long-sentence-and-count-the-number-of-lines-it-occupies.cs` | Create a paragraph, write a long sentence, and count the number of lines it occupies. | document-workflow | docx | mcp |
| 60 | `apply-the-built-in-heading-2-style-to-a-paragraph-and-adjust-its-spacing-before-and-after.cs` | Apply the built‑in 'Heading 2' style to a paragraph and adjust its spacing before and after. | document-workflow | docx | mcp |
| 61 | `set-paragraph-alignment-to-justified-and-enable-word-wrap-for-better-layout-on-narrow-page.cs` | Set paragraph alignment to justified and enable word wrap for better layout on narrow pages. | document-workflow | docx | mcp |
| 62 | `insert-a-paragraph-containing-a-hyperlink-and-configure-it-to-open-in-a-new-browser-tab.cs` | Insert a paragraph containing a hyperlink and configure it to open in a new browser tab. | document-workflow | docx | mcp |
| 63 | `apply-a-custom-paragraph-style-that-includes-a-left-indent-of-0-5-inches-and-line-spacing.cs` | Apply a custom paragraph style that includes a left indent of 0.5 inches and line spacing of 1.5. | document-workflow | docx | mcp |
| 64 | `add-a-paragraph-with-a-drop-cap-character-and-specify-the-number-of-lines-it-spans.cs` | Add a paragraph with a drop cap character and specify the number of lines it spans. | document-workflow | docx | mcp |
| 65 | `apply-a-custom-style-to-a-paragraph-that-includes-a-border-background-color-and-indentatio.cs` | Apply a custom style to a paragraph that includes a border, background color, and indentation. | document-workflow | docx | mcp |
| 66 | `create-a-paragraph-with-a-custom-tab-stop-at-2-inches-and-insert-text-that-aligns-to-it.cs` | Create a paragraph with a custom tab stop at 2 inches and insert text that aligns to it. | document-workflow | docx | mcp |
| 67 | `insert-a-table-of-contents-paragraph-that-automatically-reflects-headings-and-list-items.cs` | Insert a table of contents paragraph that automatically reflects headings and list items. | document-workflow | docx | mcp |
| 68 | `apply-a-conditional-formatting-rule-to-table-cells-that-highlights-values-above-a-threshol.cs` | Apply a conditional formatting rule to table cells that highlights values above a threshold. | document-workflow | docx | mcp |
| 69 | `add-a-paragraph-with-a-background-shading-of-light-yellow-to-highlight-important-informati.cs` | Add a paragraph with a background shading of light yellow to highlight important information. | document-workflow | docx | mcp |
| 70 | `insert-a-paragraph-that-contains-a-field-for-page-number-and-format-it-as-roman-numerals.cs` | Insert a paragraph that contains a field for page number and format it as Roman numerals. | document-workflow | docx | mcp |
| 71 | `apply-a-built-in-quote-style-to-a-paragraph-and-increase-its-left-indent-for-emphasis.cs` | Apply a built‑in 'Quote' style to a paragraph and increase its left indent for emphasis. | document-workflow | docx | mcp |
| 72 | `add-a-paragraph-with-a-custom-field-that-displays-the-document-s-author-name-dynamically.cs` | Add a paragraph with a custom field that displays the document's author name dynamically. | document-workflow | doc | mcp |
| 73 | `insert-a-paragraph-with-a-custom-line-height-of-150-percent-and-verify-visual-appearance.cs` | Insert a paragraph with a custom line height of 150 percent and verify visual appearance. | document-workflow | docx | mcp |
| 74 | `add-a-paragraph-with-a-hanging-indent-of-0-25-inches-for-citation-formatting.cs` | Add a paragraph with a hanging indent of 0.25 inches for citation formatting. | document-workflow | docx | mcp |
| 75 | `create-a-text-box-set-its-internal-margins-and-insert-a-paragraph-with-bold-text-inside.cs` | Create a text box, set its internal margins, and insert a paragraph with bold text inside. | document-workflow | docx | mcp |
| 76 | `link-two-text-boxes-sequentially-so-that-text-flows-from-the-first-to-the-second-automatic.cs` | Link two text boxes sequentially so that text flows from the first to the second automatically. | document-workflow | docx | mcp |
| 77 | `check-whether-a-given-text-box-is-the-head-of-a-linked-sequence-using-its-ishead-property.cs` | Check whether a given text box is the head of a linked sequence using its IsHead property. | document-workflow | docx | mcp |
| 78 | `break-the-forward-link-of-a-middle-text-box-to-stop-text-flow-to-the-subsequent-box.cs` | Break the forward link of a middle text box to stop text flow to the subsequent box. | document-workflow | docx | mcp |
| 79 | `retrieve-the-parent-shape-of-a-text-box-and-change-its-rotation-angle-to-45-degrees.cs` | Retrieve the parent Shape of a text box and change its rotation angle to 45 degrees. | document-workflow | docx | mcp |
| 80 | `add-a-text-box-to-a-document-header-and-ensure-it-appears-on-every-page-of-the-section.cs` | Add a text box to a document header and ensure it appears on every page of the section. | document-workflow | doc | mcp |
| 81 | `insert-a-text-box-inside-a-table-cell-and-adjust-the-cell-s-padding-to-accommodate-the-box.cs` | Insert a text box inside a table cell and adjust the cell's padding to accommodate the box. | document-workflow | docx | mcp |
| 82 | `programmatically-duplicate-a-text-box-and-place-the-copy-at-a-specified-absolute-position.cs` | Programmatically duplicate a text box and place the copy at a specified absolute position. | document-workflow | docx | mcp |
| 83 | `apply-a-shadow-effect-to-a-text-box-shape-and-export-the-document-as-a-pdf.cs` | Apply a shadow effect to a text box shape and export the document as a PDF. | document-workflow | doc, pdf | mcp |
| 84 | `link-a-text-box-to-a-paragraph-so-that-overflow-text-automatically-continues-in-the-box.cs` | Link a text box to a paragraph so that overflow text automatically continues in the box. | document-workflow | docx | mcp |
| 85 | `insert-a-text-box-inside-a-header-shape-and-ensure-it-appears-only-on-the-first-page.cs` | Insert a text box inside a header shape and ensure it appears only on the first page. | document-workflow | docx | mcp |
| 86 | `create-a-text-box-set-its-text-direction-to-right-to-left-and-insert-arabic-characters.cs` | Create a text box, set its text direction to right‑to‑left, and insert Arabic characters. | document-workflow | docx | mcp |
| 87 | `apply-a-gradient-fill-to-a-text-box-and-ensure-the-gradient-renders-correctly-in-pdf.cs` | Apply a gradient fill to a text box and ensure the gradient renders correctly in PDF. | document-workflow | pdf | mcp |
| 88 | `insert-a-text-box-then-set-its-anchor-position-to-be-relative-to-the-page-margin.cs` | Insert a text box, then set its anchor position to be relative to the page margin. | document-workflow | docx | mcp |
| 89 | `batch-process-a-set-of-documents-to-replace-all-tables-with-a-predefined-style-and-margin.cs` | Batch process a set of documents to replace all tables with a predefined style and margin settings. | document-workflow | doc | mcp |
| 90 | `iterate-through-each-document-in-a-collection-and-convert-all-numbered-lists-to-bulleted-l.cs` | Iterate through each document in a collection and convert all numbered lists to bulleted lists. | document-workflow | doc | mcp |
| 91 | `create-a-macro-like-routine-that-inserts-a-table-a-paragraph-and-a-linked-text-box-in-one.cs` | Create a macro‑like routine that inserts a table, a paragraph, and a linked text box in one operation. | document-workflow | docx | mcp |
| 92 | `generate-a-report-document-by-merging-data-rows-into-a-table-and-adding-a-summary-paragrap.cs` | Generate a report document by merging data rows into a table and adding a summary paragraph. | document-workflow | doc | mcp |
| 93 | `apply-a-custom-list-template-to-all-document-lists-and-enforce-a-maximum-of-nine-nesting-l.cs` | Apply a custom list template to all document lists and enforce a maximum of nine nesting levels. | document-workflow | doc | mcp |
| 94 | `export-a-document-containing-tables-and-text-boxes-to-an-xps-file-while-preserving-layout.cs` | Export a document containing tables and text boxes to an XPS file while preserving layout. | document-workflow | doc, xps | mcp |
| 95 | `insert-a-paragraph-with-a-field-code-that-displays-the-current-date-in-mmmm-dd-yyyy-format.cs` | Insert a paragraph with a field code that displays the current date in 'MMMM dd, yyyy' format. | document-workflow | docx | mcp |
| 96 | `set-the-document-s-compatibility-mode-to-word2010-before-adding-complex-list-structures.cs` | Set the document's compatibility mode to Word2010 before adding complex list structures. | document-workflow | doc | mcp |
| 97 | `set-the-document-s-compatibility-level-to-word2016-before-adding-complex-table-structures.cs` | Set the document's compatibility level to Word2016 before adding complex table structures. | document-workflow | doc | mcp |
| 98 | `set-the-document-s-default-font-to-calibri-size-11-and-apply-it-to-all-new-paragraphs.cs` | Set the document's default font to Calibri, size 11, and apply it to all new paragraphs. | document-workflow | doc | mcp |
| 99 | `set-the-document-s-language-property-to-french-and-adjust-paragraph-direction-for-right-to.cs` | Set the document's language property to French and adjust paragraph direction for right‑to‑left scripts. | document-workflow | doc | mcp |
| 100 | `apply-a-built-in-title-style-to-the-first-paragraph-and-ensure-it-appears-in-the-document.cs` | Apply a built‑in 'Title' style to the first paragraph and ensure it appears in the document outline. | document-workflow | doc | mcp |
| 101 | `add-a-comment-to-a-paragraph-and-programmatically-retrieve-the-comment-text-for-review.cs` | Add a comment to a paragraph and programmatically retrieve the comment text for review. | document-workflow | docx | mcp |
| 102 | `insert-a-paragraph-with-a-custom-field-that-calculates-the-sum-of-numbers-in-a-table-colum.cs` | Insert a paragraph with a custom field that calculates the sum of numbers in a table column. | document-workflow | docx | mcp |
| 103 | `generate-a-document-that-contains-a-table-a-list-and-a-paragraph-then-save-as-odt-format.cs` | Generate a document that contains a table, a list, and a paragraph, then save as ODT format. | document-workflow | doc | mcp |
| 104 | `iterate-through-all-shapes-in-a-document-and-change-any-text-box-fill-color-to-light-gray.cs` | Iterate through all shapes in a document and change any text box fill color to light gray. | document-workflow | doc | mcp |
| 105 | `add-a-page-break-before-each-new-heading-paragraph-and-ensure-the-heading-style-remains-ap.cs` | Add a page break before each new heading paragraph and ensure the heading style remains applied. | document-workflow | docx | mcp |
| 106 | `create-a-table-with-alternating-row-shading-using-a-custom-table-style-defined-programmati.cs` | Create a table with alternating row shading using a custom table style defined programmatically. | document-workflow | docx | mcp |
| 107 | `apply-a-table-style-that-includes-a-top-border-bottom-border-and-no-side-borders.cs` | Apply a table style that includes a top border, bottom border, and no side borders. | document-workflow | docx | mcp |
| 108 | `apply-a-table-style-that-includes-alternating-column-shading-for-better-readability.cs` | Apply a table style that includes alternating column shading for better readability. | document-workflow | docx | mcp |
| 109 | `apply-a-custom-paragraph-style-that-includes-both-left-and-right-borders-and-a-light-gray.cs` | Apply a custom paragraph style that includes both left and right borders and a light gray background. | document-workflow | docx | mcp |
| 110 | `create-a-numbered-list-where-each-level-uses-a-different-numbering-scheme-such-as-decimal.cs` | Create a numbered list where each level uses a different numbering scheme, such as decimal and lower‑roman. | document-workflow | docx | mcp |

## Notes for maintainers

- The selected file for each task is the verified winner recorded in the batch run.
- This category performed best with **simple primary rules**. Avoid over-prescriptive rules that add unnecessary structural assertions.
- Preserve exact file-to-task traceability when updating the category.
- Bootstrap all sample input files locally inside the example when the task refers to an existing asset.
