---
name: programming-with-documents
description: Verified C# examples for document-level workflows in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Programming with Documents

## Purpose

This folder is a **live, curated example set** for document-level programming scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free use of documented Aspose.Words APIs for creating, loading, saving, converting, protecting, traversing, and updating documents.

## Non-negotiable conventions

- Always use documented Aspose.Words APIs directly.
- Always create local sample source documents when a task refers to an existing file, folder, stream, or template.
- Never assume external files or folders already exist.
- Keep validation narrow and task-specific.
- Save the requested output document or report whenever the task expects one.

## Recommended workflow selection

- **Document workflow**: 110 examples

This category performed best with a light-touch, documented document workflow. Avoid over-prescriptive validation or speculative helper APIs.

## Validation priorities

1. The code must compile and run without manual input.
2. Required sample inputs must be bootstrapped locally inside the example.
3. Requested output files must be created successfully.
4. Validation should focus only on the exact task outcome.

## File-to-task reference

- `create-a-first-page-header-with-a-centered-logo-image-using-documentbuilder.cs`
  - Task: Create a first‑page header with a centered logo image using DocumentBuilder.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `add-an-odd-page-header-containing-a-right-aligned-chapter-title-in-a-custom-font.cs`
  - Task: Add an odd‑page header containing a right‑aligned chapter title in a custom font.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-an-even-page-footer-that-shows-page-numbers-in-roman-numerals-via-insertfield.cs`
  - Task: Insert an even‑page footer that shows page numbers in Roman numerals via InsertField.
  - Workflow: document-workflow
  - Outputs: rtf
  - Selected engine: mcp
- `configure-header-types-for-first-odd-and-even-pages-by-setting-headerfooter-type.cs`
  - Task: Configure header types for first, odd, and even pages by setting HeaderFooter.Type.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `place-an-absolutely-positioned-image-in-the-header-using-insertimage-with-left-and-top-off.cs`
  - Task: Place an absolutely positioned image in the header using InsertImage with left and top offsets.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `enable-unique-first-page-header-content-by-setting-differentfirstpageheaderfooter-to-true.cs`
  - Task: Enable unique first‑page header content by setting DifferentFirstPageHeaderFooter to true.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `activate-odd-and-even-page-headers-by-setting-section-pagesetup-differentoddandevenpages-p.cs`
  - Task: Activate odd and even page headers by setting Section.PageSetup.DifferentOddAndEvenPages property.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `copy-header-content-from-the-previous-section-into-the-current-section-using-headerfooterc.cs`
  - Task: Copy header content from the previous section into the current section using HeaderFooterCollection.Add.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `remove-all-footers-from-a-section-by-clearing-its-section-footers-collection.cs`
  - Task: Remove all footers from a section by clearing its Section.Footers collection.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-footnote-at-the-current-cursor-position-with-custom-text-via-insertfootnote.cs`
  - Task: Insert a footnote at the current cursor position with custom text via InsertFootnote.
  - Workflow: document-workflow
  - Outputs: rtf
  - Selected engine: mcp
- `add-an-endnote-with-reference-text-using-insertfootnote-and-specify-endnote-type.cs`
  - Task: Add an endnote with reference text using InsertFootnote and specify Endnote type.
  - Workflow: document-workflow
  - Outputs: rtf
  - Selected engine: mcp
- `change-footnote-numbering-style-to-lower-case-roman-numerals-by-setting-document-footnoteo.cs`
  - Task: Change footnote numbering style to lower‑case Roman numerals by setting Document.FootnoteOptions.NumberStyle.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `configure-footnote-layout-to-use-three-columns-by-assigning-document-footnoteoptions-colum.cs`
  - Task: Configure footnote layout to use three columns by assigning Document.FootnoteOptions.Columns.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `position-footnotes-at-the-bottom-of-the-page-by-adjusting-footnoteoptions-position.cs`
  - Task: Position footnotes at the bottom of the page by adjusting FootnoteOptions.Position.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `update-all-footnote-and-endnote-references-after-modifications-by-calling-document-updatef.cs`
  - Task: Update all footnote and endnote references after modifications by calling Document.UpdateFields.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `insert-a-table-of-contents-at-the-document-start-using-inserttableofcontents-with-default.cs`
  - Task: Insert a table of contents at the document start using InsertTableOfContents with default switches.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `customize-toc-entry-levels-to-include-headings-1-3-by-specifying-the-o-1-3-switch.cs`
  - Task: Customize TOC entry levels to include headings 1‑3 by specifying the \o "1-3" switch.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `include-only-specific-styles-in-the-toc-by-using-the-t-switch-with-a-comma-separated-list.cs`
  - Task: Include only specific styles in the TOC by using the \t switch with a comma‑separated list.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `update-existing-toc-fields-after-adding-sections-by-invoking-document-updatefields.cs`
  - Task: Update existing TOC fields after adding sections by invoking Document.UpdateFields.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `rebuild-toc-page-numbers-by-calling-document-updatepagelayout-method-following-updatefield.cs`
  - Task: Rebuild TOC page numbers by calling Document.UpdatePageLayout method following UpdateFields.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `extract-all-toc-entries-by-iterating-over-document-range-fields-where-field-type-equals-fi.cs`
  - Task: Extract all TOC entries by iterating over Document.Range.Fields where Field.Type equals FieldType.TOC.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `remove-a-specific-toc-field-by-locating-it-and-calling-its-remove-method.cs`
  - Task: Remove a specific TOC field by locating it and calling its Remove method.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-custom-paragraph-style-to-toc-entries-by-modifying-the-toc-field-s-style-property.cs`
  - Task: Apply a custom paragraph style to TOC entries by modifying the TOC field’s style property.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-bookmark-named-chapterstart-at-the-beginning-of-each-chapter-using-documentbuilde.cs`
  - Task: Insert a bookmark named "ChapterStart" at the beginning of each chapter using DocumentBuilder.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `navigate-to-a-bookmark-called-conclusion-and-insert-a-summary-paragraph-with-documentbuild.cs`
  - Task: Navigate to a bookmark called "Conclusion" and insert a summary paragraph with DocumentBuilder.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `delete-a-bookmark-named-draft-and-its-content-by-using-document-range-bookmarks-remove.cs`
  - Task: Delete a bookmark named "Draft" and its content by using Document.Range.Bookmarks.Remove.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `list-all-bookmark-names-by-iterating-through-document-range-bookmarks-collection-in-the-do.cs`
  - Task: List all bookmark names by iterating through Document.Range.Bookmarks collection in the document.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-multi-section-document-where-each-section-has-distinct-header-and-footer-configur.cs`
  - Task: Create a multi‑section document where each section has distinct header and footer configurations.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `set-page-orientation-to-landscape-for-sections-containing-wide-tables-via-section-pagesetu.cs`
  - Task: Set page orientation to landscape for sections containing wide tables via Section.PageSetup.Orientation.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `adjust-the-left-margin-of-the-first-page-header-to-align-with-the-main-text-margin.cs`
  - Task: Adjust the left margin of the first‑page header to align with the main text margin.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-page-number-field-formatted-as-page-x-of-y-into-the-footer-using-insertfield-with.cs`
  - Task: Insert a page number field formatted as "Page X of Y" into the footer using InsertField with switches.
  - Workflow: document-workflow
  - Outputs: rtf
  - Selected engine: mcp
- `apply-bold-and-underline-formatting-to-header-text-by-configuring-documentbuilder-font-bef.cs`
  - Task: Apply bold and underline formatting to header text by configuring DocumentBuilder.Font before insertion.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-an-even-page-footer-that-displays-the-document-title-aligned-left-using-documentbui.cs`
  - Task: Create an even‑page footer that displays the document title aligned left using DocumentBuilder.InsertField.
  - Workflow: document-workflow
  - Outputs: doc, rtf
  - Selected engine: mcp
- `copy-an-existing-header-from-one-document-to-another-by-exporting-headerfooter-as-a-node-a.cs`
  - Task: Copy an existing header from one document to another by exporting HeaderFooter as a node and importing.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `remove-all-footnotes-while-preserving-endnotes-by-clearing-document-footnotes-collection.cs`
  - Task: Remove all footnotes while preserving endnotes by clearing Document.Footnotes collection.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-new-word-document-and-insert-a-table-with-five-rows-and-three-columns.cs`
  - Task: Create a new Word document and insert a table with five rows and three columns.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `merge-the-first-two-cells-of-the-second-row-in-the-table-to-form-a-single-cell.cs`
  - Task: Merge the first two cells of the second row in the table to form a single cell.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `split-the-merged-cell-back-into-two-separate-cells-using-the-splitcell-method.cs`
  - Task: Split the merged cell back into two separate cells using the SplitCell method.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-the-built-in-grid-table-5-dark-style-to-the-entire-table-for-consistent-formatting.cs`
  - Task: Apply the built‑in 'Grid Table 5 Dark' style to the entire table for consistent formatting.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `set-a-custom-background-color-for-the-header-row-of-the-table-using-tablestyle-properties.cs`
  - Task: Set a custom background color for the header row of the table using TableStyle properties.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `populate-a-table-directly-from-a-net-datatable-containing-employee-records-and-format-head.cs`
  - Task: Populate a table directly from a .NET DataTable containing employee records and format header row.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `adjust-column-widths-proportionally-to-fit-content-after-loading-data-into-the-table.cs`
  - Task: Adjust column widths proportionally to fit content after loading data into the table.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-new-row-at-the-end-of-the-table-and-fill-each-cell-with-placeholder-text.cs`
  - Task: Insert a new row at the end of the table and fill each cell with placeholder text.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `delete-the-third-column-from-the-table-while-preserving-remaining-cell-data.cs`
  - Task: Delete the third column from the table while preserving remaining cell data.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-nested-table-inside-a-cell-of-the-outer-table-to-display-detailed-information.cs`
  - Task: Create a nested table inside a cell of the outer table to display detailed information.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `export-the-constructed-table-to-an-html-fragment-while-preserving-cell-borders.cs`
  - Task: Export the constructed table to an HTML fragment while preserving cell borders.
  - Workflow: document-workflow
  - Outputs: html
  - Selected engine: mcp
- `load-an-existing-docx-file-locate-the-first-table-and-change-its-border-thickness.cs`
  - Task: Load an existing DOCX file, locate the first table, and change its border thickness.
  - Workflow: document-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `iterate-through-all-tables-in-a-document-and-apply-a-uniform-cell-padding-of-five-points.cs`
  - Task: Iterate through all tables in a document and apply a uniform cell padding of five points.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `convert-a-datatable-with-numeric-values-into-a-table-and-format-numeric-cells-as-currency.cs`
  - Task: Convert a DataTable with numeric values into a table and format numeric cells as currency.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `add-a-caption-below-the-table-using-a-paragraph-styled-as-caption-and-reference-it.cs`
  - Task: Add a caption below the table using a paragraph styled as 'Caption' and reference it.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-bulleted-list-with-three-items-and-customize-the-bullet-character-to-a-dash.cs`
  - Task: Create a bulleted list with three items and customize the bullet character to a dash.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `generate-a-numbered-list-with-roman-numeral-formatting-for-the-first-level-items.cs`
  - Task: Generate a numbered list with Roman numeral formatting for the first level items.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `define-a-second-list-level-with-alphabetic-lower-case-numbering-and-increased-indentation.cs`
  - Task: Define a second list level with alphabetic lower‑case numbering and increased indentation.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `restart-list-numbering-at-1-for-a-new-section-after-a-heading-paragraph.cs`
  - Task: Restart list numbering at 1 for a new section after a heading paragraph.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `limit-list-nesting-to-nine-levels-and-verify-that-deeper-levels-default-to-plain-text.cs`
  - Task: Limit list nesting to nine levels and verify that deeper levels default to plain text.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-the-listformat-removenumbers-method-to-convert-a-numbered-list-back-to-plain-paragra.cs`
  - Task: Apply the ListFormat.RemoveNumbers method to convert a numbered list back to plain paragraphs.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-multi-level-list-where-each-level-uses-a-different-list-style-defined-in-the-docu.cs`
  - Task: Insert a multi‑level list where each level uses a different list style defined in the document.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `set-the-document-s-ooxmlcompliance-property-to-ooxmlcompliance-strict-to-preserve-advanced.cs`
  - Task: Set the document's OoxmlCompliance property to OoxmlCompliance.Strict to preserve advanced list settings.
  - Workflow: document-workflow
  - Outputs: doc, xml
  - Selected engine: mcp
- `create-a-paragraph-write-a-long-sentence-and-count-the-number-of-lines-it-occupies.cs`
  - Task: Create a paragraph, write a long sentence, and count the number of lines it occupies.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-the-built-in-heading-2-style-to-a-paragraph-and-adjust-its-spacing-before-and-after.cs`
  - Task: Apply the built‑in 'Heading 2' style to a paragraph and adjust its spacing before and after.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `set-paragraph-alignment-to-justified-and-enable-word-wrap-for-better-layout-on-narrow-page.cs`
  - Task: Set paragraph alignment to justified and enable word wrap for better layout on narrow pages.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-paragraph-containing-a-hyperlink-and-configure-it-to-open-in-a-new-browser-tab.cs`
  - Task: Insert a paragraph containing a hyperlink and configure it to open in a new browser tab.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-custom-paragraph-style-that-includes-a-left-indent-of-0-5-inches-and-line-spacing.cs`
  - Task: Apply a custom paragraph style that includes a left indent of 0.5 inches and line spacing of 1.5.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `add-a-paragraph-with-a-drop-cap-character-and-specify-the-number-of-lines-it-spans.cs`
  - Task: Add a paragraph with a drop cap character and specify the number of lines it spans.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-custom-style-to-a-paragraph-that-includes-a-border-background-color-and-indentatio.cs`
  - Task: Apply a custom style to a paragraph that includes a border, background color, and indentation.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-paragraph-with-a-custom-tab-stop-at-2-inches-and-insert-text-that-aligns-to-it.cs`
  - Task: Create a paragraph with a custom tab stop at 2 inches and insert text that aligns to it.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-table-of-contents-paragraph-that-automatically-reflects-headings-and-list-items.cs`
  - Task: Insert a table of contents paragraph that automatically reflects headings and list items.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-conditional-formatting-rule-to-table-cells-that-highlights-values-above-a-threshol.cs`
  - Task: Apply a conditional formatting rule to table cells that highlights values above a threshold.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `add-a-paragraph-with-a-background-shading-of-light-yellow-to-highlight-important-informati.cs`
  - Task: Add a paragraph with a background shading of light yellow to highlight important information.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-paragraph-that-contains-a-field-for-page-number-and-format-it-as-roman-numerals.cs`
  - Task: Insert a paragraph that contains a field for page number and format it as Roman numerals.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-built-in-quote-style-to-a-paragraph-and-increase-its-left-indent-for-emphasis.cs`
  - Task: Apply a built‑in 'Quote' style to a paragraph and increase its left indent for emphasis.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `add-a-paragraph-with-a-custom-field-that-displays-the-document-s-author-name-dynamically.cs`
  - Task: Add a paragraph with a custom field that displays the document's author name dynamically.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `insert-a-paragraph-with-a-custom-line-height-of-150-percent-and-verify-visual-appearance.cs`
  - Task: Insert a paragraph with a custom line height of 150 percent and verify visual appearance.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `add-a-paragraph-with-a-hanging-indent-of-0-25-inches-for-citation-formatting.cs`
  - Task: Add a paragraph with a hanging indent of 0.25 inches for citation formatting.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-text-box-set-its-internal-margins-and-insert-a-paragraph-with-bold-text-inside.cs`
  - Task: Create a text box, set its internal margins, and insert a paragraph with bold text inside.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `link-two-text-boxes-sequentially-so-that-text-flows-from-the-first-to-the-second-automatic.cs`
  - Task: Link two text boxes sequentially so that text flows from the first to the second automatically.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `check-whether-a-given-text-box-is-the-head-of-a-linked-sequence-using-its-ishead-property.cs`
  - Task: Check whether a given text box is the head of a linked sequence using its IsHead property.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `break-the-forward-link-of-a-middle-text-box-to-stop-text-flow-to-the-subsequent-box.cs`
  - Task: Break the forward link of a middle text box to stop text flow to the subsequent box.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-the-parent-shape-of-a-text-box-and-change-its-rotation-angle-to-45-degrees.cs`
  - Task: Retrieve the parent Shape of a text box and change its rotation angle to 45 degrees.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `add-a-text-box-to-a-document-header-and-ensure-it-appears-on-every-page-of-the-section.cs`
  - Task: Add a text box to a document header and ensure it appears on every page of the section.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `insert-a-text-box-inside-a-table-cell-and-adjust-the-cell-s-padding-to-accommodate-the-box.cs`
  - Task: Insert a text box inside a table cell and adjust the cell's padding to accommodate the box.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-duplicate-a-text-box-and-place-the-copy-at-a-specified-absolute-position.cs`
  - Task: Programmatically duplicate a text box and place the copy at a specified absolute position.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-shadow-effect-to-a-text-box-shape-and-export-the-document-as-a-pdf.cs`
  - Task: Apply a shadow effect to a text box shape and export the document as a PDF.
  - Workflow: document-workflow
  - Outputs: doc, pdf
  - Selected engine: mcp
- `link-a-text-box-to-a-paragraph-so-that-overflow-text-automatically-continues-in-the-box.cs`
  - Task: Link a text box to a paragraph so that overflow text automatically continues in the box.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-text-box-inside-a-header-shape-and-ensure-it-appears-only-on-the-first-page.cs`
  - Task: Insert a text box inside a header shape and ensure it appears only on the first page.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-text-box-set-its-text-direction-to-right-to-left-and-insert-arabic-characters.cs`
  - Task: Create a text box, set its text direction to right‑to‑left, and insert Arabic characters.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-gradient-fill-to-a-text-box-and-ensure-the-gradient-renders-correctly-in-pdf.cs`
  - Task: Apply a gradient fill to a text box and ensure the gradient renders correctly in PDF.
  - Workflow: document-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `insert-a-text-box-then-set-its-anchor-position-to-be-relative-to-the-page-margin.cs`
  - Task: Insert a text box, then set its anchor position to be relative to the page margin.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `batch-process-a-set-of-documents-to-replace-all-tables-with-a-predefined-style-and-margin.cs`
  - Task: Batch process a set of documents to replace all tables with a predefined style and margin settings.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `iterate-through-each-document-in-a-collection-and-convert-all-numbered-lists-to-bulleted-l.cs`
  - Task: Iterate through each document in a collection and convert all numbered lists to bulleted lists.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-macro-like-routine-that-inserts-a-table-a-paragraph-and-a-linked-text-box-in-one.cs`
  - Task: Create a macro‑like routine that inserts a table, a paragraph, and a linked text box in one operation.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `generate-a-report-document-by-merging-data-rows-into-a-table-and-adding-a-summary-paragrap.cs`
  - Task: Generate a report document by merging data rows into a table and adding a summary paragraph.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `apply-a-custom-list-template-to-all-document-lists-and-enforce-a-maximum-of-nine-nesting-l.cs`
  - Task: Apply a custom list template to all document lists and enforce a maximum of nine nesting levels.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `export-a-document-containing-tables-and-text-boxes-to-an-xps-file-while-preserving-layout.cs`
  - Task: Export a document containing tables and text boxes to an XPS file while preserving layout.
  - Workflow: document-workflow
  - Outputs: doc, xps
  - Selected engine: mcp
- `insert-a-paragraph-with-a-field-code-that-displays-the-current-date-in-mmmm-dd-yyyy-format.cs`
  - Task: Insert a paragraph with a field code that displays the current date in 'MMMM dd, yyyy' format.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `set-the-document-s-compatibility-mode-to-word2010-before-adding-complex-list-structures.cs`
  - Task: Set the document's compatibility mode to Word2010 before adding complex list structures.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `set-the-document-s-compatibility-level-to-word2016-before-adding-complex-table-structures.cs`
  - Task: Set the document's compatibility level to Word2016 before adding complex table structures.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `set-the-document-s-default-font-to-calibri-size-11-and-apply-it-to-all-new-paragraphs.cs`
  - Task: Set the document's default font to Calibri, size 11, and apply it to all new paragraphs.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `set-the-document-s-language-property-to-french-and-adjust-paragraph-direction-for-right-to.cs`
  - Task: Set the document's language property to French and adjust paragraph direction for right‑to‑left scripts.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `apply-a-built-in-title-style-to-the-first-paragraph-and-ensure-it-appears-in-the-document.cs`
  - Task: Apply a built‑in 'Title' style to the first paragraph and ensure it appears in the document outline.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `add-a-comment-to-a-paragraph-and-programmatically-retrieve-the-comment-text-for-review.cs`
  - Task: Add a comment to a paragraph and programmatically retrieve the comment text for review.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-paragraph-with-a-custom-field-that-calculates-the-sum-of-numbers-in-a-table-colum.cs`
  - Task: Insert a paragraph with a custom field that calculates the sum of numbers in a table column.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `generate-a-document-that-contains-a-table-a-list-and-a-paragraph-then-save-as-odt-format.cs`
  - Task: Generate a document that contains a table, a list, and a paragraph, then save as ODT format.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `iterate-through-all-shapes-in-a-document-and-change-any-text-box-fill-color-to-light-gray.cs`
  - Task: Iterate through all shapes in a document and change any text box fill color to light gray.
  - Workflow: document-workflow
  - Outputs: doc
  - Selected engine: mcp
- `add-a-page-break-before-each-new-heading-paragraph-and-ensure-the-heading-style-remains-ap.cs`
  - Task: Add a page break before each new heading paragraph and ensure the heading style remains applied.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-table-with-alternating-row-shading-using-a-custom-table-style-defined-programmati.cs`
  - Task: Create a table with alternating row shading using a custom table style defined programmatically.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-table-style-that-includes-a-top-border-bottom-border-and-no-side-borders.cs`
  - Task: Apply a table style that includes a top border, bottom border, and no side borders.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-table-style-that-includes-alternating-column-shading-for-better-readability.cs`
  - Task: Apply a table style that includes alternating column shading for better readability.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-custom-paragraph-style-that-includes-both-left-and-right-borders-and-a-light-gray.cs`
  - Task: Apply a custom paragraph style that includes both left and right borders and a light gray background.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-numbered-list-where-each-level-uses-a-different-numbering-scheme-such-as-decimal.cs`
  - Task: Create a numbered list where each level uses a different numbering scheme, such as decimal and lower‑roman.
  - Workflow: document-workflow
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Assuming external input files exist**
  - Symptom: Runtime failures when loading a missing document, template, or folder.
  - Preferred fix: Create the sample file or folder locally inside the example before reopening it.
- **Over-validating unrelated document structure**
  - Symptom: Examples fail even though the requested operation succeeded.
  - Preferred fix: Validate only the exact requested output, property, or conversion result.
- **Invented APIs or fake namespaces**
  - Symptom: Build failures caused by non-existent Aspose.Words members or namespaces.
  - Preferred fix: Use only documented `Document`, `DocumentBuilder`, save/load options, property collections, and protection APIs.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required package

```bash
dotnet add package Aspose.Words --version 26.3.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\programming-with-documents\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve exact file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the verified winner from the latest batch report rather than a merely compiling draft.
- Bootstrap file-based inputs locally instead of depending on machine-specific paths.
