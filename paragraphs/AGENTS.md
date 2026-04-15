---
name: paragraphs
description: Verified C# examples for paragraph scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Paragraphs

## Purpose

This folder is a **live, curated example set** for paragraph scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free paragraph creation, formatting, traversal, export, and editing workflows using the real Aspose.Words DOM.

## Non-negotiable conventions

- Always use real `Paragraph` nodes and `ParagraphFormat` APIs for paragraph-level work.
- Prefer `DocumentBuilder` for simple paragraph insertion and direct DOM traversal for reporting, filtering, and structural edits.
- Bootstrap all required input assets locally for DOCX, DOC, or TXT workflows.
- Keep validation narrow and task-specific. Validate the exact paragraph change or output required by the task.
- For built-in paragraph styles, prefer `ParagraphFormat.StyleIdentifier`.
- For custom paragraph styles, prefer `ParagraphFormat.StyleName`.
- Do not use fake namespaces such as `Aspose.Words.Paragraphs`.
- Do not use unsupported APIs such as `Paragraph.GetLinesCount()` or `Run.Replace()`.

## Recommended workflow selection

- **Paragraph formatting workflow**: 16 examples
- **Insertion workflow**: 6 examples
- **Traversal/reporting workflow**: 6 examples
- **Text export workflow**: 4 examples
- **Editing workflow**: 3 examples

Use the simplest builder-based workflow for insertion and formatting tasks. Switch to direct paragraph traversal only when the task explicitly requires counting, filtering, exporting, or structural editing.

## Validation priorities

1. The code must compile and run without manual input.
2. Paragraph-targeted tasks must verify the exact paragraph or paragraph scope affected by the operation.
3. TXT import and export tasks must create and use local sample files instead of machine-specific paths.
4. Formatting tasks must verify the actual `ParagraphFormat` property on the intended paragraph.
5. Reporting tasks must produce deterministic output without relying on unsupported layout or paragraph-count helpers.

## Quality verification summary

The publish-ready set was checked for the final category folder and all **35** examples are present. A spot-quality review also confirmed that the final files do **not** contain the earlier bad patterns seen during generation, including fake namespaces like `Aspose.Words.Paragraphs`, unsupported calls such as `Paragraph.GetLinesCount()`, and invalid `Run.Replace()` usage.

## File-to-task reference

- `insert-a-new-paragraph-with-custom-text-using-documentbuilder-writeln-inside-a-document-lo.cs`
  - Task: Insert a new paragraph with custom text using DocumentBuilder.Writeln inside a document loop.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: mcp
- `insert-an-empty-paragraph-at-a-specific-node-using-documentbuilder-insertparagraph-for-pre.cs`
  - Task: Insert an empty paragraph at a specific node using DocumentBuilder.InsertParagraph for precise placement.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: mcp
- `set-paragraph-alignment-to-center-by-modifying-builder-currentparagraph-paragraphformat-al.cs`
  - Task: Set paragraph alignment to center by modifying builder.CurrentParagraph.ParagraphFormat.Alignment property for the current paragraph.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `adjust-paragraph-line-spacing-to-1-5-lines-by-setting-builder-currentparagraph-paragraphfo.cs`
  - Task: Adjust paragraph line spacing to 1.5 lines by setting builder.CurrentParagraph.ParagraphFormat.LineSpacing to 1.5.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `set-first-line-indent-of-a-paragraph-to-half-an-inch-using-paragraphformat-firstlineindent.cs`
  - Task: Set first line indent of a paragraph to half an inch using ParagraphFormat.FirstLineIndent property.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `apply-built-in-heading1-style-to-the-current-paragraph-by-assigning-paragraphformat-stylei.cs`
  - Task: Apply built‑in Heading1 style to the current paragraph by assigning ParagraphFormat.StyleIdentifier to StyleIdentifier.Heading1.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-custom-paragraph-style-named-mystyle-by-setting-paragraphformat-stylename-property.cs`
  - Task: Apply a custom paragraph style named MyStyle by setting ParagraphFormat.StyleName property to the style name.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `use-styleidentifier-enumeration-to-select-a-locale-independent-style-for-a-paragraph-ensur.cs`
  - Task: Use StyleIdentifier enumeration to select a locale‑independent style for a paragraph, ensuring consistent formatting.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-the-number-of-lines-in-a-paragraph-by-calling-paragraph-getlinescount-and-store-t.cs`
  - Task: Retrieve the number of lines in a paragraph by calling Paragraph.GetLinesCount() and store the result.
  - Workflow: traversal-reporting
  - Outputs: docx
  - Selected engine: mcp
- `log-each-paragraph-s-line-count-to-the-console-for-debugging-paragraph-layout-issues.cs`
  - Task: Log each paragraph’s line count to the console for debugging paragraph layout issues.
  - Workflow: traversal-reporting
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-style-separator-run-between-two-text-runs-to-combine-different-paragraph-styles-w.cs`
  - Task: Insert a style separator run between two text runs to combine different paragraph styles within one paragraph.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `search-document-runs-for-styleidentifier-styleseparator-to-identify-existing-style-separat.cs`
  - Task: Search document runs for StyleIdentifier.StyleSeparator to identify existing style separators for further processing.
  - Workflow: traversal-reporting
  - Outputs: docx
  - Selected engine: mcp
- `enable-bi-directional-marks-before-saving-to-plain-text-by-setting-document-addbidimarks-p.cs`
  - Task: Enable bi‑directional marks before saving to plain text by setting Document.AddBidiMarks property to true.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `export-a-document-to-txt-format-while-preserving-headers-and-footers-by-enabling-exporthea.cs`
  - Task: Export a document to TXT format while preserving headers and footers by enabling ExportHeadersFooters option during save.
  - Workflow: text-export
  - Outputs: txt
  - Selected engine: mcp
- `export-a-document-to-txt-without-headers-and-footers-by-setting-exportheadersfooters-optio.cs`
  - Task: Export a document to TXT without headers and footers by setting ExportHeadersFooters option to false.
  - Workflow: text-export
  - Outputs: txt
  - Selected engine: mcp
- `load-a-plain-text-file-with-detectnumberingwithwhitespaces-enabled-to-automatically-recogn.cs`
  - Task: Load a plain‑text file with DetectNumberingWithWhitespaces enabled to automatically recognize list items.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `convert-detected-plain-text-list-numbering-into-word-list-structures-after-loading-the-doc.cs`
  - Task: Convert detected plain‑text list numbering into Word list structures after loading the document successfully.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `navigate-to-a-specific-paragraph-index-using-documentbuilder-movetoparagraph-before-applyi.cs`
  - Task: Navigate to a specific paragraph index using DocumentBuilder.MoveToParagraph before applying formatting changes programmatically.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `modify-paragraph-spacing-before-and-after-by-setting-paragraphformat-spacebefore-and-space.cs`
  - Task: Modify paragraph spacing before and after by setting ParagraphFormat.SpaceBefore and SpaceAfter properties appropriately.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `set-paragraph-shading-background-color-to-light-gray-using-paragraphformat-shading-backgro.cs`
  - Task: Set paragraph shading background color to light gray using ParagraphFormat.Shading.BackgroundPatternColor property.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `apply-right-to-left-paragraph-direction-by-setting-paragraphformat-bidi-property-to-true-f.cs`
  - Task: Apply right‑to‑left paragraph direction by setting ParagraphFormat.Bidi property to true for Arabic text.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `insert-an-empty-paragraph-after-a-bookmark-by-moving-the-builder-to-the-bookmark-and-calli.cs`
  - Task: Insert an empty paragraph after a bookmark by moving the builder to the bookmark and calling InsertParagraph.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: mcp
- `remove-all-paragraphs-that-use-a-specific-style-identifier-such-as-styleidentifier-quote-f.cs`
  - Task: Remove all paragraphs that use a specific style identifier, such as StyleIdentifier.Quote, from the document.
  - Workflow: traversal-reporting
  - Outputs: docx
  - Selected engine: mcp
- `replace-text-within-a-paragraph-while-preserving-its-original-formatting-by-using-run-repl.cs`
  - Task: Replace text within a paragraph while preserving its original formatting by using Run.Replace method.
  - Workflow: editing
  - Outputs: docx
  - Selected engine: mcp
- `merge-consecutive-paragraphs-that-share-identical-formatting-into-a-single-paragraph-to-si.cs`
  - Task: Merge consecutive paragraphs that share identical formatting into a single paragraph to simplify document structure.
  - Workflow: editing
  - Outputs: docx
  - Selected engine: mcp
- `split-a-long-paragraph-into-multiple-shorter-paragraphs-at-specified-character-positions-t.cs`
  - Task: Split a long paragraph into multiple shorter paragraphs at specified character positions to improve readability.
  - Workflow: editing
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-paragraph-with-a-page-break-before-it-by-setting-paragraphformat-pagebreakbefore.cs`
  - Task: Insert a paragraph with a page break before it by setting ParagraphFormat.PageBreakBefore property to true.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: mcp
- `set-paragraph-outline-level-to-2-for-subheadings-to-ensure-proper-inclusion-in-generated-t.cs`
  - Task: Set paragraph outline level to 2 for subheadings to ensure proper inclusion in generated table of contents.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-specific-paragraph-style-to-all-heading-paragraphs-using-a-style-selector-loop-acr.cs`
  - Task: Apply a specific paragraph style to all heading paragraphs using a style selector loop across the document.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `count-the-number-of-paragraphs-that-contain-a-given-keyword-and-output-the-total-count-res.cs`
  - Task: Count the number of paragraphs that contain a given keyword and output the total count result.
  - Workflow: traversal-reporting
  - Outputs: docx
  - Selected engine: mcp
- `add-a-paragraph-containing-a-hyperlink-run-and-assign-the-built-in-hyperlink-style-to-it.cs`
  - Task: Add a paragraph containing a hyperlink run and assign the built‑in Hyperlink style to it.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-paragraph-with-a-field-code-such-as-date-and-format-the-field-result-using-paragr.cs`
  - Task: Insert a paragraph with a field code such as DATE and format the field result using paragraph formatting.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: mcp
- `determine-whether-a-paragraph-resides-inside-a-table-by-checking-paragraph-isintable-prope.cs`
  - Task: Determine whether a paragraph resides inside a table by checking Paragraph.IsInTable property value.
  - Workflow: traversal-reporting
  - Outputs: docx
  - Selected engine: mcp
- `export-paragraphs-to-plain-text-with-line-numbers-prefixed-by-enabling-line-numbering-opti.cs`
  - Task: Export paragraphs to plain text with line numbers prefixed by enabling line numbering option in SaveOptions.
  - Workflow: text-export
  - Outputs: docx
  - Selected engine: mcp
- `load-a-txt-file-with-list-detection-disabled-and-compare-the-resulting-document-s-list-for.cs`
  - Task: Load a TXT file with list detection disabled and compare the resulting document's list formatting to enabled mode.
  - Workflow: text-export
  - Outputs: txt, docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Fake Paragraph namespace**
  - Symptom: Compiler error that `Aspose.Words.Paragraphs` does not exist.
  - Preferred fix: Use only documented Aspose.Words namespaces such as `Aspose.Words`, `Aspose.Words.Saving`, and `Aspose.Words.Loading` where needed.
- **Unsupported paragraph line-count API**
  - Symptom: Compiler errors around `Paragraph.GetLinesCount()` or invented layout convenience members.
  - Preferred fix: Use a compile-safe approximation for line-count-oriented tasks instead of inventing unavailable APIs.
- **Invalid run replacement API**
  - Symptom: Compiler error that `Run` does not expose `Replace`.
  - Preferred fix: Replace text by editing `Run.Text` while preserving run formatting or by rebuilding runs carefully.
- **Wrong paragraph target during validation**
  - Symptom: Runtime verification fails because formatting was applied to a different paragraph than the one being checked.
  - Preferred fix: Capture or re-find the exact target paragraph and validate that paragraph only.
- **Transient MCP variation on simple tasks**
  - Symptom: First attempt may target the wrong paragraph scope or add unnecessary validation noise.
  - Preferred fix: Keep validation minimal and focused on the exact task requirement.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`
- Supporting packages seen in the verified set:
  - `Newtonsoft.Json` for JSON-oriented reporting patterns when needed
  - `Aspose.Drawing.Common` only where drawing/color interop appears in auxiliary examples

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Newtonsoft.Json
dotnet add package Aspose.Drawing.Common
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\paragraphs\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the verified publish-ready winner rather than a merely compiling draft.
- For simple paragraph tasks, avoid adding broad document-wide validation if the task only targets one paragraph or one output artifact.
