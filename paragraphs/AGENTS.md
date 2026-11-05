---
name: paragraphs
description: Verified C# examples for paragraphs scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Paragraphs

## Purpose

Provide concise, reliable guidance for an AI coding assistant to generate correct C# snippets that manipulate paragraphs with Aspose.Words, ensuring only documented APIs are used and document integrity is preserved.

## Non-negotiable conventions

- Never reference namespaces or members that are not part of the official Aspose.Words API.
- Do not call Run.Replace; modify Run.Text directly to preserve formatting.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Insertion workflow**: 6 examples
- **Paragraph Formatting workflow**: 16 examples
- **Traversal Reporting workflow**: 6 examples
- **Text Export workflow**: 4 examples
- **Editing workflow**: 3 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Check that every referenced type, property, and method exists in Aspose.Words 26.10.0 before compilation.

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
  - Task: Apply built-in Heading1 style to the current paragraph by assigning ParagraphFormat.StyleIdentifier to StyleIdentifier.Heading1.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-custom-paragraph-style-named-mystyle-by-setting-paragraphformat-stylename-property.cs`
  - Task: Apply a custom paragraph style named MyStyle by setting ParagraphFormat.StyleName property to the style name.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `use-styleidentifier-enumeration-to-select-a-locale-independent-style-for-a-paragraph-ensur.cs`
  - Task: Use StyleIdentifier enumeration to select a locale-independent style for a paragraph, ensuring consistent formatting.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-the-number-of-lines-in-a-paragraph-by-calling-paragraph-getlinescount-and-store-t.cs`
  - Task: Retrieve the number of lines in a paragraph by calling Paragraph.GetLinesCount() and store the result.
  - Workflow: traversal-reporting
  - Outputs: docx
  - Selected engine: mcp
- `log-each-paragraph-s-line-count-to-the-console-for-debugging-paragraph-layout-issues.cs`
  - Task: Log each paragraph's line count to the console for debugging paragraph layout issues.
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
  - Task: Enable bi-directional marks before saving to plain text by setting Document.AddBidiMarks property to true.
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
  - Task: Load a plain-text file with DetectNumberingWithWhitespaces enabled to automatically recognize list items.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: mcp
- `convert-detected-plain-text-list-numbering-into-word-list-structures-after-loading-the-doc.cs`
  - Task: Convert detected plain-text list numbering into Word list structures after loading the document successfully.
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
  - Task: Apply right-to-left paragraph direction by setting ParagraphFormat.Bidi property to true for Arabic text.
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
  - Task: Add a paragraph containing a hyperlink run and assign the built-in Hyperlink style to it.
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
  - Selected engine: existing_repo

## Common failure patterns and preferred agent fixes

- **Use of non-existent API members**
  - Symptom: Compiler reports missing definitions for properties such as Document.AddBidiMarks, TxtSaveOptions.ExportHeadersFooters, or StyleIdentifier.StyleSeparator.
  - Root cause: The code references members that were never implemented in the Aspose.Words library.
  - Preferred fix: Replace the call with the supported alternative (e.g., configure TxtSaveOptions.AddBidiMarks, use the Unicode unit separator character for style separators, or omit unsupported options).
  - Prevention: Consult the official Aspose.Words API reference for the exact version before using a property or method.

- **Incorrect overload usage**
  - Symptom: Build fails because DocumentBuilder.MoveTo or MoveToParagraph is called with an invalid number of arguments.
  - Root cause: The code supplies parameters that do not match any overload defined in the library.
  - Preferred fix: Use the overload signatures documented for the current version, providing both paragraph index and character index for MoveToParagraph, or call MoveTo with a Node argument.
  - Prevention: Review method signatures in the API docs and pass the required arguments exactly as specified.

- **Applying paragraph style to a run**
  - Symptom: Runtime throws ArgumentException stating that the style is not a character style when assigning a StyleIdentifier to Font.StyleIdentifier.
  - Root cause: A paragraph-level style (e.g., Heading1) is being applied to a Run, which only accepts character styles.
  - Preferred fix: Assign a character style such as StyleIdentifier.DefaultParagraphFont to the run, or leave the run with the document's default character formatting.
  - Prevention: Distinguish between paragraph and character styles; use Font.StyleIdentifier only for character styles.

- **Assuming unsupported layout helpers**
  - Symptom: Compiler errors appear for calls to Paragraph.GetLinesCount or other invented layout APIs.
  - Root cause: The code attempts to use methods that are not part of the Aspose.Words API for line-count or layout information.
  - Preferred fix: Remove the unsupported call and, if line count is needed, approximate it by measuring the rendered height or by using a custom layout engine outside Aspose.Words.
  - Prevention: Limit paragraph-related logic to properties and methods documented in the API reference.

- **Over-broad paragraph count validation**
  - Symptom: Runtime failures occur when the test asserts the total number of paragraphs instead of the targeted subset.
  - Root cause: The validation logic checks the entire document rather than the specific paragraph(s) affected by the operation.
  - Preferred fix: Validate only the paragraph(s) that the task modifies or queries, using precise node references.
  - Prevention: Scope assertions to the exact nodes involved in the operation.

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
Copy-Item ..\paragraphs\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update the examples whenever a new Aspose.Words version introduces or deprecates API members, and verify all referenced members against the official documentation.
