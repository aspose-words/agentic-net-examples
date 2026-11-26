---
name: paragraphs
description: Verified C# examples for paragraphs scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Paragraphs

## Purpose

Assist developers in implementing paragraph-level manipulations with Aspose.Words by enforcing supported APIs, correct namespace usage, and precise validation while avoiding common compilation and runtime pitfalls.

## Non-negotiable conventions

- Do not call Run.Replace; modify Run.Text directly.
- Do not reference non-existent namespaces such as Aspose.Words.Building.
- Do not use Paragraph.GetLinesCount or any layout-specific line-count API.
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

1. Ensure that all required Aspose.Words namespaces are imported and that the project references the correct assembly version.
2. Confirm that every invoked member exists in the target version before compilation.
3. Scope assertions to the exact nodes affected by the operation.

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
  - Selected engine: llm
- `export-a-document-to-txt-without-headers-and-footers-by-setting-exportheadersfooters-optio.cs`
  - Task: Export a document to TXT without headers and footers by setting ExportHeadersFooters option to false.
  - Workflow: text-export
  - Outputs: txt
  - Selected engine: mcp
- `load-a-plain-text-file-with-detectnumberingwithwhitespaces-enabled-to-automatically-recogn.cs`
  - Task: Load a plain-text file with DetectNumberingWithWhitespaces enabled to automatically recognize list items.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: existing_repo
- `convert-detected-plain-text-list-numbering-into-word-list-structures-after-loading-the-doc.cs`
  - Task: Convert detected plain-text list numbering into Word list structures after loading the document successfully.
  - Workflow: paragraph-formatting
  - Outputs: docx
  - Selected engine: llm
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
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Incorrect namespace reference**
  - Symptom: Compiler reports CS0234 indicating that a referenced namespace such as Aspose.Words.Building does not exist.
  - Root cause: The code imports a namespace that is not part of the Aspose.Words library for the targeted version.
  - Preferred fix: Replace the invalid namespace with the correct one, e.g., Aspose.Words.Saving or Aspose.Words.Loading.
  - Prevention: Consult the official API reference for the exact namespace names before adding using directives.

- **Missing method overload**
  - Symptom: Compiler error CS1501 shows that InsertParagraph does not accept the supplied number of arguments.
  - Root cause: The code calls InsertParagraph with parameters that are not defined in the method signature.
  - Preferred fix: Call InsertParagraph() without arguments after moving the builder to the desired node.
  - Prevention: Verify method signatures in the documentation or IntelliSense before passing arguments.

- **Member not found on type**
  - Symptom: CS1061 or CS0117 errors indicate that members like Document.AddBidiMarks, TxtSaveOptions.ExportHeadersFooters, Paragraph.IsInTable, or Run.IsStyleSeparator are unavailable.
  - Root cause: The code uses properties or methods that are not implemented in the current Aspose.Words version.
  - Preferred fix: Use supported alternatives such as TxtSaveOptions (which includes headers by default) or detect table membership via Paragraph.GetAncestor(NodeType.Table).
  - Prevention: Check the version-specific API list and guard against unavailable members with conditional compilation or runtime checks.

- **Unsupported line-count API**
  - Symptom: Compilation fails because Paragraph.GetLinesCount does not exist.
  - Root cause: The code attempts to call a layout-specific helper that is not part of the public API.
  - Preferred fix: Remove the call and, if needed, approximate line count using document layout services or omit the feature.
  - Prevention: Restrict code to documented APIs; avoid assumptions about internal layout calculations.

- **Invalid run replacement**
  - Symptom: CS1061 indicates that Run.Replace is missing.
  - Root cause: The code tries to replace text via a non-existent method.
  - Preferred fix: Assign the new string to Run.Text while preserving formatting.
  - Prevention: Use Run.Text property for text updates and keep formatting unchanged.

- **Over-broad paragraph validation**
  - Symptom: Runtime failures occur when the test asserts the total number of paragraphs instead of the targeted one.
  - Root cause: Validation logic checks the whole document rather than the specific paragraph affected by the operation.
  - Preferred fix: Validate only the paragraph(s) directly involved in the task.
  - Prevention: Design assertions to target the exact nodes under test.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
```

```powershell
Copy-Item ..\paragraphs\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update the documentation whenever a new Aspose.Words version introduces or deprecates members referenced in the examples.
- Add unit tests for each verified workflow to detect regressions in API signatures.
- Periodically review the rules summary to incorporate newly discovered best-practice patterns.
