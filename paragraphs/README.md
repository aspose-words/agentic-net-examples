# Paragraphs Examples for Aspose.Words for .NET

The Paragraphs category covers insertion, formatting, traversal, and export of Word document paragraphs using Aspose.Words 26.10.0. Verified examples demonstrate safe DocumentBuilder usage, style-separator handling, bi-directional text options, TXT save options, and plain-text list detection.

## Snapshot

- Category: Paragraphs
- Slug: paragraphs
- Total examples: 35
- Publish-ready successful examples: 35 / 35
- Source run: 20261102_031946_08fb5b
- Insertion examples: 6
- Paragraph Formatting examples: 16
- Traversal Reporting examples: 6
- Text Export examples: 4
- Editing examples: 3

## Category rules that shaped these examples

- Never reference namespaces or members that are not part of the official Aspose.Words API.
- Do not call Run.Replace; modify Run.Text directly to preserve formatting.
- Use Paragraph.ParagraphFormat for all alignment, spacing, indent, bidi, shading, outline level, and page-break settings.
- Apply built-in styles via ParagraphFormat.StyleIdentifier and custom styles via ParagraphFormat.StyleName.
- Validate only the specific paragraph or range required for the task, avoiding blanket paragraph-count assertions.

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
Copy-Item ..\paragraphs\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `paragraphs/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0

# PowerShell example
Copy-Item ..\paragraphs\insert-a-new-paragraph-with-custom-text-using-documentbuilder-writeln-inside-a-document-lo.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `insert-a-new-paragraph-with-custom-text-using-documentbuilder-writeln-inside-a-document-lo.cs` | Insert a new paragraph with custom text using DocumentBuilder.Writeln inside a document loop. | insertion | docx | mcp |
| 2 | `insert-an-empty-paragraph-at-a-specific-node-using-documentbuilder-insertparagraph-for-pre.cs` | Insert an empty paragraph at a specific node using DocumentBuilder.InsertParagraph for precise placement. | insertion | docx | mcp |
| 3 | `set-paragraph-alignment-to-center-by-modifying-builder-currentparagraph-paragraphformat-al.cs` | Set paragraph alignment to center by modifying builder.CurrentParagraph.ParagraphFormat.Alignment property for the current paragraph. | paragraph-formatting | docx | mcp |
| 4 | `adjust-paragraph-line-spacing-to-1-5-lines-by-setting-builder-currentparagraph-paragraphfo.cs` | Adjust paragraph line spacing to 1.5 lines by setting builder.CurrentParagraph.ParagraphFormat.LineSpacing to 1.5. | paragraph-formatting | docx | mcp |
| 5 | `set-first-line-indent-of-a-paragraph-to-half-an-inch-using-paragraphformat-firstlineindent.cs` | Set first line indent of a paragraph to half an inch using ParagraphFormat.FirstLineIndent property. | paragraph-formatting | docx | mcp |
| 6 | `apply-built-in-heading1-style-to-the-current-paragraph-by-assigning-paragraphformat-stylei.cs` | Apply built-in Heading1 style to the current paragraph by assigning ParagraphFormat.StyleIdentifier to StyleIdentifier.Heading1. | paragraph-formatting | docx | mcp |
| 7 | `apply-a-custom-paragraph-style-named-mystyle-by-setting-paragraphformat-stylename-property.cs` | Apply a custom paragraph style named MyStyle by setting ParagraphFormat.StyleName property to the style name. | paragraph-formatting | docx | mcp |
| 8 | `use-styleidentifier-enumeration-to-select-a-locale-independent-style-for-a-paragraph-ensur.cs` | Use StyleIdentifier enumeration to select a locale-independent style for a paragraph, ensuring consistent formatting. | paragraph-formatting | docx | mcp |
| 9 | `retrieve-the-number-of-lines-in-a-paragraph-by-calling-paragraph-getlinescount-and-store-t.cs` | Retrieve the number of lines in a paragraph by calling Paragraph.GetLinesCount() and store the result. | traversal-reporting | docx | mcp |
| 10 | `log-each-paragraph-s-line-count-to-the-console-for-debugging-paragraph-layout-issues.cs` | Log each paragraph's line count to the console for debugging paragraph layout issues. | traversal-reporting | docx | mcp |
| 11 | `insert-a-style-separator-run-between-two-text-runs-to-combine-different-paragraph-styles-w.cs` | Insert a style separator run between two text runs to combine different paragraph styles within one paragraph. | paragraph-formatting | docx | mcp |
| 12 | `search-document-runs-for-styleidentifier-styleseparator-to-identify-existing-style-separat.cs` | Search document runs for StyleIdentifier.StyleSeparator to identify existing style separators for further processing. | traversal-reporting | docx | mcp |
| 13 | `enable-bi-directional-marks-before-saving-to-plain-text-by-setting-document-addbidimarks-p.cs` | Enable bi-directional marks before saving to plain text by setting Document.AddBidiMarks property to true. | paragraph-formatting | docx | mcp |
| 14 | `export-a-document-to-txt-format-while-preserving-headers-and-footers-by-enabling-exporthea.cs` | Export a document to TXT format while preserving headers and footers by enabling ExportHeadersFooters option during save. | text-export | txt | mcp |
| 15 | `export-a-document-to-txt-without-headers-and-footers-by-setting-exportheadersfooters-optio.cs` | Export a document to TXT without headers and footers by setting ExportHeadersFooters option to false. | text-export | txt | mcp |
| 16 | `load-a-plain-text-file-with-detectnumberingwithwhitespaces-enabled-to-automatically-recogn.cs` | Load a plain-text file with DetectNumberingWithWhitespaces enabled to automatically recognize list items. | paragraph-formatting | docx | mcp |
| 17 | `convert-detected-plain-text-list-numbering-into-word-list-structures-after-loading-the-doc.cs` | Convert detected plain-text list numbering into Word list structures after loading the document successfully. | paragraph-formatting | docx | mcp |
| 18 | `navigate-to-a-specific-paragraph-index-using-documentbuilder-movetoparagraph-before-applyi.cs` | Navigate to a specific paragraph index using DocumentBuilder.MoveToParagraph before applying formatting changes programmatically. | paragraph-formatting | docx | mcp |
| 19 | `modify-paragraph-spacing-before-and-after-by-setting-paragraphformat-spacebefore-and-space.cs` | Modify paragraph spacing before and after by setting ParagraphFormat.SpaceBefore and SpaceAfter properties appropriately. | paragraph-formatting | docx | mcp |
| 20 | `set-paragraph-shading-background-color-to-light-gray-using-paragraphformat-shading-backgro.cs` | Set paragraph shading background color to light gray using ParagraphFormat.Shading.BackgroundPatternColor property. | paragraph-formatting | docx | mcp |
| 21 | `apply-right-to-left-paragraph-direction-by-setting-paragraphformat-bidi-property-to-true-f.cs` | Apply right-to-left paragraph direction by setting ParagraphFormat.Bidi property to true for Arabic text. | paragraph-formatting | docx | mcp |
| 22 | `insert-an-empty-paragraph-after-a-bookmark-by-moving-the-builder-to-the-bookmark-and-calli.cs` | Insert an empty paragraph after a bookmark by moving the builder to the bookmark and calling InsertParagraph. | insertion | docx | mcp |
| 23 | `remove-all-paragraphs-that-use-a-specific-style-identifier-such-as-styleidentifier-quote-f.cs` | Remove all paragraphs that use a specific style identifier, such as StyleIdentifier.Quote, from the document. | traversal-reporting | docx | mcp |
| 24 | `replace-text-within-a-paragraph-while-preserving-its-original-formatting-by-using-run-repl.cs` | Replace text within a paragraph while preserving its original formatting by using Run.Replace method. | editing | docx | mcp |
| 25 | `merge-consecutive-paragraphs-that-share-identical-formatting-into-a-single-paragraph-to-si.cs` | Merge consecutive paragraphs that share identical formatting into a single paragraph to simplify document structure. | editing | docx | mcp |
| 26 | `split-a-long-paragraph-into-multiple-shorter-paragraphs-at-specified-character-positions-t.cs` | Split a long paragraph into multiple shorter paragraphs at specified character positions to improve readability. | editing | docx | mcp |
| 27 | `insert-a-paragraph-with-a-page-break-before-it-by-setting-paragraphformat-pagebreakbefore.cs` | Insert a paragraph with a page break before it by setting ParagraphFormat.PageBreakBefore property to true. | insertion | docx | mcp |
| 28 | `set-paragraph-outline-level-to-2-for-subheadings-to-ensure-proper-inclusion-in-generated-t.cs` | Set paragraph outline level to 2 for subheadings to ensure proper inclusion in generated table of contents. | paragraph-formatting | docx | mcp |
| 29 | `apply-a-specific-paragraph-style-to-all-heading-paragraphs-using-a-style-selector-loop-acr.cs` | Apply a specific paragraph style to all heading paragraphs using a style selector loop across the document. | paragraph-formatting | docx | mcp |
| 30 | `count-the-number-of-paragraphs-that-contain-a-given-keyword-and-output-the-total-count-res.cs` | Count the number of paragraphs that contain a given keyword and output the total count result. | traversal-reporting | docx | mcp |
| 31 | `add-a-paragraph-containing-a-hyperlink-run-and-assign-the-built-in-hyperlink-style-to-it.cs` | Add a paragraph containing a hyperlink run and assign the built-in Hyperlink style to it. | insertion | docx | mcp |
| 32 | `insert-a-paragraph-with-a-field-code-such-as-date-and-format-the-field-result-using-paragr.cs` | Insert a paragraph with a field code such as DATE and format the field result using paragraph formatting. | insertion | docx | mcp |
| 33 | `determine-whether-a-paragraph-resides-inside-a-table-by-checking-paragraph-isintable-prope.cs` | Determine whether a paragraph resides inside a table by checking Paragraph.IsInTable property value. | traversal-reporting | docx | mcp |
| 34 | `export-paragraphs-to-plain-text-with-line-numbers-prefixed-by-enabling-line-numbering-opti.cs` | Export paragraphs to plain text with line numbers prefixed by enabling line numbering option in SaveOptions. | text-export | docx | mcp |
| 35 | `load-a-txt-file-with-list-detection-disabled-and-compare-the-resulting-document-s-list-for.cs` | Load a TXT file with list detection disabled and compare the resulting document's list formatting to enabled mode. | text-export | txt, docx | existing_repo |

## Common failure patterns seen during generation and how they were corrected

### Use of non-existent API members

- Symptom: Compiler reports missing definitions for properties such as Document.AddBidiMarks, TxtSaveOptions.ExportHeadersFooters, or StyleIdentifier.StyleSeparator.
- Root cause: The code references members that were never implemented in the Aspose.Words library.
- Preferred fix: Replace the call with the supported alternative (e.g., configure TxtSaveOptions.AddBidiMarks, use the Unicode unit separator character for style separators, or omit unsupported options).
- Prevention: Consult the official Aspose.Words API reference for the exact version before using a property or method.

### Incorrect overload usage

- Symptom: Build fails because DocumentBuilder.MoveTo or MoveToParagraph is called with an invalid number of arguments.
- Root cause: The code supplies parameters that do not match any overload defined in the library.
- Preferred fix: Use the overload signatures documented for the current version, providing both paragraph index and character index for MoveToParagraph, or call MoveTo with a Node argument.
- Prevention: Review method signatures in the API docs and pass the required arguments exactly as specified.

### Applying paragraph style to a run

- Symptom: Runtime throws ArgumentException stating that the style is not a character style when assigning a StyleIdentifier to Font.StyleIdentifier.
- Root cause: A paragraph-level style (e.g., Heading1) is being applied to a Run, which only accepts character styles.
- Preferred fix: Assign a character style such as StyleIdentifier.DefaultParagraphFont to the run, or leave the run with the document's default character formatting.
- Prevention: Distinguish between paragraph and character styles; use Font.StyleIdentifier only for character styles.

### Assuming unsupported layout helpers

- Symptom: Compiler errors appear for calls to Paragraph.GetLinesCount or other invented layout APIs.
- Root cause: The code attempts to use methods that are not part of the Aspose.Words API for line-count or layout information.
- Preferred fix: Remove the unsupported call and, if line count is needed, approximate it by measuring the rendered height or by using a custom layout engine outside Aspose.Words.
- Prevention: Limit paragraph-related logic to properties and methods documented in the API reference.

### Over-broad paragraph count validation

- Symptom: Runtime failures occur when the test asserts the total number of paragraphs instead of the targeted subset.
- Root cause: The validation logic checks the entire document rather than the specific paragraph(s) affected by the operation.
- Preferred fix: Validate only the paragraph(s) that the task modifies or queries, using precise node references.
- Prevention: Scope assertions to the exact nodes involved in the operation.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Update the examples whenever a new Aspose.Words version introduces or deprecates API members, and verify all referenced members against the official documentation.
