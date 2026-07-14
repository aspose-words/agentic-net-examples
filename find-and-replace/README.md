# Find and Replace Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Find and Replace category. Each file is a standalone console example selected from the verified 26.5.0 run.

## Snapshot

- Category: Find and Replace
- Slug: find-and-replace
- Total examples: 63
- Publish-ready successful examples: 63 / 63
- Source run: 20260619_131835_59df5f
- Callback Reporting examples: 6
- Formatting Aware Replace examples: 7
- Input Bootstrap examples: 2
- Plain Replace examples: 41
- Regex Replace examples: 7

## Category rules that shaped these examples

- Do not use System.Drawing in this category.
- Do not assume source documents or folders already exist; bootstrap them locally when needed.
- Do not invent unsupported find-replace APIs or callback hooks.
- Do not skip writing the requested output artifact when the task expects one.
- Use Range.Replace with the correct FindReplaceOptions or Regex only when the task requires it.
- Implement IReplacingCallback only through documented Aspose.Words APIs when custom replacement logic is required.
- Bootstrap local DOC, DOCX, HTML, TXT, stream, or folder inputs whenever the task implies an existing source.
- Use Newtonsoft.Json for JSON reporting tasks and Aspose.Drawing instead of System.Drawing when drawing-related types are needed.
- Initialize all non-nullable reference type properties to avoid CS8618 warnings.
- Avoid CS8600, CS8602, and CS8604 by guarding maybe-null values before dereference or assignment.
- Declare nullable locals when a value may legitimately be null and null-check before use.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.5.0
- Aspose.Drawing.Common 25.11.0
- Newtonsoft.Json

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0
dotnet add package Aspose.Drawing.Common --version 25.11.0
dotnet add package Newtonsoft.Json

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\find-and-replace\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `find-and-replace/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0
dotnet add package Aspose.Drawing.Common --version 25.11.0
dotnet add package Newtonsoft.Json

# PowerShell example
Copy-Item ..\find-and-replace\load-a-docx-file-into-a-document-object-and-replace-all-literal-string-occurrences.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-into-a-document-object-and-replace-all-literal-string-occurrences.cs` | Load a DOCX file into a Document object and replace all literal string occurrences. | Plain Replace | docx | mcp |
| 2 | `perform-a-case-sensitive-replace-using-findreplaceoptions-with-matchcase-enabled-on-the-lo.cs` | Perform a case-sensitive replace using FindReplaceOptions with MatchCase enabled on the loaded document. | Plain Replace | docx | mcp |
| 3 | `replace-whole-words-only-by-setting-findreplaceoptions-matchwholeword-to-true-during-repla.cs` | Replace whole words only by setting FindReplaceOptions.MatchWholeWord to true during replacement. | Plain Replace | docx | mcp |
| 4 | `execute-a-regular-expression-replace-that-converts-dates-from-mm-dd-yyyy-to-yyyy-mm-dd-for.cs` | Execute a regular expression replace that converts dates from MM-DD-YYYY to YYYY-MM-DD format. | Formatting Aware Replace | docx | mcp |
| 5 | `apply-metacharacters-in-the-replacement-string-to-insert-line-breaks-after-each-replaced-p.cs` | Apply metacharacters in the replacement string to insert line breaks after each replaced paragraph. | Regex Replace | docx | mcp |
| 6 | `search-and-replace-text-within-header-sections-by-accessing-the-document-s-headersfooters.cs` | Search and replace text within header sections by accessing the document's HeadersFooters collection. | Plain Replace | docx | mcp |
| 7 | `search-and-replace-text-within-footer-sections-while-preserving-existing-page-numbers-usin.cs` | Search and replace text within footer sections while preserving existing page numbers using FindReplaceOptions. | Plain Replace | docx | mcp |
| 8 | `ignore-field-codes-during-replacement-by-enabling-findreplaceoptions-ignorefields-before-c.cs` | Ignore field codes during replacement by enabling FindReplaceOptions.IgnoreFields before calling Replace. | Plain Replace | docx | mcp |
| 9 | `implement-a-custom-replaceevaluator-delegate-to-replace-only-words-longer-than-ten-charact.cs` | Implement a custom ReplaceEvaluator delegate to replace only words longer than ten characters. | Plain Replace | docx | mcp |
| 10 | `count-the-number-of-replacements-performed-by-storing-the-integer-result-returned-from-the.cs` | Count the number of replacements performed by storing the integer result returned from the Replace method. | Plain Replace | docx | mcp |
| 11 | `chain-multiple-replace-calls-to-sequentially-replace-email-addresses-phone-numbers-and-url.cs` | Chain multiple Replace calls to sequentially replace email addresses, phone numbers, and URLs in the document. | Plain Replace | docx | mcp |
| 12 | `combine-findreplaceoptions-with-documentbuilder-to-insert-dynamic-content-after-each-succe.cs` | Combine FindReplaceOptions with DocumentBuilder to insert dynamic content after each successful replacement. | Plain Replace | docx | mcp |
| 13 | `enable-regex-mode-by-setting-findreplaceoptions-useregularexpressions-to-true-for-pattern.cs` | Enable regex mode by setting FindReplaceOptions.UseRegularExpressions to true for pattern-based replacements. | Regex Replace | docx | mcp |
| 14 | `replace-tab-characters-with-four-spaces-using-a-regular-expression-pattern-matching-the-t.cs` | Replace tab characters with four spaces using a regular expression pattern matching the \t escape sequence. | Regex Replace | docx | mcp |
| 15 | `replace-placeholder-tokens-like-name-with-actual-values-using-a-dictionary-inside-a-replac.cs` | Replace placeholder tokens like {{Name}} with actual values using a dictionary inside a ReplaceEvaluator. | Plain Replace | docx | mcp |
| 16 | `execute-a-case-insensitive-replace-that-changes-all-instances-of-color-to-colour-throughou.cs` | Execute a case-insensitive replace that changes all instances of 'color' to 'colour' throughout the document. | Formatting Aware Replace | docx | mcp |
| 17 | `apply-whole-word-matching-to-avoid-partial-replacements-inside-longer-words-such-as-catalo.cs` | Apply whole-word matching to avoid partial replacements inside longer words such as 'catalogue'. | Callback Reporting | docx | mcp |
| 18 | `replace-multiple-line-break-sequences-with-a-single-paragraph-break-using-the-metacharacte.cs` | Replace multiple line break sequences with a single paragraph break using the metacharacter \n in the replacement. | Regex Replace | docx | existing_repo |
| 19 | `perform-find-and-replace-on-a-document-loaded-from-a-memory-stream-without-writing-to-disk.cs` | Perform find and replace on a document loaded from a memory stream without writing to disk. | Input Bootstrap | docx | mcp |
| 20 | `configure-findreplaceoptions-replacingcallback-to-log-each-replacement-occurrence-to-a-cus.cs` | Configure FindReplaceOptions.ReplacingCallback to log each replacement occurrence to a custom logger. | Callback Reporting | docx | mcp |
| 21 | `use-a-regular-expression-to-locate-and-replace-all-html-tags-embedded-in-the-document-text.cs` | Use a regular expression to locate and replace all HTML tags embedded in the document text. | Input Bootstrap | html | mcp |
| 22 | `update-the-copyright-year-in-headers-and-footers-by-replacing-the-old-year-with-the-curren.cs` | Update the copyright year in headers and footers by replacing the old year with the current year. | Plain Replace | docx | mcp |
| 23 | `set-matchcase-false-and-matchwholeword-true-to-replace-whole-words-case-insensitively-acro.cs` | Set MatchCase false and MatchWholeWord true to replace whole words case-insensitively across the document. | Plain Replace | docx | mcp |
| 24 | `implement-a-progress-reporter-that-receives-the-replacement-count-after-each-replace-call.cs` | Implement a progress reporter that receives the replacement count after each Replace call. | Callback Reporting | docx | mcp |
| 25 | `replace-unicode-em-dashes-with-standard-hyphens-using-a-regular-expression-for-the-dash-ch.cs` | Replace Unicode em dashes with standard hyphens using a regular expression for the dash character. | Plain Replace | docx | mcp |
| 26 | `mask-phone-numbers-by-finding-them-with-a-regular-expression-and-replacing-with-asterisks.cs` | Mask phone numbers by finding them with a regular expression and replacing with asterisks for privacy. | Plain Replace | docx | mcp |
| 27 | `apply-findreplaceoptions-ignorefields-to-skip-replacing-field-result-texts-while-updating.cs` | Apply FindReplaceOptions.IgnoreFields to skip replacing field result texts while updating field codes. | Plain Replace | docx | existing_repo |
| 28 | `add-a-prefix-to-each-matched-word-using-a-custom-replaceevaluator-delegate-during-replacem.cs` | Add a prefix to each matched word using a custom ReplaceEvaluator delegate during replacement. | Plain Replace | docx | mcp |
| 29 | `replace-deprecated-terms-with-new-equivalents-across-multiple-languages-using-culture-spec.cs` | Replace deprecated terms with new equivalents across multiple languages using culture-specific regular expressions. | Plain Replace | docx | mcp |
| 30 | `perform-conditional-replacement-that-only-changes-text-appearing-after-a-specific-heading.cs` | Perform conditional replacement that only changes text appearing after a specific heading using a callback. | Callback Reporting | docx | mcp |
| 31 | `replace-double-spaces-with-a-single-space-throughout-the-document-using-a-simple-regular-e.cs` | Replace double spaces with a single space throughout the document using a simple regular expression. | Plain Replace | docx | mcp |
| 32 | `add-a-suffix-to-each-matched-numeric-value-using-a-custom-replaceevaluator-during-the-repl.cs` | Add a suffix to each matched numeric value using a custom ReplaceEvaluator during the replace operation. | Plain Replace | docx | mcp |
| 33 | `replace-language-code-tags-with-full-language-names-using-a-regular-expression-for-accurat.cs` | Replace language code tags with full language names using a regular expression for accurate localization. | Plain Replace | docx | mcp |
| 34 | `replace-placeholder-merge-fields-with-actual-data-by-evaluating-each-field-inside-a-replac.cs` | Replace placeholder merge fields with actual data by evaluating each field inside a ReplaceEvaluator delegate. | Plain Replace | docx | mcp |
| 35 | `replace-the-word-foo-only-when-it-appears-at-the-start-of-a-paragraph-using-a-custom-evalu.cs` | Replace the word 'foo' only when it appears at the start of a paragraph using a custom evaluator. | Plain Replace | docx | mcp |
| 36 | `replace-the-first-page-header-differently-from-other-headers-by-checking-section-propertie.cs` | Replace the first page header differently from other headers by checking section properties during replacement. | Plain Replace | docx | mcp |
| 37 | `insert-a-page-break-after-each-replaced-heading-using-the-metacharacter-f-in-the-replaceme.cs` | Insert a page break after each replaced heading using the metacharacter \f in the replacement string. | Regex Replace | docx | mcp |
| 38 | `replace-color-names-with-their-hexadecimal-representations-using-a-regular-expression-for.cs` | Replace color names with their hexadecimal representations using a regular expression for color conversion. | Formatting Aware Replace | docx | mcp |
| 39 | `update-version-numbers-from-1-0-0-to-2-0-0-across-the-document-using-a-regular-expression.cs` | Update version numbers from 1.0.0 to 2.0.0 across the document using a regular expression. | Plain Replace | docx | mcp |
| 40 | `replace-a-custom-delimiter-with-a-comma-while-preserving-surrounding-whitespace-using-a-re.cs` | Replace a custom delimiter with a comma while preserving surrounding whitespace using a regular expression. | Plain Replace | docx | mcp |
| 41 | `mask-email-addresses-by-locating-them-with-a-regular-expression-and-replacing-with-a-maske.cs` | Mask email addresses by locating them with a regular expression and replacing with a masked version. | Plain Replace | docx | mcp |
| 42 | `expand-macro-names-to-their-full-code-using-a-custom-replaceevaluator-delegate-during-repl.cs` | Expand macro names to their full code using a custom ReplaceEvaluator delegate during replacement. | Plain Replace | docx | mcp |
| 43 | `convert-date-strings-from-january-1-2020-to-01-01-2020-using-a-regular-expression.cs` | Convert date strings from 'January 1, 2020' to '01/01/2020' using a regular expression. | Plain Replace | docx | mcp |
| 44 | `replace-abbreviations-with-their-full-forms-while-preserving-surrounding-punctuation-using.cs` | Replace abbreviations with their full forms while preserving surrounding punctuation using a regular expression. | Plain Replace | docx | mcp |
| 45 | `replace-placeholders-surrounded-by-double-brackets-with-values-from-a-json-configuration-f.cs` | Replace placeholders surrounded by double brackets with values from a JSON configuration file using a callback. | Callback Reporting | json | mcp |
| 46 | `replace-text-in-headers-containing-a-specific-keyword-while-leaving-other-headers-unchange.cs` | Replace text in headers containing a specific keyword while leaving other headers unchanged. | Plain Replace | docx | mcp |
| 47 | `convert-all-uppercase-words-to-title-case-using-a-custom-replaceevaluator-during-the-repla.cs` | Convert all uppercase words to title case using a custom ReplaceEvaluator during the replacement process. | Plain Replace | docx | mcp |
| 48 | `replace-a-specific-bullet-character-with-a-different-bullet-style-throughout-the-document.cs` | Replace a specific bullet character with a different bullet style throughout the document using a regular expression. | Formatting Aware Replace | docx | mcp |
| 49 | `replace-a-custom-tag-attribute-value-across-the-entire-document-using-a-regular-expression.cs` | Replace a custom tag attribute value across the entire document using a regular expression. | Plain Replace | docx | mcp |
| 50 | `change-url-scheme-from-http-to-https-while-preserving-the-rest-of-each-hyperlink-using-a-r.cs` | Change URL scheme from http to https while preserving the rest of each hyperlink using a regular expression. | Plain Replace | docx | mcp |
| 51 | `convert-markdown-style-headings-to-word-heading-styles-throughout-the-document-using-a-reg.cs` | Convert markdown style headings to Word heading styles throughout the document using a regular expression. | Formatting Aware Replace | docx | mcp |
| 52 | `replace-the-copyright-symbol-with-the-unicode-character-using-a-regular-expression.cs` | Replace the copyright symbol with the Unicode  character using a regular expression. | Plain Replace | docx | mcp |
| 53 | `insert-a-page-number-field-after-each-replaced-heading-using-documentbuilder-for-automatic.cs` | Insert a page number field after each replaced heading using DocumentBuilder for automatic numbering. | Plain Replace | docx | mcp |
| 54 | `replace-numeric-formats-like-1-234-56-with-a-localized-format-using-a-regular-expression.cs` | Replace numeric formats like 1,234.56 with a localized format using a regular expression. | Formatting Aware Replace | docx | mcp |
| 55 | `add-a-hyperlink-to-each-replaced-email-address-using-documentbuilder-during-the-replacemen.cs` | Add a hyperlink to each replaced email address using DocumentBuilder during the replacement operation. | Plain Replace | docx | existing_repo |
| 56 | `ignore-case-and-match-whole-words-when-updating-product-names-across-the-document-using-fi.cs` | Ignore case and match whole words when updating product names across the document using FindReplaceOptions. | Plain Replace | docx | mcp |
| 57 | `replace-placeholders-surrounded-by-percent-signs-with-environment-variable-values-using-a.cs` | Replace placeholders surrounded by percent signs with environment variable values using a custom callback. | Callback Reporting | docx | mcp |
| 58 | `replace-multiple-consecutive-spaces-with-a-single-space-throughout-the-document-using-a-re.cs` | Replace multiple consecutive spaces with a single space throughout the document using a regular expression. | Plain Replace | docx | mcp |
| 59 | `add-a-custom-style-to-each-replaced-paragraph-using-documentbuilder-during-the-replacement.cs` | Add a custom style to each replaced paragraph using DocumentBuilder during the replacement process. | Formatting Aware Replace | docx | mcp |
| 60 | `standardize-date-range-patterns-to-a-unified-format-across-the-document-using-a-regular-ex.cs` | Standardize date range patterns to a unified format across the document using a regular expression. | Regex Replace | docx | mcp |
| 61 | `insert-a-table-of-figures-after-each-replaced-figure-caption-using-documentbuilder-for-aut.cs` | Insert a table of figures after each replaced figure caption using DocumentBuilder for automatic generation. | Plain Replace | docx | mcp |
| 62 | `use-a-custom-replaceevaluator-to-replace-only-the-first-occurrence-of-a-pattern-in-each-se.cs` | Use a custom ReplaceEvaluator to replace only the first occurrence of a pattern in each section. | Regex Replace | docx | mcp |
| 63 | `replace-a-placeholder-with-a-generated-qr-code-image-using-documentbuilder-during-the-repl.cs` | Replace a placeholder with a generated QR code image using DocumentBuilder during the replacement. | Plain Replace | docx | existing_repo |

## Common failure patterns seen during generation and how they were corrected

### Invented find-replace APIs

- Symptom: Code tries to use unsupported helpers such as Document.ReplaceAll or Paragraph.Replace.
- Fix: Use Document.Range.Replace with supported FindReplaceOptions and callback APIs only.

### Callback misuse

- Symptom: Custom replacement logic fails because IReplacingCallback or ReplacingArgs are used incorrectly.
- Fix: Implement IReplacingCallback with the documented Replacing method and assign it through FindReplaceOptions.ReplacingCallback.

### Metacharacter source content mismatch

- Symptom: Line-break or metacharacter replacement finds no matches because the source document was built with plain literal text instead of real control structures.
- Fix: Create real line breaks or control characters in the source document and validate that Range.Replace returns a replacement count greater than zero.

### Invalid formatting insertion

- Symptom: Replacement logic inserts formatted content into invalid node parents or assumes every match belongs to a single Run.
- Fix: Use supported run, font, paragraph, or builder logic and preserve valid parent-child insertion structure.

### Drawing ambiguity

- Symptom: Compile errors caused by System.Drawing usage or ambiguous Font references.
- Fix: Use Aspose.Drawing only and prefer explicit Aspose.Drawing type names when ambiguity is possible.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.5.0 run.
- Preserve file-to-task traceability when updating this folder.
- Keep examples standalone and bootstrap local inputs inside the example whenever external sources are mentioned.
