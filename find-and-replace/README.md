# Find and Replace Examples for Aspose.Words for .NET

The Find and Replace category demonstrates how to perform literal, case-sensitive, whole-word, and regular-expression replacements in Word documents using Aspose.Words. Verified examples show correct usage of Document.Range.Replace, FindReplaceOptions, and IReplacingCallback with locally bootstrapped sample files.

## Snapshot

- Category: Find And Replace
- Slug: find-and-replace
- Total examples: 63
- Publish-ready successful examples: 63 / 63
- Source run: 20261102_031946_08fb5b
- Plain Replace examples: 41
- Formatting Aware Replace examples: 7
- Regex Replace examples: 7
- Callback Reporting examples: 6
- Input Bootstrap examples: 2

## Category rules that shaped these examples

- Do not reference System.Drawing; use Aspose.Drawing for any drawing-related types.
- Do not assume that source files, streams, or folders exist; generate minimal sample inputs within the example code.
- Do not call undocumented helpers such as Document.ReplaceAll or Paragraph.Replace.
- Use Document.Range.Replace with a literal string, a Regex, or a FindReplaceOptions instance for all supported replacement scenarios.
- Implement IReplacingCallback when custom replacement logic is required and assign it to FindReplaceOptions.ReplacingCallback.
- When a task mentions placeholders, headers, footers, or other sections, access the corresponding Nodes collection and apply the same replace pattern.
- Initialize all non-nullable reference-type fields or properties before use to avoid CS8618 warnings.
- Guard any value that may be null with an explicit null-check before dereferencing to prevent CS8600, CS8602, and CS8604.
- Declare local variables as nullable when a null value is legitimate and handle the null case explicitly.

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
dotnet add package Aspose.Drawing.Common --version 26.10.0
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
dotnet add package Aspose.Words --version 26.10.0

# PowerShell example
Copy-Item ..\find-and-replace\load-a-docx-file-into-a-document-object-and-replace-all-literal-string-occurrences.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-into-a-document-object-and-replace-all-literal-string-occurrences.cs` | Load a DOCX file into a Document object and replace all literal string occurrences. | plain-replace | docx | mcp |
| 2 | `perform-a-case-sensitive-replace-using-findreplaceoptions-with-matchcase-enabled-on-the-lo.cs` | Perform a case-sensitive replace using FindReplaceOptions with MatchCase enabled on the loaded document. | plain-replace | docx | mcp |
| 3 | `replace-whole-words-only-by-setting-findreplaceoptions-matchwholeword-to-true-during-repla.cs` | Replace whole words only by setting FindReplaceOptions.MatchWholeWord to true during replacement. | plain-replace | docx | mcp |
| 4 | `execute-a-regular-expression-replace-that-converts-dates-from-mm-dd-yyyy-to-yyyy-mm-dd-for.cs` | Execute a regular expression replace that converts dates from MM-DD-YYYY to YYYY-MM-DD format. | formatting-aware-replace | docx | mcp |
| 5 | `apply-metacharacters-in-the-replacement-string-to-insert-line-breaks-after-each-replaced-p.cs` | Apply metacharacters in the replacement string to insert line breaks after each replaced paragraph. | regex-replace | docx | mcp |
| 6 | `search-and-replace-text-within-header-sections-by-accessing-the-document-s-headersfooters.cs` | Search and replace text within header sections by accessing the document's HeadersFooters collection. | plain-replace | docx | mcp |
| 7 | `search-and-replace-text-within-footer-sections-while-preserving-existing-page-numbers-usin.cs` | Search and replace text within footer sections while preserving existing page numbers using FindReplaceOptions. | plain-replace | docx | mcp |
| 8 | `ignore-field-codes-during-replacement-by-enabling-findreplaceoptions-ignorefields-before-c.cs` | Ignore field codes during replacement by enabling FindReplaceOptions.IgnoreFields before calling Replace. | plain-replace | docx | mcp |
| 9 | `implement-a-custom-replaceevaluator-delegate-to-replace-only-words-longer-than-ten-charact.cs` | Implement a custom ReplaceEvaluator delegate to replace only words longer than ten characters. | plain-replace | docx | llm |
| 10 | `count-the-number-of-replacements-performed-by-storing-the-integer-result-returned-from-the.cs` | Count the number of replacements performed by storing the integer result returned from the Replace method. | plain-replace | docx | mcp |
| 11 | `chain-multiple-replace-calls-to-sequentially-replace-email-addresses-phone-numbers-and-url.cs` | Chain multiple Replace calls to sequentially replace email addresses, phone numbers, and URLs in the document. | plain-replace | docx | mcp |
| 12 | `combine-findreplaceoptions-with-documentbuilder-to-insert-dynamic-content-after-each-succe.cs` | Combine FindReplaceOptions with DocumentBuilder to insert dynamic content after each successful replacement. | plain-replace | docx | mcp |
| 13 | `enable-regex-mode-by-setting-findreplaceoptions-useregularexpressions-to-true-for-pattern.cs` | Enable regex mode by setting FindReplaceOptions.UseRegularExpressions to true for pattern-based replacements. | regex-replace | docx | mcp |
| 14 | `replace-tab-characters-with-four-spaces-using-a-regular-expression-pattern-matching-the-t.cs` | Replace tab characters with four spaces using a regular expression pattern matching the \t escape sequence. | regex-replace | docx | mcp |
| 15 | `replace-placeholder-tokens-like-name-with-actual-values-using-a-dictionary-inside-a-replac.cs` | Replace placeholder tokens like {{Name}} with actual values using a dictionary inside a ReplaceEvaluator. | plain-replace | docx | mcp |
| 16 | `execute-a-case-insensitive-replace-that-changes-all-instances-of-color-to-colour-throughou.cs` | Execute a case-insensitive replace that changes all instances of 'color' to 'colour' throughout the document. | formatting-aware-replace | docx | mcp |
| 17 | `apply-whole-word-matching-to-avoid-partial-replacements-inside-longer-words-such-as-catalo.cs` | Apply whole-word matching to avoid partial replacements inside longer words such as 'catalogue'. | callback-reporting | docx | mcp |
| 18 | `replace-multiple-line-break-sequences-with-a-single-paragraph-break-using-the-metacharacte.cs` | Replace multiple line break sequences with a single paragraph break using the metacharacter \n in the replacement. | regex-replace | docx | existing_repo |
| 19 | `perform-find-and-replace-on-a-document-loaded-from-a-memory-stream-without-writing-to-disk.cs` | Perform find and replace on a document loaded from a memory stream without writing to disk. | input-bootstrap | docx | mcp |
| 20 | `configure-findreplaceoptions-replacingcallback-to-log-each-replacement-occurrence-to-a-cus.cs` | Configure FindReplaceOptions.ReplacingCallback to log each replacement occurrence to a custom logger. | callback-reporting | docx | mcp |
| 21 | `use-a-regular-expression-to-locate-and-replace-all-html-tags-embedded-in-the-document-text.cs` | Use a regular expression to locate and replace all HTML tags embedded in the document text. | input-bootstrap | html | mcp |
| 22 | `update-the-copyright-year-in-headers-and-footers-by-replacing-the-old-year-with-the-curren.cs` | Update the copyright year in headers and footers by replacing the old year with the current year. | plain-replace | docx | mcp |
| 23 | `set-matchcase-false-and-matchwholeword-true-to-replace-whole-words-case-insensitively-acro.cs` | Set MatchCase false and MatchWholeWord true to replace whole words case-insensitively across the document. | plain-replace | docx | mcp |
| 24 | `implement-a-progress-reporter-that-receives-the-replacement-count-after-each-replace-call.cs` | Implement a progress reporter that receives the replacement count after each Replace call. | callback-reporting | docx | mcp |
| 25 | `replace-unicode-em-dashes-with-standard-hyphens-using-a-regular-expression-for-the-dash-ch.cs` | Replace Unicode em dashes with standard hyphens using a regular expression for the dash character. | plain-replace | docx | mcp |
| 26 | `mask-phone-numbers-by-finding-them-with-a-regular-expression-and-replacing-with-asterisks.cs` | Mask phone numbers by finding them with a regular expression and replacing with asterisks for privacy. | plain-replace | docx | mcp |
| 27 | `apply-findreplaceoptions-ignorefields-to-skip-replacing-field-result-texts-while-updating.cs` | Apply FindReplaceOptions.IgnoreFields to skip replacing field result texts while updating field codes. | plain-replace | docx | mcp |
| 28 | `add-a-prefix-to-each-matched-word-using-a-custom-replaceevaluator-delegate-during-replacem.cs` | Add a prefix to each matched word using a custom ReplaceEvaluator delegate during replacement. | plain-replace | docx | mcp |
| 29 | `replace-deprecated-terms-with-new-equivalents-across-multiple-languages-using-culture-spec.cs` | Replace deprecated terms with new equivalents across multiple languages using culture-specific regular expressions. | plain-replace | docx | mcp |
| 30 | `perform-conditional-replacement-that-only-changes-text-appearing-after-a-specific-heading.cs` | Perform conditional replacement that only changes text appearing after a specific heading using a callback. | callback-reporting | docx | mcp |
| 31 | `replace-double-spaces-with-a-single-space-throughout-the-document-using-a-simple-regular-e.cs` | Replace double spaces with a single space throughout the document using a simple regular expression. | plain-replace | docx | mcp |
| 32 | `add-a-suffix-to-each-matched-numeric-value-using-a-custom-replaceevaluator-during-the-repl.cs` | Add a suffix to each matched numeric value using a custom ReplaceEvaluator during the replace operation. | plain-replace | docx | mcp |
| 33 | `replace-language-code-tags-with-full-language-names-using-a-regular-expression-for-accurat.cs` | Replace language code tags with full language names using a regular expression for accurate localization. | plain-replace | docx | mcp |
| 34 | `replace-placeholder-merge-fields-with-actual-data-by-evaluating-each-field-inside-a-replac.cs` | Replace placeholder merge fields with actual data by evaluating each field inside a ReplaceEvaluator delegate. | plain-replace | docx | mcp |
| 35 | `replace-the-word-foo-only-when-it-appears-at-the-start-of-a-paragraph-using-a-custom-evalu.cs` | Replace the word 'foo' only when it appears at the start of a paragraph using a custom evaluator. | plain-replace | docx | llm |
| 36 | `replace-the-first-page-header-differently-from-other-headers-by-checking-section-propertie.cs` | Replace the first page header differently from other headers by checking section properties during replacement. | plain-replace | docx | mcp |
| 37 | `insert-a-page-break-after-each-replaced-heading-using-the-metacharacter-f-in-the-replaceme.cs` | Insert a page break after each replaced heading using the metacharacter \f in the replacement string. | regex-replace | docx | mcp |
| 38 | `replace-color-names-with-their-hexadecimal-representations-using-a-regular-expression-for.cs` | Replace color names with their hexadecimal representations using a regular expression for color conversion. | formatting-aware-replace | docx | mcp |
| 39 | `update-version-numbers-from-1-0-0-to-2-0-0-across-the-document-using-a-regular-expression.cs` | Update version numbers from 1.0.0 to 2.0.0 across the document using a regular expression. | plain-replace | docx | mcp |
| 40 | `replace-a-custom-delimiter-with-a-comma-while-preserving-surrounding-whitespace-using-a-re.cs` | Replace a custom delimiter with a comma while preserving surrounding whitespace using a regular expression. | plain-replace | docx | mcp |
| 41 | `mask-email-addresses-by-locating-them-with-a-regular-expression-and-replacing-with-a-maske.cs` | Mask email addresses by locating them with a regular expression and replacing with a masked version. | plain-replace | docx | mcp |
| 42 | `expand-macro-names-to-their-full-code-using-a-custom-replaceevaluator-delegate-during-repl.cs` | Expand macro names to their full code using a custom ReplaceEvaluator delegate during replacement. | plain-replace | docx | mcp |
| 43 | `convert-date-strings-from-january-1-2020-to-01-01-2020-using-a-regular-expression.cs` | Convert date strings from 'January 1, 2020' to '01/01/2020' using a regular expression. | plain-replace | docx | mcp |
| 44 | `replace-abbreviations-with-their-full-forms-while-preserving-surrounding-punctuation-using.cs` | Replace abbreviations with their full forms while preserving surrounding punctuation using a regular expression. | plain-replace | docx | mcp |
| 45 | `replace-placeholders-surrounded-by-double-brackets-with-values-from-a-json-configuration-f.cs` | Replace placeholders surrounded by double brackets with values from a JSON configuration file using a callback. | callback-reporting | json | mcp |
| 46 | `replace-text-in-headers-containing-a-specific-keyword-while-leaving-other-headers-unchange.cs` | Replace text in headers containing a specific keyword while leaving other headers unchanged. | plain-replace | docx | mcp |
| 47 | `convert-all-uppercase-words-to-title-case-using-a-custom-replaceevaluator-during-the-repla.cs` | Convert all uppercase words to title case using a custom ReplaceEvaluator during the replacement process. | plain-replace | docx | mcp |
| 48 | `replace-a-specific-bullet-character-with-a-different-bullet-style-throughout-the-document.cs` | Replace a specific bullet character with a different bullet style throughout the document using a regular expression. | formatting-aware-replace | docx | mcp |
| 49 | `replace-a-custom-tag-attribute-value-across-the-entire-document-using-a-regular-expression.cs` | Replace a custom tag attribute value across the entire document using a regular expression. | plain-replace | docx | mcp |
| 50 | `change-url-scheme-from-http-to-https-while-preserving-the-rest-of-each-hyperlink-using-a-r.cs` | Change URL scheme from http to https while preserving the rest of each hyperlink using a regular expression. | plain-replace | docx | mcp |
| 51 | `convert-markdown-style-headings-to-word-heading-styles-throughout-the-document-using-a-reg.cs` | Convert markdown style headings to Word heading styles throughout the document using a regular expression. | formatting-aware-replace | docx | mcp |
| 52 | `replace-the-copyright-symbol-with-the-unicode-character-using-a-regular-expression.cs` | Replace the copyright symbol with the Unicode © character using a regular expression. | plain-replace | docx | mcp |
| 53 | `insert-a-page-number-field-after-each-replaced-heading-using-documentbuilder-for-automatic.cs` | Insert a page number field after each replaced heading using DocumentBuilder for automatic numbering. | plain-replace | docx | mcp |
| 54 | `replace-numeric-formats-like-1-234-56-with-a-localized-format-using-a-regular-expression.cs` | Replace numeric formats like 1,234.56 with a localized format using a regular expression. | formatting-aware-replace | docx | mcp |
| 55 | `add-a-hyperlink-to-each-replaced-email-address-using-documentbuilder-during-the-replacemen.cs` | Add a hyperlink to each replaced email address using DocumentBuilder during the replacement operation. | plain-replace | docx | llm |
| 56 | `ignore-case-and-match-whole-words-when-updating-product-names-across-the-document-using-fi.cs` | Ignore case and match whole words when updating product names across the document using FindReplaceOptions. | plain-replace | docx | mcp |
| 57 | `replace-placeholders-surrounded-by-percent-signs-with-environment-variable-values-using-a.cs` | Replace placeholders surrounded by percent signs with environment variable values using a custom callback. | callback-reporting | docx | mcp |
| 58 | `replace-multiple-consecutive-spaces-with-a-single-space-throughout-the-document-using-a-re.cs` | Replace multiple consecutive spaces with a single space throughout the document using a regular expression. | plain-replace | docx | mcp |
| 59 | `add-a-custom-style-to-each-replaced-paragraph-using-documentbuilder-during-the-replacement.cs` | Add a custom style to each replaced paragraph using DocumentBuilder during the replacement process. | formatting-aware-replace | docx | mcp |
| 60 | `standardize-date-range-patterns-to-a-unified-format-across-the-document-using-a-regular-ex.cs` | Standardize date range patterns to a unified format across the document using a regular expression. | regex-replace | docx | mcp |
| 61 | `insert-a-table-of-figures-after-each-replaced-figure-caption-using-documentbuilder-for-aut.cs` | Insert a table of figures after each replaced figure caption using DocumentBuilder for automatic generation. | plain-replace | docx | mcp |
| 62 | `use-a-custom-replaceevaluator-to-replace-only-the-first-occurrence-of-a-pattern-in-each-se.cs` | Use a custom ReplaceEvaluator to replace only the first occurrence of a pattern in each section. | regex-replace | docx | mcp |
| 63 | `replace-a-placeholder-with-a-generated-qr-code-image-using-documentbuilder-during-the-repl.cs` | Replace a placeholder with a generated QR code image using DocumentBuilder during the replacement. | plain-replace | docx | llm |

## Common failure patterns seen during generation and how they were corrected

### Missing FindReplaceOptions property

- Symptom: The compiler reports that FindReplaceOptions does not contain definitions such as MatchWholeWord, UseRegularExpressions, or FindWholeWords.
- Root cause: The code references property names that were renamed or never existed in the Aspose.Words API version 26.10.0.
- Preferred fix: Use the correct property names: MatchCase for case sensitivity, FindWholeWordsOnly for whole-word matching, and rely on the Regex overload to enable regular-expression mode.
- Prevention: Consult the official Aspose.Words API reference for the exact property names and avoid hard-coding undocumented members.

### Undefined ReplaceEvaluator type

- Symptom: Compilation fails with CS0246 because the type ReplaceEvaluator cannot be found.
- Root cause: The example attempts to use a non-existent delegate type instead of the documented IReplacingCallback interface.
- Preferred fix: Implement a class that inherits from IReplacingCallback and assign an instance to FindReplaceOptions.ReplacingCallback.
- Prevention: Always use IReplacingCallback for custom replacement logic and verify its namespace (Aspose.Words.Replacing).

### Incorrect overload arguments

- Symptom: The compiler reports CS1503 errors when a Regex and a FindReplaceOptions object are passed to an overload expecting string arguments.
- Root cause: The code calls the string-based Replace overload instead of the Regex overload that accepts a Regex, a replacement string, and FindReplaceOptions.
- Preferred fix: Call Document.Range.Replace with the Regex overload: Replace(Regex pattern, string replacement, FindReplaceOptions options).
- Prevention: Match the argument types exactly to the overload signature; use IntelliSense or the API docs to select the correct method.

### Metacharacter replacement without real line breaks

- Symptom: Runtime throws an InvalidOperationException because no replacements were made when attempting to replace multiple line-break sequences with "\n" or "&p".
- Root cause: The source document contains plain text literals instead of actual ControlChar.LineBreak characters, so the regex never matches.
- Preferred fix: Insert real line-break characters using ControlChar.LineBreak when building the sample document, and match them with a regex such as "(\v){2,}".
- Prevention: Bootstrap the input document with the exact control characters that the replacement logic expects.

### Callback misuse leading to null parent

- Symptom: Runtime throws an InvalidOperationException or ArgumentException when the callback attempts to move the DocumentBuilder to a node with a null parent.
- Root cause: The IReplacingCallback implementation accesses args.MatchNode without verifying that it belongs to the main story or attempts to insert breaks outside a valid paragraph.
- Preferred fix: Validate that args.MatchNode is not null and that its parent story is the main document before performing insertions; use DocumentBuilder.MoveTo(args.MatchNode) only on nodes within the main story.
- Prevention: Add defensive null checks and story-type checks inside the Replacing method to ensure safe node manipulation.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Update the example code whenever a new Aspose.Words version changes property names or overload signatures, and run the verification suite after each upgrade.
