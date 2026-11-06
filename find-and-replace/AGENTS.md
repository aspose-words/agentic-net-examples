---
name: find-and-replace
description: Verified C# examples for find-and-replace scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Find and Replace

## Purpose

Provide an AI coding assistant with clear, reproducible patterns for implementing robust find-and-replace operations in Aspose.Words, ensuring that all required namespaces are imported, documents are created or loaded locally, and only supported APIs are used.

## Non-negotiable conventions

- Do not reference System.Drawing; use Aspose.Drawing for any drawing-related types.
- Do not assume that source files, streams, or folders exist; generate minimal sample inputs within the example code.
- Do not call undocumented helpers such as Document.ReplaceAll or Paragraph.Replace.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Plain Replace workflow**: 41 examples
- **Formatting Aware Replace workflow**: 7 examples
- **Regex Replace workflow**: 7 examples
- **Callback Reporting workflow**: 6 examples
- **Input Bootstrap workflow**: 2 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Verify that the replacement count returned by Document.Range.Replace is greater than zero; otherwise throw an informative exception.

## File-to-task reference

- `load-a-docx-file-into-a-document-object-and-replace-all-literal-string-occurrences.cs`
  - Task: Load a DOCX file into a Document object and replace all literal string occurrences.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `perform-a-case-sensitive-replace-using-findreplaceoptions-with-matchcase-enabled-on-the-lo.cs`
  - Task: Perform a case-sensitive replace using FindReplaceOptions with MatchCase enabled on the loaded document.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-whole-words-only-by-setting-findreplaceoptions-matchwholeword-to-true-during-repla.cs`
  - Task: Replace whole words only by setting FindReplaceOptions.MatchWholeWord to true during replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `execute-a-regular-expression-replace-that-converts-dates-from-mm-dd-yyyy-to-yyyy-mm-dd-for.cs`
  - Task: Execute a regular expression replace that converts dates from MM-DD-YYYY to YYYY-MM-DD format.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: mcp
- `apply-metacharacters-in-the-replacement-string-to-insert-line-breaks-after-each-replaced-p.cs`
  - Task: Apply metacharacters in the replacement string to insert line breaks after each replaced paragraph.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: mcp
- `search-and-replace-text-within-header-sections-by-accessing-the-document-s-headersfooters.cs`
  - Task: Search and replace text within header sections by accessing the document's HeadersFooters collection.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `search-and-replace-text-within-footer-sections-while-preserving-existing-page-numbers-usin.cs`
  - Task: Search and replace text within footer sections while preserving existing page numbers using FindReplaceOptions.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `ignore-field-codes-during-replacement-by-enabling-findreplaceoptions-ignorefields-before-c.cs`
  - Task: Ignore field codes during replacement by enabling FindReplaceOptions.IgnoreFields before calling Replace.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `implement-a-custom-replaceevaluator-delegate-to-replace-only-words-longer-than-ten-charact.cs`
  - Task: Implement a custom ReplaceEvaluator delegate to replace only words longer than ten characters.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: llm
- `count-the-number-of-replacements-performed-by-storing-the-integer-result-returned-from-the.cs`
  - Task: Count the number of replacements performed by storing the integer result returned from the Replace method.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `chain-multiple-replace-calls-to-sequentially-replace-email-addresses-phone-numbers-and-url.cs`
  - Task: Chain multiple Replace calls to sequentially replace email addresses, phone numbers, and URLs in the document.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `combine-findreplaceoptions-with-documentbuilder-to-insert-dynamic-content-after-each-succe.cs`
  - Task: Combine FindReplaceOptions with DocumentBuilder to insert dynamic content after each successful replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `enable-regex-mode-by-setting-findreplaceoptions-useregularexpressions-to-true-for-pattern.cs`
  - Task: Enable regex mode by setting FindReplaceOptions.UseRegularExpressions to true for pattern-based replacements.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-tab-characters-with-four-spaces-using-a-regular-expression-pattern-matching-the-t.cs`
  - Task: Replace tab characters with four spaces using a regular expression pattern matching the \t escape sequence.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-placeholder-tokens-like-name-with-actual-values-using-a-dictionary-inside-a-replac.cs`
  - Task: Replace placeholder tokens like {{Name}} with actual values using a dictionary inside a ReplaceEvaluator.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `execute-a-case-insensitive-replace-that-changes-all-instances-of-color-to-colour-throughou.cs`
  - Task: Execute a case-insensitive replace that changes all instances of 'color' to 'colour' throughout the document.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: mcp
- `apply-whole-word-matching-to-avoid-partial-replacements-inside-longer-words-such-as-catalo.cs`
  - Task: Apply whole-word matching to avoid partial replacements inside longer words such as 'catalogue'.
  - Workflow: callback-reporting
  - Outputs: docx
  - Selected engine: mcp
- `replace-multiple-line-break-sequences-with-a-single-paragraph-break-using-the-metacharacte.cs`
  - Task: Replace multiple line break sequences with a single paragraph break using the metacharacter \n in the replacement.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: existing_repo
- `perform-find-and-replace-on-a-document-loaded-from-a-memory-stream-without-writing-to-disk.cs`
  - Task: Perform find and replace on a document loaded from a memory stream without writing to disk.
  - Workflow: input-bootstrap
  - Outputs: docx
  - Selected engine: mcp
- `configure-findreplaceoptions-replacingcallback-to-log-each-replacement-occurrence-to-a-cus.cs`
  - Task: Configure FindReplaceOptions.ReplacingCallback to log each replacement occurrence to a custom logger.
  - Workflow: callback-reporting
  - Outputs: docx
  - Selected engine: mcp
- `use-a-regular-expression-to-locate-and-replace-all-html-tags-embedded-in-the-document-text.cs`
  - Task: Use a regular expression to locate and replace all HTML tags embedded in the document text.
  - Workflow: input-bootstrap
  - Outputs: html
  - Selected engine: mcp
- `update-the-copyright-year-in-headers-and-footers-by-replacing-the-old-year-with-the-curren.cs`
  - Task: Update the copyright year in headers and footers by replacing the old year with the current year.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `set-matchcase-false-and-matchwholeword-true-to-replace-whole-words-case-insensitively-acro.cs`
  - Task: Set MatchCase false and MatchWholeWord true to replace whole words case-insensitively across the document.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `implement-a-progress-reporter-that-receives-the-replacement-count-after-each-replace-call.cs`
  - Task: Implement a progress reporter that receives the replacement count after each Replace call.
  - Workflow: callback-reporting
  - Outputs: docx
  - Selected engine: mcp
- `replace-unicode-em-dashes-with-standard-hyphens-using-a-regular-expression-for-the-dash-ch.cs`
  - Task: Replace Unicode em dashes with standard hyphens using a regular expression for the dash character.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `mask-phone-numbers-by-finding-them-with-a-regular-expression-and-replacing-with-asterisks.cs`
  - Task: Mask phone numbers by finding them with a regular expression and replacing with asterisks for privacy.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `apply-findreplaceoptions-ignorefields-to-skip-replacing-field-result-texts-while-updating.cs`
  - Task: Apply FindReplaceOptions.IgnoreFields to skip replacing field result texts while updating field codes.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `add-a-prefix-to-each-matched-word-using-a-custom-replaceevaluator-delegate-during-replacem.cs`
  - Task: Add a prefix to each matched word using a custom ReplaceEvaluator delegate during replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-deprecated-terms-with-new-equivalents-across-multiple-languages-using-culture-spec.cs`
  - Task: Replace deprecated terms with new equivalents across multiple languages using culture-specific regular expressions.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `perform-conditional-replacement-that-only-changes-text-appearing-after-a-specific-heading.cs`
  - Task: Perform conditional replacement that only changes text appearing after a specific heading using a callback.
  - Workflow: callback-reporting
  - Outputs: docx
  - Selected engine: mcp
- `replace-double-spaces-with-a-single-space-throughout-the-document-using-a-simple-regular-e.cs`
  - Task: Replace double spaces with a single space throughout the document using a simple regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `add-a-suffix-to-each-matched-numeric-value-using-a-custom-replaceevaluator-during-the-repl.cs`
  - Task: Add a suffix to each matched numeric value using a custom ReplaceEvaluator during the replace operation.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-language-code-tags-with-full-language-names-using-a-regular-expression-for-accurat.cs`
  - Task: Replace language code tags with full language names using a regular expression for accurate localization.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-placeholder-merge-fields-with-actual-data-by-evaluating-each-field-inside-a-replac.cs`
  - Task: Replace placeholder merge fields with actual data by evaluating each field inside a ReplaceEvaluator delegate.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-the-word-foo-only-when-it-appears-at-the-start-of-a-paragraph-using-a-custom-evalu.cs`
  - Task: Replace the word 'foo' only when it appears at the start of a paragraph using a custom evaluator.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: llm
- `replace-the-first-page-header-differently-from-other-headers-by-checking-section-propertie.cs`
  - Task: Replace the first page header differently from other headers by checking section properties during replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-page-break-after-each-replaced-heading-using-the-metacharacter-f-in-the-replaceme.cs`
  - Task: Insert a page break after each replaced heading using the metacharacter \f in the replacement string.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-color-names-with-their-hexadecimal-representations-using-a-regular-expression-for.cs`
  - Task: Replace color names with their hexadecimal representations using a regular expression for color conversion.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: mcp
- `update-version-numbers-from-1-0-0-to-2-0-0-across-the-document-using-a-regular-expression.cs`
  - Task: Update version numbers from 1.0.0 to 2.0.0 across the document using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-a-custom-delimiter-with-a-comma-while-preserving-surrounding-whitespace-using-a-re.cs`
  - Task: Replace a custom delimiter with a comma while preserving surrounding whitespace using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `mask-email-addresses-by-locating-them-with-a-regular-expression-and-replacing-with-a-maske.cs`
  - Task: Mask email addresses by locating them with a regular expression and replacing with a masked version.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `expand-macro-names-to-their-full-code-using-a-custom-replaceevaluator-delegate-during-repl.cs`
  - Task: Expand macro names to their full code using a custom ReplaceEvaluator delegate during replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `convert-date-strings-from-january-1-2020-to-01-01-2020-using-a-regular-expression.cs`
  - Task: Convert date strings from 'January 1, 2020' to '01/01/2020' using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-abbreviations-with-their-full-forms-while-preserving-surrounding-punctuation-using.cs`
  - Task: Replace abbreviations with their full forms while preserving surrounding punctuation using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-placeholders-surrounded-by-double-brackets-with-values-from-a-json-configuration-f.cs`
  - Task: Replace placeholders surrounded by double brackets with values from a JSON configuration file using a callback.
  - Workflow: callback-reporting
  - Outputs: json
  - Selected engine: mcp
- `replace-text-in-headers-containing-a-specific-keyword-while-leaving-other-headers-unchange.cs`
  - Task: Replace text in headers containing a specific keyword while leaving other headers unchanged.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `convert-all-uppercase-words-to-title-case-using-a-custom-replaceevaluator-during-the-repla.cs`
  - Task: Convert all uppercase words to title case using a custom ReplaceEvaluator during the replacement process.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-a-specific-bullet-character-with-a-different-bullet-style-throughout-the-document.cs`
  - Task: Replace a specific bullet character with a different bullet style throughout the document using a regular expression.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-a-custom-tag-attribute-value-across-the-entire-document-using-a-regular-expression.cs`
  - Task: Replace a custom tag attribute value across the entire document using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `change-url-scheme-from-http-to-https-while-preserving-the-rest-of-each-hyperlink-using-a-r.cs`
  - Task: Change URL scheme from http to https while preserving the rest of each hyperlink using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `convert-markdown-style-headings-to-word-heading-styles-throughout-the-document-using-a-reg.cs`
  - Task: Convert markdown style headings to Word heading styles throughout the document using a regular expression.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-the-copyright-symbol-with-the-unicode-character-using-a-regular-expression.cs`
  - Task: Replace the copyright symbol with the Unicode © character using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-page-number-field-after-each-replaced-heading-using-documentbuilder-for-automatic.cs`
  - Task: Insert a page number field after each replaced heading using DocumentBuilder for automatic numbering.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-numeric-formats-like-1-234-56-with-a-localized-format-using-a-regular-expression.cs`
  - Task: Replace numeric formats like 1,234.56 with a localized format using a regular expression.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: mcp
- `add-a-hyperlink-to-each-replaced-email-address-using-documentbuilder-during-the-replacemen.cs`
  - Task: Add a hyperlink to each replaced email address using DocumentBuilder during the replacement operation.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: llm
- `ignore-case-and-match-whole-words-when-updating-product-names-across-the-document-using-fi.cs`
  - Task: Ignore case and match whole words when updating product names across the document using FindReplaceOptions.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-placeholders-surrounded-by-percent-signs-with-environment-variable-values-using-a.cs`
  - Task: Replace placeholders surrounded by percent signs with environment variable values using a custom callback.
  - Workflow: callback-reporting
  - Outputs: docx
  - Selected engine: mcp
- `replace-multiple-consecutive-spaces-with-a-single-space-throughout-the-document-using-a-re.cs`
  - Task: Replace multiple consecutive spaces with a single space throughout the document using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `add-a-custom-style-to-each-replaced-paragraph-using-documentbuilder-during-the-replacement.cs`
  - Task: Add a custom style to each replaced paragraph using DocumentBuilder during the replacement process.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: mcp
- `standardize-date-range-patterns-to-a-unified-format-across-the-document-using-a-regular-ex.cs`
  - Task: Standardize date range patterns to a unified format across the document using a regular expression.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-table-of-figures-after-each-replaced-figure-caption-using-documentbuilder-for-aut.cs`
  - Task: Insert a table of figures after each replaced figure caption using DocumentBuilder for automatic generation.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
- `use-a-custom-replaceevaluator-to-replace-only-the-first-occurrence-of-a-pattern-in-each-se.cs`
  - Task: Use a custom ReplaceEvaluator to replace only the first occurrence of a pattern in each section.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-a-placeholder-with-a-generated-qr-code-image-using-documentbuilder-during-the-repl.cs`
  - Task: Replace a placeholder with a generated QR code image using DocumentBuilder during the replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: llm

## Common failure patterns and preferred agent fixes

- **Missing FindReplaceOptions property**
  - Symptom: The compiler reports that FindReplaceOptions does not contain definitions such as MatchWholeWord, UseRegularExpressions, or FindWholeWords.
  - Root cause: The code references property names that were renamed or never existed in the Aspose.Words API version 26.10.0.
  - Preferred fix: Use the correct property names: MatchCase for case sensitivity, FindWholeWordsOnly for whole-word matching, and rely on the Regex overload to enable regular-expression mode.
  - Prevention: Consult the official Aspose.Words API reference for the exact property names and avoid hard-coding undocumented members.

- **Undefined ReplaceEvaluator type**
  - Symptom: Compilation fails with CS0246 because the type ReplaceEvaluator cannot be found.
  - Root cause: The example attempts to use a non-existent delegate type instead of the documented IReplacingCallback interface.
  - Preferred fix: Implement a class that inherits from IReplacingCallback and assign an instance to FindReplaceOptions.ReplacingCallback.
  - Prevention: Always use IReplacingCallback for custom replacement logic and verify its namespace (Aspose.Words.Replacing).

- **Incorrect overload arguments**
  - Symptom: The compiler reports CS1503 errors when a Regex and a FindReplaceOptions object are passed to an overload expecting string arguments.
  - Root cause: The code calls the string-based Replace overload instead of the Regex overload that accepts a Regex, a replacement string, and FindReplaceOptions.
  - Preferred fix: Call Document.Range.Replace with the Regex overload: Replace(Regex pattern, string replacement, FindReplaceOptions options).
  - Prevention: Match the argument types exactly to the overload signature; use IntelliSense or the API docs to select the correct method.

- **Metacharacter replacement without real line breaks**
  - Symptom: Runtime throws an InvalidOperationException because no replacements were made when attempting to replace multiple line-break sequences with "\n" or "&p".
  - Root cause: The source document contains plain text literals instead of actual ControlChar.LineBreak characters, so the regex never matches.
  - Preferred fix: Insert real line-break characters using ControlChar.LineBreak when building the sample document, and match them with a regex such as "(\v){2,}".
  - Prevention: Bootstrap the input document with the exact control characters that the replacement logic expects.

- **Callback misuse leading to null parent**
  - Symptom: Runtime throws an InvalidOperationException or ArgumentException when the callback attempts to move the DocumentBuilder to a node with a null parent.
  - Root cause: The IReplacingCallback implementation accesses args.MatchNode without verifying that it belongs to the main story or attempts to insert breaks outside a valid paragraph.
  - Preferred fix: Validate that args.MatchNode is not null and that its parent story is the main document before performing insertions; use DocumentBuilder.MoveTo(args.MatchNode) only on nodes within the main story.
  - Prevention: Add defensive null checks and story-type checks inside the Replacing method to ensure safe node manipulation.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.10.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
dotnet add package Aspose.Drawing.Common --version 26.10.0
dotnet add package Newtonsoft.Json
```

```powershell
Copy-Item ..\find-and-replace\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update the example code whenever a new Aspose.Words version changes property names or overload signatures, and run the verification suite after each upgrade.
