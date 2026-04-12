---
name: find-and-replace
description: Verified C# examples for find-and-replace scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Find and Replace

## Purpose

This folder is a **live, curated example set** for find-and-replace scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free search, replacement, regex, callback, and replacement-report workflows using direct Aspose.Words APIs.

## Non-negotiable conventions

- Use native `Range.Replace` APIs directly.
- Bootstrap local source documents, files, streams, or folders whenever the task implies an existing source.
- Use `FindReplaceOptions` and `IReplacingCallback` only through documented Aspose.Words APIs.
- Use `Newtonsoft.Json` for JSON reporting tasks and `Aspose.Drawing` instead of `System.Drawing` when drawing-related types are needed.
- For metacharacter tasks, create real document control structures rather than plain literal backslash sequences.
- Guard maybe-null values to avoid nullable-reference warnings such as `CS8600`, `CS8602`, and `CS8604`.

## Recommended workflow selection

- **Plain replace workflow**: 41 examples
- **Regex replace workflow**: 7 examples
- **Formatting-aware replace workflow**: 7 examples
- **Callback / reporting workflow**: 6 examples
- **Input-bootstrap workflow**: 2 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. Source documents, files, streams, or folders must be bootstrapped locally whenever the task implies an existing input.
3. Replacements must be performed through real Aspose.Words processing flow and validated when the task expects changes.
4. Requested report or export files must actually be written.
5. Drawing-related types must use `Aspose.Drawing` and not `System.Drawing`.

## File-to-task reference

- `load-a-docx-file-into-a-document-object-and-replace-all-literal-string-occurrences.cs`
  - Task: Load a DOCX file into a Document object and replace all literal string occurrences.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `perform-a-case-sensitive-replace-using-findreplaceoptions-with-matchcase-enabled-on-the-lo.cs`
  - Task: Perform a case‑sensitive replace using FindReplaceOptions with MatchCase enabled on the loaded document.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-whole-words-only-by-setting-findreplaceoptions-matchwholeword-to-true-during-repla.cs`
  - Task: Replace whole words only by setting FindReplaceOptions.MatchWholeWord to true during replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `execute-a-regular-expression-replace-that-converts-dates-from-mm-dd-yyyy-to-yyyy-mm-dd-for.cs`
  - Task: Execute a regular expression replace that converts dates from MM‑DD‑YYYY to YYYY‑MM‑DD format.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: verified
- `apply-metacharacters-in-the-replacement-string-to-insert-line-breaks-after-each-replaced-p.cs`
  - Task: Apply metacharacters in the replacement string to insert line breaks after each replaced paragraph.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: verified
- `search-and-replace-text-within-header-sections-by-accessing-the-document-s-headersfooters.cs`
  - Task: Search and replace text within header sections by accessing the document's HeadersFooters collection.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `search-and-replace-text-within-footer-sections-while-preserving-existing-page-numbers-usin.cs`
  - Task: Search and replace text within footer sections while preserving existing page numbers using FindReplaceOptions.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `ignore-field-codes-during-replacement-by-enabling-findreplaceoptions-ignorefields-before-c.cs`
  - Task: Ignore field codes during replacement by enabling FindReplaceOptions.IgnoreFields before calling Replace.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `implement-a-custom-replaceevaluator-delegate-to-replace-only-words-longer-than-ten-charact.cs`
  - Task: Implement a custom ReplaceEvaluator delegate to replace only words longer than ten characters.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `count-the-number-of-replacements-performed-by-storing-the-integer-result-returned-from-the.cs`
  - Task: Count the number of replacements performed by storing the integer result returned from the Replace method.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `chain-multiple-replace-calls-to-sequentially-replace-email-addresses-phone-numbers-and-url.cs`
  - Task: Chain multiple Replace calls to sequentially replace email addresses, phone numbers, and URLs in the document.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `combine-findreplaceoptions-with-documentbuilder-to-insert-dynamic-content-after-each-succe.cs`
  - Task: Combine FindReplaceOptions with DocumentBuilder to insert dynamic content after each successful replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `enable-regex-mode-by-setting-findreplaceoptions-useregularexpressions-to-true-for-pattern.cs`
  - Task: Enable regex mode by setting FindReplaceOptions.UseRegularExpressions to true for pattern‑based replacements.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-tab-characters-with-four-spaces-using-a-regular-expression-pattern-matching-the-t.cs`
  - Task: Replace tab characters with four spaces using a regular expression pattern matching the \t escape sequence.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-placeholder-tokens-like-name-with-actual-values-using-a-dictionary-inside-a-replac.cs`
  - Task: Replace placeholder tokens like {{Name}} with actual values using a dictionary inside a ReplaceEvaluator.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `execute-a-case-insensitive-replace-that-changes-all-instances-of-color-to-colour-throughou.cs`
  - Task: Execute a case‑insensitive replace that changes all instances of ‘color’ to ‘colour’ throughout the document.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: verified
- `apply-whole-word-matching-to-avoid-partial-replacements-inside-longer-words-such-as-catalo.cs`
  - Task: Apply whole‑word matching to avoid partial replacements inside longer words such as ‘catalogue’.
  - Workflow: callback-reporting
  - Outputs: docx
  - Selected engine: verified
- `replace-multiple-line-break-sequences-with-a-single-paragraph-break-using-the-metacharacte.cs`
  - Task: Replace multiple line break sequences with a single paragraph break using the metacharacter \n in the replacement.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: verified
- `perform-find-and-replace-on-a-document-loaded-from-a-memory-stream-without-writing-to-disk.cs`
  - Task: Perform find and replace on a document loaded from a memory stream without writing to disk.
  - Workflow: input-bootstrap
  - Outputs: docx
  - Selected engine: verified
- `configure-findreplaceoptions-replacingcallback-to-log-each-replacement-occurrence-to-a-cus.cs`
  - Task: Configure FindReplaceOptions.ReplacingCallback to log each replacement occurrence to a custom logger.
  - Workflow: callback-reporting
  - Outputs: docx
  - Selected engine: verified
- `use-a-regular-expression-to-locate-and-replace-all-html-tags-embedded-in-the-document-text.cs`
  - Task: Use a regular expression to locate and replace all HTML tags embedded in the document text.
  - Workflow: input-bootstrap
  - Outputs: html
  - Selected engine: verified
- `update-the-copyright-year-in-headers-and-footers-by-replacing-the-old-year-with-the-curren.cs`
  - Task: Update the copyright year in headers and footers by replacing the old year with the current year.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `set-matchcase-false-and-matchwholeword-true-to-replace-whole-words-case-insensitively-acro.cs`
  - Task: Set MatchCase false and MatchWholeWord true to replace whole words case‑insensitively across the document.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `implement-a-progress-reporter-that-receives-the-replacement-count-after-each-replace-call.cs`
  - Task: Implement a progress reporter that receives the replacement count after each Replace call.
  - Workflow: callback-reporting
  - Outputs: docx
  - Selected engine: verified
- `replace-unicode-em-dashes-with-standard-hyphens-using-a-regular-expression-for-the-dash-ch.cs`
  - Task: Replace Unicode em dashes with standard hyphens using a regular expression for the dash character.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `mask-phone-numbers-by-finding-them-with-a-regular-expression-and-replacing-with-asterisks.cs`
  - Task: Mask phone numbers by finding them with a regular expression and replacing with asterisks for privacy.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `apply-findreplaceoptions-ignorefields-to-skip-replacing-field-result-texts-while-updating.cs`
  - Task: Apply FindReplaceOptions.IgnoreFields to skip replacing field result texts while updating field codes.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `add-a-prefix-to-each-matched-word-using-a-custom-replaceevaluator-delegate-during-replacem.cs`
  - Task: Add a prefix to each matched word using a custom ReplaceEvaluator delegate during replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-deprecated-terms-with-new-equivalents-across-multiple-languages-using-culture-spec.cs`
  - Task: Replace deprecated terms with new equivalents across multiple languages using culture‑specific regular expressions.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `perform-conditional-replacement-that-only-changes-text-appearing-after-a-specific-heading.cs`
  - Task: Perform conditional replacement that only changes text appearing after a specific heading using a callback.
  - Workflow: callback-reporting
  - Outputs: docx
  - Selected engine: verified
- `replace-double-spaces-with-a-single-space-throughout-the-document-using-a-simple-regular-e.cs`
  - Task: Replace double spaces with a single space throughout the document using a simple regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `add-a-suffix-to-each-matched-numeric-value-using-a-custom-replaceevaluator-during-the-repl.cs`
  - Task: Add a suffix to each matched numeric value using a custom ReplaceEvaluator during the replace operation.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-language-code-tags-with-full-language-names-using-a-regular-expression-for-accurat.cs`
  - Task: Replace language code tags with full language names using a regular expression for accurate localization.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-placeholder-merge-fields-with-actual-data-by-evaluating-each-field-inside-a-replac.cs`
  - Task: Replace placeholder merge fields with actual data by evaluating each field inside a ReplaceEvaluator delegate.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-the-word-foo-only-when-it-appears-at-the-start-of-a-paragraph-using-a-custom-evalu.cs`
  - Task: Replace the word ‘foo’ only when it appears at the start of a paragraph using a custom evaluator.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-the-first-page-header-differently-from-other-headers-by-checking-section-propertie.cs`
  - Task: Replace the first page header differently from other headers by checking section properties during replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `insert-a-page-break-after-each-replaced-heading-using-the-metacharacter-f-in-the-replaceme.cs`
  - Task: Insert a page break after each replaced heading using the metacharacter \f in the replacement string.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-color-names-with-their-hexadecimal-representations-using-a-regular-expression-for.cs`
  - Task: Replace color names with their hexadecimal representations using a regular expression for color conversion.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: verified
- `update-version-numbers-from-1-0-0-to-2-0-0-across-the-document-using-a-regular-expression.cs`
  - Task: Update version numbers from 1.0.0 to 2.0.0 across the document using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-a-custom-delimiter-with-a-comma-while-preserving-surrounding-whitespace-using-a-re.cs`
  - Task: Replace a custom delimiter with a comma while preserving surrounding whitespace using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `mask-email-addresses-by-locating-them-with-a-regular-expression-and-replacing-with-a-maske.cs`
  - Task: Mask email addresses by locating them with a regular expression and replacing with a masked version.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `expand-macro-names-to-their-full-code-using-a-custom-replaceevaluator-delegate-during-repl.cs`
  - Task: Expand macro names to their full code using a custom ReplaceEvaluator delegate during replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `convert-date-strings-from-january-1-2020-to-01-01-2020-using-a-regular-expression.cs`
  - Task: Convert date strings from ‘January 1, 2020’ to ‘01/01/2020’ using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-abbreviations-with-their-full-forms-while-preserving-surrounding-punctuation-using.cs`
  - Task: Replace abbreviations with their full forms while preserving surrounding punctuation using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-placeholders-surrounded-by-double-brackets-with-values-from-a-json-configuration-f.cs`
  - Task: Replace placeholders surrounded by double brackets with values from a JSON configuration file using a callback.
  - Workflow: callback-reporting
  - Outputs: json
  - Selected engine: verified
- `replace-text-in-headers-containing-a-specific-keyword-while-leaving-other-headers-unchange.cs`
  - Task: Replace text in headers containing a specific keyword while leaving other headers unchanged.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `convert-all-uppercase-words-to-title-case-using-a-custom-replaceevaluator-during-the-repla.cs`
  - Task: Convert all uppercase words to title case using a custom ReplaceEvaluator during the replacement process.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-a-specific-bullet-character-with-a-different-bullet-style-throughout-the-document.cs`
  - Task: Replace a specific bullet character with a different bullet style throughout the document using a regular expression.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-a-custom-tag-attribute-value-across-the-entire-document-using-a-regular-expression.cs`
  - Task: Replace a custom tag attribute value across the entire document using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `change-url-scheme-from-http-to-https-while-preserving-the-rest-of-each-hyperlink-using-a-r.cs`
  - Task: Change URL scheme from http to https while preserving the rest of each hyperlink using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `convert-markdown-style-headings-to-word-heading-styles-throughout-the-document-using-a-reg.cs`
  - Task: Convert markdown style headings to Word heading styles throughout the document using a regular expression.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-the-copyright-symbol-with-the-unicode-character-using-a-regular-expression.cs`
  - Task: Replace the copyright symbol with the Unicode © character using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `insert-a-page-number-field-after-each-replaced-heading-using-documentbuilder-for-automatic.cs`
  - Task: Insert a page number field after each replaced heading using DocumentBuilder for automatic numbering.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-numeric-formats-like-1-234-56-with-a-localized-format-using-a-regular-expression.cs`
  - Task: Replace numeric formats like 1,234.56 with a localized format using a regular expression.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: verified
- `add-a-hyperlink-to-each-replaced-email-address-using-documentbuilder-during-the-replacemen.cs`
  - Task: Add a hyperlink to each replaced email address using DocumentBuilder during the replacement operation.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `ignore-case-and-match-whole-words-when-updating-product-names-across-the-document-using-fi.cs`
  - Task: Ignore case and match whole words when updating product names across the document using FindReplaceOptions.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-placeholders-surrounded-by-percent-signs-with-environment-variable-values-using-a.cs`
  - Task: Replace placeholders surrounded by percent signs with environment variable values using a custom callback.
  - Workflow: callback-reporting
  - Outputs: docx
  - Selected engine: verified
- `replace-multiple-consecutive-spaces-with-a-single-space-throughout-the-document-using-a-re.cs`
  - Task: Replace multiple consecutive spaces with a single space throughout the document using a regular expression.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `add-a-custom-style-to-each-replaced-paragraph-using-documentbuilder-during-the-replacement.cs`
  - Task: Add a custom style to each replaced paragraph using DocumentBuilder during the replacement process.
  - Workflow: formatting-aware-replace
  - Outputs: docx
  - Selected engine: verified
- `standardize-date-range-patterns-to-a-unified-format-across-the-document-using-a-regular-ex.cs`
  - Task: Standardize date range patterns to a unified format across the document using a regular expression.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: verified
- `insert-a-table-of-figures-after-each-replaced-figure-caption-using-documentbuilder-for-aut.cs`
  - Task: Insert a table of figures after each replaced figure caption using DocumentBuilder for automatic generation.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified
- `use-a-custom-replaceevaluator-to-replace-only-the-first-occurrence-of-a-pattern-in-each-se.cs`
  - Task: Use a custom ReplaceEvaluator to replace only the first occurrence of a pattern in each section.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: verified
- `replace-a-placeholder-with-a-generated-qr-code-image-using-documentbuilder-during-the-repl.cs`
  - Task: Replace a placeholder with a generated QR code image using DocumentBuilder during the replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: verified

## Common failure patterns and preferred agent fixes

- **Invented find-replace APIs**
  - Symptom: Code tries to use unsupported helpers such as Document.ReplaceAll or Paragraph.Replace.
  - Preferred fix: Use Document.Range.Replace with supported FindReplaceOptions and callback APIs only.

- **Callback misuse**
  - Symptom: Custom replacement logic fails because IReplacingCallback or ReplacingArgs are used incorrectly.
  - Preferred fix: Implement IReplacingCallback with the documented Replacing method and assign it through FindReplaceOptions.ReplacingCallback.

- **Metacharacter source content mismatch**
  - Symptom: Line-break or metacharacter replacement finds no matches because the source document was built with plain literal text instead of real control structures.
  - Preferred fix: Create real line breaks or control characters in the source document and validate that Range.Replace returns a replacement count greater than zero.

- **Invalid formatting insertion**
  - Symptom: Replacement logic inserts formatted content into invalid node parents or assumes every match belongs to a single Run.
  - Preferred fix: Use supported run, font, paragraph, or builder logic and preserve valid parent-child insertion structure.

- **Drawing ambiguity**
  - Symptom: Compile errors caused by System.Drawing usage or ambiguous Font references.
  - Preferred fix: Use Aspose.Drawing only and prefer explicit Aspose.Drawing type names when ambiguity is possible.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`
- Additional package: `Aspose.Drawing.Common`
- Additional package: `Newtonsoft.Json`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.Drawing.Common
dotnet add package Newtonsoft.Json
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\find-and-replace\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer direct Aspose.Words find-and-replace APIs over speculative shortcuts.
