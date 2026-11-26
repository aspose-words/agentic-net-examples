---
name: find-and-replace
description: Verified C# examples for find-and-replace scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Find and Replace

## Purpose

Provide an AI coding agent with concise guidance to implement reliable find-and-replace operations in Aspose.Words, covering the correct API usage, required namespaces, and common pitfalls to avoid compilation and runtime errors.

## Non-negotiable conventions

- Do not reference System.Drawing; use Aspose.Drawing for any drawing-related types.
- Do not assume source files or folders exist; always bootstrap them within the example.
- Do not invent unsupported find-replace helpers such as Document.ReplaceAll or Paragraph.Replace.
- Do not omit the final Save call when the task expects an output artifact.
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

1. Validate that Document.Range.Replace returns a replacement count greater than zero before proceeding to save or report.

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
  - Selected engine: llm
- `implement-a-custom-replaceevaluator-delegate-to-replace-only-words-longer-than-ten-charact.cs`
  - Task: Implement a custom ReplaceEvaluator delegate to replace only words longer than ten characters.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: mcp
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
  - Selected engine: llm
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
  - Selected engine: llm
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
  - Selected engine: mcp
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
  - Selected engine: llm
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
  - Selected engine: llm
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
  - Selected engine: mcp
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
  - Selected engine: llm
- `use-a-custom-replaceevaluator-to-replace-only-the-first-occurrence-of-a-pattern-in-each-se.cs`
  - Task: Use a custom ReplaceEvaluator to replace only the first occurrence of a pattern in each section.
  - Workflow: regex-replace
  - Outputs: docx
  - Selected engine: mcp
- `replace-a-placeholder-with-a-generated-qr-code-image-using-documentbuilder-during-the-repl.cs`
  - Task: Replace a placeholder with a generated QR code image using DocumentBuilder during the replacement.
  - Workflow: plain-replace
  - Outputs: docx
  - Selected engine: existing_repo

## Common failure patterns and preferred agent fixes

- **Incorrect FindReplaceOptions property**
  - Symptom: Compilation fails with CS0117 because properties such as MatchWholeWord, UseRegularExpressions, or FindWholeWords are not found on FindReplaceOptions.
  - Root cause: The code references outdated or misspelled property names that do not exist in the current Aspose.Words API.
  - Preferred fix: Use the correct property names: FindWholeWordsOnly, MatchCase, IgnoreFields, etc., as defined in Aspose.Words.Replacing.FindReplaceOptions.
  - Prevention: Consult the official Aspose.Words documentation for the exact property names and enable IntelliSense to catch misspellings.

- **Missing FindReplaceOptions type**
  - Symptom: Compilation fails with CS0246 indicating that FindReplaceOptions cannot be found.
  - Root cause: The required namespace Aspose.Words.Replacing is not imported, so the type is unavailable.
  - Preferred fix: Add `using Aspose.Words.Replacing;` at the top of the file.
  - Prevention: Always include the Replacing namespace whenever FindReplaceOptions or related types are used.

- **Missing ReplaceEvaluator type**
  - Symptom: Compilation fails with CS0246 because ReplaceEvaluator is undefined.
  - Root cause: The code attempts to use a non-existent delegate name; the correct interface is IReplacingCallback.
  - Preferred fix: Implement IReplacingCallback and assign the instance to FindReplaceOptions.ReplacingCallback.
  - Prevention: Reference the Aspose.Words.Replacing namespace and use the documented IReplacingCallback interface for custom logic.

- **Missing Field type**
  - Symptom: Compilation fails with CS0103 or CS0246 when the code references the Field class without a proper using directive.
  - Root cause: The Field class resides in Aspose.Words.Fields, which is not imported.
  - Preferred fix: Add `using Aspose.Words.Fields;` and reference Field members through that namespace.
  - Prevention: Include the Fields namespace whenever field manipulation is required.

- **Improper source content for metacharacter replacement**
  - Symptom: Runtime throws InvalidOperationException because no line-break characters are found, resulting in zero replacements.
  - Root cause: The sample document contains plain text without actual ControlChar.LineBreak or paragraph breaks, so metacharacter patterns never match.
  - Preferred fix: Insert real line-break control characters (e.g., ControlChar.LineBreak) into the source document before performing the replace.
  - Prevention: Validate that the document contains the expected control characters or use a regex that matches the actual content.

- **Invalid field handling when ignoring fields**
  - Symptom: Runtime throws InvalidOperationException indicating that a field code was unexpectedly modified.
  - Root cause: FindReplaceOptions.IgnoreFields is set, but the replacement logic still attempts to modify field codes or the document contains only fields matching the search term.
  - Preferred fix: Ensure that the search pattern does not target field codes, or verify that at least one non-field occurrence exists before replacement.
  - Prevention: When using IgnoreFields, design the find pattern to exclude field codes and always check the replacement count.

- **Aspose.Drawing vs System.Drawing type conflict**
  - Symptom: Compilation fails with CS0029 because Aspose.Drawing.Color cannot be assigned to System.Drawing.Color.
  - Root cause: The code mixes types from the two drawing libraries, violating type compatibility.
  - Preferred fix: Use only Aspose.Drawing types (e.g., Aspose.Drawing.Color) and avoid referencing System.Drawing.
  - Prevention: Reference Aspose.Drawing exclusively and qualify type names when ambiguity is possible.

- **DocumentBuilder API misuse**
  - Symptom: Compilation fails with CS1503 or CS1061 because constructors or methods such as DocumentBuilderOptions, InsertParagraphAfter, or ParagraphStyleName are not found.
  - Root cause: The code attempts to use non-existent overloads or members that belong to older versions of the API.
  - Preferred fix: Instantiate DocumentBuilder with a Document instance and use supported methods like InsertParagraph, MoveTo, or ParagraphFormat properties.
  - Prevention: Refer to the current Aspose.Words API reference for DocumentBuilder usage and rely on IntelliSense to select valid members.

- **Incorrect overload usage for regex replace**
  - Symptom: Compilation fails with CS1503 when a string is passed where a Regex is expected, or vice versa.
  - Root cause: The code calls the overload that expects a Regex object but supplies a plain string, or supplies a Regex but provides a FindReplaceOptions where a string is required.
  - Preferred fix: Use the overload `Range.Replace(Regex pattern, string replacement, FindReplaceOptions options)` with a proper Regex instance.
  - Prevention: Match the method signature exactly: provide a Regex object for the pattern and a string for the replacement.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
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

- Periodically verify the Aspose.Words version used in the project to ensure property names and method signatures have not changed.
