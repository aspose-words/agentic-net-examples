---
name: find-and-replace
description: C# examples for find-and-replace using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - find-and-replace

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **find-and-replace** category.
This folder contains standalone C# examples for find-and-replace operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **find-and-replace**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (59/61 files) ← category-specific
- `using System;` (57/61 files)
- `using Aspose.Words.Replacing;` (48/61 files)
- `using System.Text.RegularExpressions;` (42/61 files)
- `using System.IO;` (12/61 files)
- `using System.Text;` (5/61 files)
- `using System.Collections.Generic;` (5/61 files)
- `using Aspose.Words.Fields;` (4/61 files)
- `using System.Drawing;` (3/61 files)
- `using System.Globalization;` (3/61 files)
- `using Aspose.Words.Drawing;` (1/61 files)
- `using System.Linq;` (1/61 files)

## Common Code Pattern

Most files follow this pattern:

```csharp
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
// ... operations ...
doc.Save("output.docx");
```

## Files in this folder

| File | Key APIs | Description |
|------|----------|-------------|
| [add-custom-style-each-replaced-paragraph-documentbuilde...](./add-custom-style-each-replaced-paragraph-documentbuilder-during-replacement-process.cs) | `ParagraphFormat`, `Font`, `Document` | Add custom style each replaced paragraph documentbuilder during replacement p... |
| [add-hyperlink-each-replaced-email-address-documentbuild...](./add-hyperlink-each-replaced-email-address-documentbuilder-during-replacement-operation.cs) | `DocumentBuilder`, `Font`, `Document` | Add hyperlink each replaced email address documentbuilder during replacement... |
| [add-prefix-each-matched-word-custom-replaceevaluator-de...](./add-prefix-each-matched-word-custom-replaceevaluator-delegate-during-replacement.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Add prefix each matched word custom replaceevaluator delegate during replacement |
| [add-suffix-each-matched-numeric-value-custom-replaceeva...](./add-suffix-each-matched-numeric-value-custom-replaceevaluator-during-replace-operation.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Add suffix each matched numeric value custom replaceevaluator during replace... |
| [apply-findreplaceoptions-ignorefields-skip-replacing-fi...](./apply-findreplaceoptions-ignorefields-skip-replacing-field-result-texts-while.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Apply findreplaceoptions ignorefields skip replacing field result texts while |
| [apply-metacharacters-replacement-string-insert-line-bre...](./apply-metacharacters-replacement-string-insert-line-breaks-after-each-replaced.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Apply metacharacters replacement string insert line breaks after each replaced |
| [apply-whole-word-matching-avoid-partial-replacements-in...](./apply-whole-word-matching-avoid-partial-replacements-inside-longer-words-such-as.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Apply whole word matching avoid partial replacements inside longer words such as |
| [chain-multiple-replace-calls-sequentially-replace-email...](./chain-multiple-replace-calls-sequentially-replace-email-addresses-phone-numbers-urls.cs) | `Range`, `Document`, `RegexOptions` | Chain multiple replace calls sequentially replace email addresses phone numbe... |
| [change-url-scheme-http-https-while-preserving-rest-each...](./change-url-scheme-http-https-while-preserving-rest-each-hyperlink-regular-expression.cs) | `Document`, `Text`, `Input` | Change url scheme http https while preserving rest each hyperlink regular exp... |
| [combine-findreplaceoptions-documentbuilder-insert-dynam...](./combine-findreplaceoptions-documentbuilder-insert-dynamic-content-after-each.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Combine findreplaceoptions documentbuilder insert dynamic content after each |
| [configure-findreplaceoptions-replacingcallback-log-each...](./configure-findreplaceoptions-replacingcallback-log-each-replacement-occurrence-custom.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Configure findreplaceoptions replacingcallback log each replacement occurrenc... |
| [convert-all-uppercase-words-title-case-custom-replaceev...](./convert-all-uppercase-words-title-case-custom-replaceevaluator-during-replacement.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Convert all uppercase words title case custom replaceevaluator during replace... |
| [convert-date-strings-january-1-2020-01-01-2020-regular-...](./convert-date-strings-january-1-2020-01-01-2020-regular-expression.cs) | `Match`, `Document`, `DocumentBuilder` | Convert date strings january 1 2020 01 01 2020 regular expression |
| [convert-markdown-style-headings-word-heading-styles-thr...](./convert-markdown-style-headings-word-heading-styles-throughout-document-regular.cs) | `ParagraphFormat`, `StyleIdentifier`, `Document` | Convert markdown style headings word heading styles throughout document regular |
| [count-number-replacements-performed-storing-integer-res...](./count-number-replacements-performed-storing-integer-result-returned-replace-method.cs) | `Document`, `DocumentBuilder`, `Range` | Count number replacements performed storing integer result returned replace m... |
| [custom-replaceevaluator-replace-only-first-occurrence-p...](./custom-replaceevaluator-replace-only-first-occurrence-pattern-each-section.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Custom replaceevaluator replace only first occurrence pattern each section |
| [docx-file-document-object-replace-all-literal-string-oc...](./docx-file-document-object-replace-all-literal-string-occurrences.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Docx file document object replace all literal string occurrences |
| [enable-regex-mode-setting-findreplaceoptions-useregular...](./enable-regex-mode-setting-findreplaceoptions-useregularexpressions-true-pattern-based.cs) | `FindReplaceOptions`, `Document`, `Text` | Enable regex mode setting findreplaceoptions useregularexpressions true patte... |
| [execute-case-insensitive-replace-that-changes-all-insta...](./execute-case-insensitive-replace-that-changes-all-instances-color-colour-throughout.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Execute case insensitive replace that changes all instances color colour thro... |
| [execute-regular-expression-replace-that-converts-dates-...](./execute-regular-expression-replace-that-converts-dates-mm-dd-yyyy-yyyy-mm-dd-format.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Execute regular expression replace that converts dates mm dd yyyy yyyy mm dd... |
| [expand-macro-names-their-full-code-custom-replaceevalua...](./expand-macro-names-their-full-code-custom-replaceevaluator-delegate-during-replacement.cs) | `Document`, `FindReplaceOptions`, `Match` | Expand macro names their full code custom replaceevaluator delegate during re... |
| [ignore-case-match-whole-words-when-updating-product-nam...](./ignore-case-match-whole-words-when-updating-product-names-across-document.cs) | `Document`, `FindReplaceOptions`, `Input` | Ignore case match whole words when updating product names across document |
| [ignore-field-codes-during-replacement-enabling-findrepl...](./ignore-field-codes-during-replacement-enabling-findreplaceoptions-ignorefields-before.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Ignore field codes during replacement enabling findreplaceoptions ignorefield... |
| [implement-custom-replaceevaluator-delegate-replace-only...](./implement-custom-replaceevaluator-delegate-replace-only-words-longer-than-ten.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Implement custom replaceevaluator delegate replace only words longer than ten |
| [implement-progress-reporter-that-receives-replacement-c...](./implement-progress-reporter-that-receives-replacement-count-after-each-replace-call.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Implement progress reporter that receives replacement count after each replac... |
| [insert-page-break-after-each-replaced-heading-metachara...](./insert-page-break-after-each-replaced-heading-metacharacter-f-replacement-string.cs) | `Range`, `ParagraphFormat`, `StyleIdentifier` | Insert page break after each replaced heading metacharacter f replacement string |
| [insert-page-number-field-after-each-replaced-heading-do...](./insert-page-number-field-after-each-replaced-heading-documentbuilder-automatic.cs) | `StyleIdentifier`, `Document`, `Input` | Insert page number field after each replaced heading documentbuilder automatic |
| [insert-table-figures-after-each-replaced-figure-caption...](./insert-table-figures-after-each-replaced-figure-caption-documentbuilder-automatic.cs) | `DocumentBuilder`, `Document`, `InsertTableOfFiguresAfterCaption` | Insert table figures after each replaced figure caption documentbuilder autom... |
| [mask-email-addresses-locating-them-regular-expression-r...](./mask-email-addresses-locating-them-regular-expression-replacing-masked-version.cs) | `Document`, `DocumentBuilder`, `Text` | Mask email addresses locating them regular expression replacing masked version |
| [mask-phone-numbers-finding-them-regular-expression-repl...](./mask-phone-numbers-finding-them-regular-expression-replacing-asterisks-privacy.cs) | `Text`, `Value` | Mask phone numbers finding them regular expression replacing asterisks privacy |
| ... | | *and 31 more files* |

## Category Statistics
- Total examples: 61

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for find-and-replace patterns.


## Command Reference

### Build and Run

Files in this folder are standalone `.cs` examples. Run one example at a time by copying it into a temporary console project as `Program.cs`.

```bash
# Create a temporary console project from the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\find-and-replace\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

### Notes for Agents and Developers

- Treat every `.cs` file in `find-and-replace/` as a full console program, not a snippet.
- Run one file at a time by copying it to `Program.cs`.
- If a sample needs input documents, images, fonts, or data files, place them in the temporary project directory before running.
- See the root [AGENTS.md](../AGENTS.md) for repository-wide prerequisites, project file template, and testing guidance.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
