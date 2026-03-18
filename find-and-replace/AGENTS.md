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

- `using Aspose.Words;` (62/62 files) ← category-specific
- `using System;` (58/62 files)
- `using Aspose.Words.Replacing;` (50/62 files)
- `using System.Text.RegularExpressions;` (43/62 files)
- `using System.Collections.Generic;` (6/62 files)
- `using System.Text;` (5/62 files)
- `using System.IO;` (5/62 files)
- `using Aspose.Words.Fields;` (4/62 files)
- `using System.Drawing;` (3/62 files)
- `using System.Globalization;` (3/62 files)
- `using Aspose.Words.Drawing;` (2/62 files)
- `using System.Linq;` (1/62 files)

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
| [add-custom-style-each-replaced-paragraph-documentbuilde...](./add-custom-style-each-replaced-paragraph-documentbuilder-during-replacement-process.cs) | `Aspose`, `ParagraphFormat`, `Font` | Add custom style each replaced paragraph documentbuilder during replacement p... |
| [add-hyperlink-each-replaced-email-address-documentbuild...](./add-hyperlink-each-replaced-email-address-documentbuilder-during-replacement-operation.cs) | `DocumentBuilder`, `Font`, `Document` | Add hyperlink each replaced email address documentbuilder during replacement... |
| [add-prefix-each-matched-word-custom-replaceevaluator-de...](./add-prefix-each-matched-word-custom-replaceevaluator-delegate-during-replacement.cs) | `Aspose`, `Document`, `DocumentBuilder` | Add prefix each matched word custom replaceevaluator delegate during replacement |
| [add-suffix-each-matched-numeric-value-custom-replaceeva...](./add-suffix-each-matched-numeric-value-custom-replaceevaluator-during-replace-operation.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Add suffix each matched numeric value custom replaceevaluator during replace... |
| [apply-findreplaceoptions-ignorefields-skip-replacing-fi...](./apply-findreplaceoptions-ignorefields-skip-replacing-field-result-texts-while.cs) | `Document`, `DocumentBuilder`, `Aspose` | Apply findreplaceoptions ignorefields skip replacing field result texts while |
| [apply-metacharacters-replacement-string-insert-line-bre...](./apply-metacharacters-replacement-string-insert-line-breaks-after-each-replaced.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Apply metacharacters replacement string insert line breaks after each replaced |
| [apply-whole-word-matching-avoid-partial-replacements-in...](./apply-whole-word-matching-avoid-partial-replacements-inside-longer-words-such-as.cs) | `Document`, `DocumentBuilder`, `Aspose` | Apply whole word matching avoid partial replacements inside longer words such as |
| [chain-multiple-replace-calls-sequentially-replace-email...](./chain-multiple-replace-calls-sequentially-replace-email-addresses-phone-numbers-urls.cs) | `Range`, `Document`, `Aspose` | Chain multiple replace calls sequentially replace email addresses phone numbe... |
| [change-url-scheme-http-https-while-preserving-rest-each...](./change-url-scheme-http-https-while-preserving-rest-each-hyperlink-regular-expression.cs) | `System`, `Document`, `Aspose` | Change url scheme http https while preserving rest each hyperlink regular exp... |
| [combine-findreplaceoptions-documentbuilder-insert-dynam...](./combine-findreplaceoptions-documentbuilder-insert-dynamic-content-after-each.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Combine findreplaceoptions documentbuilder insert dynamic content after each |
| [configure-findreplaceoptions-replacingcallback-log-each...](./configure-findreplaceoptions-replacingcallback-log-each-replacement-occurrence-custom.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Configure findreplaceoptions replacingcallback log each replacement occurrenc... |
| [convert-all-uppercase-words-title-case-custom-replaceev...](./convert-all-uppercase-words-title-case-custom-replaceevaluator-during-replacement.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Convert all uppercase words title case custom replaceevaluator during replace... |
| [convert-date-strings-january-1-2020-01-01-2020-regular-...](./convert-date-strings-january-1-2020-01-01-2020-regular-expression.cs) | `Match`, `Document`, `DocumentBuilder` | Convert date strings january 1 2020 01 01 2020 regular expression |
| [convert-markdown-style-headings-word-heading-styles-thr...](./convert-markdown-style-headings-word-heading-styles-throughout-document-regular.cs) | `ParagraphFormat`, `StyleIdentifier`, `Aspose` | Convert markdown style headings word heading styles throughout document regular |
| [count-number-replacements-performed-storing-integer-res...](./count-number-replacements-performed-storing-integer-result-returned-replace-method.cs) | `Document`, `DocumentBuilder`, `Aspose` | Count number replacements performed storing integer result returned replace m... |
| [custom-replaceevaluator-replace-only-first-occurrence-p...](./custom-replaceevaluator-replace-only-first-occurrence-pattern-each-section.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Custom replaceevaluator replace only first occurrence pattern each section |
| [docx-file-document-object-replace-all-literal-string-oc...](./docx-file-document-object-replace-all-literal-string-occurrences.cs) | `Document`, `Range`, `Aspose` | Docx file document object replace all literal string occurrences |
| [enable-regex-mode-setting-findreplaceoptions-useregular...](./enable-regex-mode-setting-findreplaceoptions-useregularexpressions-true-pattern-based.cs) | `FindReplaceOptions`, `Document`, `Aspose` | Enable regex mode setting findreplaceoptions useregularexpressions true patte... |
| [execute-case-insensitive-replace-that-changes-all-insta...](./execute-case-insensitive-replace-that-changes-all-instances-color-colour-throughout.cs) | `Document`, `Aspose`, `FindReplaceOptions` | Execute case insensitive replace that changes all instances color colour thro... |
| [execute-regular-expression-replace-that-converts-dates-...](./execute-regular-expression-replace-that-converts-dates-mm-dd-yyyy-yyyy-mm-dd-format.cs) | `Document`, `DocumentBuilder`, `Aspose` | Execute regular expression replace that converts dates mm dd yyyy yyyy mm dd... |
| [expand-macro-names-their-full-code-custom-replaceevalua...](./expand-macro-names-their-full-code-custom-replaceevaluator-delegate-during-replacement.cs) | `Document`, `FindReplaceOptions`, `System` | Expand macro names their full code custom replaceevaluator delegate during re... |
| [ignore-case-match-whole-words-when-updating-product-nam...](./ignore-case-match-whole-words-when-updating-product-names-across-document.cs) | `Document`, `Aspose`, `FindReplaceOptions` | Ignore case match whole words when updating product names across document |
| [ignore-field-codes-during-replacement-enabling-findrepl...](./ignore-field-codes-during-replacement-enabling-findreplaceoptions-ignorefields-before.cs) | `Document`, `DocumentBuilder`, `Aspose` | Ignore field codes during replacement enabling findreplaceoptions ignorefield... |
| [implement-custom-replaceevaluator-delegate-replace-only...](./implement-custom-replaceevaluator-delegate-replace-only-words-longer-than-ten.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Implement custom replaceevaluator delegate replace only words longer than ten |
| [implement-progress-reporter-that-receives-replacement-c...](./implement-progress-reporter-that-receives-replacement-count-after-each-replace-call.cs) | `Document`, `DocumentBuilder`, `FindReplaceOptions` | Implement progress reporter that receives replacement count after each replac... |
| [insert-page-break-after-each-replaced-heading-metachara...](./insert-page-break-after-each-replaced-heading-metacharacter-f-replacement-string.cs) | `Range`, `ParagraphFormat`, `StyleIdentifier` | Insert page break after each replaced heading metacharacter f replacement string |
| [insert-page-number-field-after-each-replaced-heading-do...](./insert-page-number-field-after-each-replaced-heading-documentbuilder-automatic.cs) | `StyleIdentifier`, `Aspose`, `Document` | Insert page number field after each replaced heading documentbuilder automatic |
| [insert-table-figures-after-each-replaced-figure-caption...](./insert-table-figures-after-each-replaced-figure-caption-documentbuilder-automatic.cs) | `Aspose`, `DocumentBuilder`, `Document` | Insert table figures after each replaced figure caption documentbuilder autom... |
| [mask-email-addresses-locating-them-regular-expression-r...](./mask-email-addresses-locating-them-regular-expression-replacing-masked-version.cs) | `Document`, `DocumentBuilder`, `Aspose` | Mask email addresses locating them regular expression replacing masked version |
| [mask-phone-numbers-finding-them-regular-expression-repl...](./mask-phone-numbers-finding-them-regular-expression-replacing-asterisks-privacy.cs) | `Document`, `FindReplaceOptions`, `System` | Mask phone numbers finding them regular expression replacing asterisks privacy |
| ... | | *and 32 more files* |

## Category Statistics
- Total examples: 62

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for find-and-replace patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082348`
<!-- AUTOGENERATED:END -->
