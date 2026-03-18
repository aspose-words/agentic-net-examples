---
name: hyphenation
description: C# examples for hyphenation using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - hyphenation

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **hyphenation** category.
This folder contains standalone C# examples for hyphenation operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **hyphenation**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (39/39 files) ← category-specific
- `using System;` (37/39 files)
- `using System.IO;` (20/39 files)
- `using Aspose.Words.Settings;` (16/39 files)
- `using Aspose.Words.Saving;` (7/39 files)
- `using System.Collections.Generic;` (6/39 files)
- `using System.Globalization;` (6/39 files)
- `using Aspose.Words.Layout;` (4/39 files)
- `using Aspose.Words.Tables;` (3/39 files)
- `using System.Diagnostics;` (2/39 files)
- `using Aspose.Words.Loading;` (2/39 files)
- `using System.Net.Http;` (2/39 files)

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
| [api-query-hyphenation-status-each-word-paragraph-log-re...](./api-query-hyphenation-status-each-word-paragraph-log-results.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Api query hyphenation status each word paragraph log results |
| [apply-hyphenation-document-then-programmatically-adjust...](./apply-hyphenation-document-then-programmatically-adjust-line-spacing-improve.cs) | `HyphenationOptions`, `Aspose`, `Document` | Apply hyphenation document then programmatically adjust line spacing improve |
| [apply-hyphenation-only-selected-range-documentbuilder-v...](./apply-hyphenation-only-selected-range-documentbuilder-verify-layout-changes.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Apply hyphenation only selected range documentbuilder verify layout changes |
| [apply-hyphenation-specific-paragraph-disabling-it-surro...](./apply-hyphenation-specific-paragraph-disabling-it-surrounding-sections.cs) | `CurrentParagraph`, `ParagraphFormat`, `Document` | Apply hyphenation specific paragraph disabling it surrounding sections |
| [batch-convert-docx-files-pdf-while-preserving-hyphenati...](./batch-convert-docx-files-pdf-while-preserving-hyphenation-log-any-documents-that-fail.cs) | `Document`, `System`, `Aspose` | Batch convert docx files pdf while preserving hyphenation log any documents t... |
| [batch-process-collection-docx-files-applying-language-s...](./batch-process-collection-docx-files-applying-language-specific-hyphenation-exporting.cs) | `HyphenationOptions`, `Aspose`, `Hyphenation` | Batch process collection docx files applying language specific hyphenation ex... |
| [check-whether-word-hyphenation-will-be-hyphenated-api-b...](./check-whether-word-hyphenation-will-be-hyphenated-api-before-document-generation.cs) | `Aspose`, `HyphenationOptions`, `Document` | Check whether word hyphenation will be hyphenated api before document generation |
| [compare-layout-differences-between-document-saved-hyphe...](./compare-layout-differences-between-document-saved-hyphenation-disabled-same-document.cs) | `Document`, `HyphenationOptions`, `DocumentBuilder` | Compare layout differences between document saved hyphenation disabled same d... |
| [configure-hyphenation-respect-compound-word-rules-micro...](./configure-hyphenation-respect-compound-word-rules-microsoft-word-german-language.cs) | `Hyphenation`, `HyphenationOptions`, `Document` | Configure hyphenation respect compound word rules microsoft word german language |
| [console-application-that-accepts-document-path-hyphenat...](./console-application-that-accepts-document-path-hyphenation-language-code-outputs.cs) | `HyphenationOptions`, `Document`, `Aspose` | Console application that accepts document path hyphenation language code outputs |
| [develop-function-that-returns-true-if-given-word-will-b...](./develop-function-that-returns-true-if-given-word-will-be-hyphenated-under-current.cs) | `Document`, `DocumentBuilder`, `FirstSection` | Develop function that returns true if given word will be hyphenated under cur... |
| [disable-hyphenation-headings-while-keeping-it-enabled-b...](./disable-hyphenation-headings-while-keeping-it-enabled-body-paragraphs-report.cs) | `ParagraphFormat`, `Aspose`, `HyphenationOptions` | Disable hyphenation headings while keeping it enabled body paragraphs report |
| [docx-containing-mixed-languages-set-appropriate-hyphena...](./docx-containing-mixed-languages-set-appropriate-hyphenation-language-each-section.cs) | `Hyphenation`, `Document`, `System` | Docx containing mixed languages set appropriate hyphenation language each sec... |
| [docx-disable-hyphenation-footnotes-only-compare-footnot...](./docx-disable-hyphenation-footnotes-only-compare-footnote-layout-before-after.cs) | `Document`, `Aspose`, `NodeType` | Docx disable hyphenation footnotes only compare footnote layout before after |
| [docx-enable-hyphenation-export-document-docx-preserving...](./docx-enable-hyphenation-export-document-docx-preserving-hyphenation-marks.cs) | `HyphenationOptions`, `Document`, `Aspose` | Docx enable hyphenation export document docx preserving hyphenation marks |
| [docx-file-register-custom-hunspell-dictionary-enable-au...](./docx-file-register-custom-hunspell-dictionary-enable-automatic-hyphenation.cs) | `Document`, `Aspose`, `System` | Docx file register custom hunspell dictionary enable automatic hyphenation |
| [docx-set-hyphenation-language-ru-ru-result-as-pdf](./docx-set-hyphenation-language-ru-ru-result-as-pdf.cs) | `Document`, `Aspose`, `PDF` | Docx set hyphenation language ru ru result as pdf |
| [enable-hyphenation-document-then-export-docx-verify-hyp...](./enable-hyphenation-document-then-export-docx-verify-hyphenation-marks-are-retained.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Enable hyphenation document then export docx verify hyphenation marks are ret... |
| [enable-hyphenation-globally-document-then-override-it-s...](./enable-hyphenation-globally-document-then-override-it-single-table-cell.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Enable hyphenation globally document then override it single table cell |
| [export-hyphenated-document-pdf-compare-file-size-non-hy...](./export-hyphenated-document-pdf-compare-file-size-non-hyphenated-version.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Export hyphenated document pdf compare file size non hyphenated version |
| [implement-error-handling-incompatible-hyphenation-dicti...](./implement-error-handling-incompatible-hyphenation-dictionaries-provide-descriptive.cs) | `Aspose`, `WarningInfoCollection`, `Hyphenation` | Implement error handling incompatible hyphenation dictionaries provide descri... |
| [integrate-hyphenation-dictionary-updates-ci-pipeline-ke...](./integrate-hyphenation-dictionary-updates-ci-pipeline-keep-language-patterns-current.cs) | `Hyphenation`, `Document`, `System` | Integrate hyphenation dictionary updates ci pipeline keep language patterns c... |
| [list-all-available-hyphenation-dictionaries-system-disp...](./list-all-available-hyphenation-dictionaries-system-display-their-language-codes.cs) | `Hyphenation`, `System`, `Value` | List all available hyphenation dictionaries system display their language codes |
| [measure-pagination-differences-after-enabling-hyphenati...](./measure-pagination-differences-after-enabling-hyphenation-multi-section-report.cs) | `Aspose`, `HyphenationOptions`, `System` | Measure pagination differences after enabling hyphenation multi section report |
| [measure-rendering-time-differences-between-pdf-generati...](./measure-rendering-time-differences-between-pdf-generation-without-hyphenation-enabled.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Measure rendering time differences between pdf generation without hyphenation... |
| [multiple-documents-folder-enable-hyphenation-each-as-hy...](./multiple-documents-folder-enable-hyphenation-each-as-hyphenated-pdf.cs) | `HyphenationOptions`, `Document`, `System` | Multiple documents folder enable hyphenation each as hyphenated pdf |
| [new-document-enable-automatic-hyphenation-start-it-as-d...](./new-document-enable-automatic-hyphenation-start-it-as-docx-file.cs) | `HyphenationOptions`, `Document`, `Aspose` | New document enable automatic hyphenation start it as docx file |
| [pdf-file-enable-hyphenation-render-result-image-visual-...](./pdf-file-enable-hyphenation-render-result-image-visual-inspection.cs) | `PdfLoadOptions`, `Aspose`, `HyphenationOptions` | Pdf file enable hyphenation render result image visual inspection |
| [pdf-hyphenated-document-ensure-that-hyphenation-marks-a...](./pdf-hyphenated-document-ensure-that-hyphenation-marks-are-not-visible-output.cs) | `Aspose`, `Document`, `Words` | Pdf hyphenated document ensure that hyphenation marks are not visible output |
| [programmatically-adjust-paragraph-justification-after-h...](./programmatically-adjust-paragraph-justification-after-hyphenation-prevent-excessive.cs) | `Document`, `DocumentBuilder`, `Aspose` | Programmatically adjust paragraph justification after hyphenation prevent exc... |
| ... | | *and 9 more files* |

## Category Statistics
- Total examples: 39

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for hyphenation patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082414`
<!-- AUTOGENERATED:END -->
