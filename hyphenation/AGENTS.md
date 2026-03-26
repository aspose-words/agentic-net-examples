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

- `using Aspose.Words;` (37/37 files) ← category-specific
- `using System;` (36/37 files)
- `using System.IO;` (21/37 files)
- `using Aspose.Words.Settings;` (13/37 files)
- `using Aspose.Words.Saving;` (7/37 files)
- `using System.Collections.Generic;` (5/37 files)
- `using System.Globalization;` (4/37 files)
- `using Aspose.Words.Tables;` (4/37 files)
- `using Aspose.Words.Layout;` (4/37 files)
- `using System.Diagnostics;` (2/37 files)
- `using Aspose.Words.Drawing;` (1/37 files)
- `using Aspose.Words.Notes;` (1/37 files)

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
| [apply-hyphenation-document-then-programmatically-adjust...](./apply-hyphenation-document-then-programmatically-adjust-line-spacing-improve.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Apply hyphenation document then programmatically adjust line spacing improve |
| [apply-hyphenation-only-selected-range-documentbuilder-v...](./apply-hyphenation-only-selected-range-documentbuilder-verify-layout-changes.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Apply hyphenation only selected range documentbuilder verify layout changes |
| [apply-hyphenation-specific-paragraph-disabling-it-surro...](./apply-hyphenation-specific-paragraph-disabling-it-surrounding-sections.cs) | `CurrentParagraph`, `ParagraphFormat`, `Document` | Apply hyphenation specific paragraph disabling it surrounding sections |
| [batch-convert-docx-files-pdf-while-preserving-hyphenati...](./batch-convert-docx-files-pdf-while-preserving-hyphenation-log-any-documents-that-fail.cs) | `Document`, `AppDomain`, `CurrentDomain` | Batch convert docx files pdf while preserving hyphenation log any documents t... |
| [batch-process-collection-docx-files-applying-language-s...](./batch-process-collection-docx-files-applying-language-specific-hyphenation-exporting.cs) | `HyphenationOptions`, `Hyphenation`, `Document` | Batch process collection docx files applying language specific hyphenation ex... |
| [check-whether-word-hyphenation-will-be-hyphenated-api-b...](./check-whether-word-hyphenation-will-be-hyphenated-api-before-document-generation.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Check whether word hyphenation will be hyphenated api before document generation |
| [compare-layout-differences-between-document-saved-hyphe...](./compare-layout-differences-between-document-saved-hyphenation-disabled-same-document.cs) | `Document`, `HyphenationOptions`, `DocumentBuilder` | Compare layout differences between document saved hyphenation disabled same d... |
| [configure-hyphenation-respect-compound-word-rules-micro...](./configure-hyphenation-respect-compound-word-rules-microsoft-word-german-language.cs) | `Hyphenation`, `HyphenationOptions`, `Document` | Configure hyphenation respect compound word rules microsoft word german language |
| [console-application-that-accepts-document-path-hyphenat...](./console-application-that-accepts-document-path-hyphenation-language-code-outputs.cs) | `HyphenationOptions`, `Document`, `Hyphenation` | Console application that accepts document path hyphenation language code outputs |
| [develop-function-that-returns-true-if-given-word-will-b...](./develop-function-that-returns-true-if-given-word-will-be-hyphenated-under-current.cs) | `Document`, `DocumentBuilder`, `FirstSection` | Develop function that returns true if given word will be hyphenated under cur... |
| [disable-hyphenation-headings-while-keeping-it-enabled-b...](./disable-hyphenation-headings-while-keeping-it-enabled-body-paragraphs-report.cs) | `ParagraphFormat`, `HyphenationOptions`, `Document` | Disable hyphenation headings while keeping it enabled body paragraphs report |
| [docx-disable-hyphenation-footnotes-only-compare-footnot...](./docx-disable-hyphenation-footnotes-only-compare-footnote-layout-before-after.cs) | `Document`, `NodeType`, `DocumentBuilder` | Docx disable hyphenation footnotes only compare footnote layout before after |
| [docx-enable-hyphenation-export-document-docx-preserving...](./docx-enable-hyphenation-export-document-docx-preserving-hyphenation-marks.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Docx enable hyphenation export document docx preserving hyphenation marks |
| [docx-file-register-custom-hunspell-dictionary-enable-au...](./docx-file-register-custom-hunspell-dictionary-enable-automatic-hyphenation.cs) | `Document`, `DocumentBuilder`, `Hyphenation` | Docx file register custom hunspell dictionary enable automatic hyphenation |
| [enable-hyphenation-document-then-export-docx-verify-hyp...](./enable-hyphenation-document-then-export-docx-verify-hyphenation-marks-are-retained.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Enable hyphenation document then export docx verify hyphenation marks are ret... |
| [enable-hyphenation-globally-document-then-override-it-s...](./enable-hyphenation-globally-document-then-override-it-single-table-cell.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Enable hyphenation globally document then override it single table cell |
| [export-hyphenated-document-pdf-compare-file-size-non-hy...](./export-hyphenated-document-pdf-compare-file-size-non-hyphenated-version.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Export hyphenated document pdf compare file size non hyphenated version |
| [implement-error-handling-incompatible-hyphenation-dicti...](./implement-error-handling-incompatible-hyphenation-dictionaries-provide-descriptive.cs) | `WarningInfoCollection`, `Hyphenation` | Implement error handling incompatible hyphenation dictionaries provide descri... |
| [integrate-hyphenation-dictionary-updates-ci-pipeline-ke...](./integrate-hyphenation-dictionary-updates-ci-pipeline-keep-language-patterns-current.cs) | `Hyphenation`, `Document`, `FolderHyphenationCallback` | Integrate hyphenation dictionary updates ci pipeline keep language patterns c... |
| [list-all-available-hyphenation-dictionaries-system-disp...](./list-all-available-hyphenation-dictionaries-system-display-their-language-codes.cs) | `Hyphenation`, `Value`, `Text` | List all available hyphenation dictionaries system display their language codes |
| [measure-pagination-differences-after-enabling-hyphenati...](./measure-pagination-differences-after-enabling-hyphenation-multi-section-report.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Measure pagination differences after enabling hyphenation multi section report |
| [measure-rendering-time-differences-between-pdf-generati...](./measure-rendering-time-differences-between-pdf-generation-without-hyphenation-enabled.cs) | `HyphenationOptions`, `Document`, `DocumentBuilder` | Measure rendering time differences between pdf generation without hyphenation... |
| [multiple-documents-folder-enable-hyphenation-each-as-hy...](./multiple-documents-folder-enable-hyphenation-each-as-hyphenated-pdf.cs) | `HyphenationOptions`, `Document` | Multiple documents folder enable hyphenation each as hyphenated pdf |
| [new-document-enable-automatic-hyphenation-start-it-as-d...](./new-document-enable-automatic-hyphenation-start-it-as-docx-file.cs) | `HyphenationOptions`, `Document`, `HyphenatedDocument` | New document enable automatic hyphenation start it as docx file |
| [pdf-file-enable-hyphenation-render-result-image-visual-...](./pdf-file-enable-hyphenation-render-result-image-visual-inspection.cs) | `Document`, `DocumentBuilder`, `ImageSaveOptions` | Pdf file enable hyphenation render result image visual inspection |
| [pdf-hyphenated-document-ensure-that-hyphenation-marks-a...](./pdf-hyphenated-document-ensure-that-hyphenation-marks-are-not-visible-output.cs) | `Document`, `PdfSaveOptions`, `HyphenatedDocument` | Pdf hyphenated document ensure that hyphenation marks are not visible output |
| [programmatically-adjust-paragraph-justification-after-h...](./programmatically-adjust-paragraph-justification-after-hyphenation-prevent-excessive.cs) | `Document`, `DocumentBuilder`, `HyphenationOptions` | Programmatically adjust paragraph justification after hyphenation prevent exc... |
| [register-external-libreoffice-dictionary-github-apply-i...](./register-external-libreoffice-dictionary-github-apply-it-documents-written-spanish.cs) | `Hyphenation`, `Document`, `DocumentBuilder` | Register external libreoffice dictionary github apply it documents written sp... |
| [retrieve-hyphenation-patterns-french-language-log-them-...](./retrieve-hyphenation-patterns-french-language-log-them-debugging-purposes.cs) | `Hyphenation`, `FileMode`, `FileAccess` | Retrieve hyphenation patterns french language log them debugging purposes |
| ... | | *and 7 more files* |

## Category Statistics
- Total examples: 37

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for hyphenation patterns.


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
Copy-Item ..\hyphenation\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

### Notes for Agents and Developers

- Treat every `.cs` file in `hyphenation/` as a full console program, not a snippet.
- Run one file at a time by copying it to `Program.cs`.
- If a sample needs input documents, images, fonts, or data files, place them in the temporary project directory before running.
- See the root [AGENTS.md](../AGENTS.md) for repository-wide prerequisites, project file template, and testing guidance.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
