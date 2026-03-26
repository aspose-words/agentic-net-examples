---
name: comparison
description: C# examples for comparison using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - comparison

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **comparison** category.
This folder contains standalone C# examples for comparison operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **comparison**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (30/30 files) ← category-specific
- `using System;` (29/30 files)
- `using Aspose.Words.Comparing;` (17/30 files)
- `using System.IO;` (9/30 files)
- `using Aspose.Words.Tables;` (2/30 files)
- `using Aspose.Words.Saving;` (2/30 files)
- `using System.Linq;` (1/30 files)
- `using Aspose.Words.Replacing;` (1/30 files)
- `using System.Drawing;` (1/30 files)
- `using Aspose.Words.Drawing;` (1/30 files)
- `using Aspose.Words.Loading;` (1/30 files)
- `using System.Collections.Generic;` (1/30 files)

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
| [accept-only-formatting-revisions-while-rejecting-all-co...](./accept-only-formatting-revisions-while-rejecting-all-content-changes-compared-document.cs) | `Document`, `DocumentBuilder`, `Original` | Accept only formatting revisions while rejecting all content changes compared... |
| [after-accepting-revisions-remove-all-revision-marks-cle...](./after-accepting-revisions-remove-all-revision-marks-clearing-revisions-collection.cs) | `Document`, `Revisions` | After accepting revisions remove all revision marks clearing revisions collec... |
| [after-comparison-programmatically-accept-all-revisions-...](./after-comparison-programmatically-accept-all-revisions-cleaned-document-as-docx.cs) | `Document`, `Original`, `Edited` | After comparison programmatically accept all revisions cleaned document as docx |
| [apply-custom-author-name-timestamp-when-calling-compare...](./apply-custom-author-name-timestamp-when-calling-compare-attribute-revisions-specific.cs) | `Document`, `Original`, `Edited` | Apply custom author name timestamp when calling compare attribute revisions s... |
| [batch-process-folder-document-pairs-generating-compared...](./batch-process-folder-document-pairs-generating-compared-versions-revision-tracking.cs) | `Document`, `AppContext`, `StringComparison` | Batch process folder document pairs generating compared versions revision tra... |
| [check-doc1-revisions-count-equals-zero-determine-that-t...](./check-doc1-revisions-count-equals-zero-determine-that-two-documents-are-identical.cs) | `Document`, `Revisions`, `InvalidOperationException` | Check doc1 revisions count equals zero determine that two documents are ident... |
| [compare-doc-file-docx-file-verify-that-table-structure-...](./compare-doc-file-docx-file-verify-that-table-structure-differences-are-detected.cs) | `Document`, `DocumentBuilder`, `CompareOptions` | Compare doc file docx file verify that table structure differences are detected |
| [compare-document-against-itself-confirm-that-no-revisio...](./compare-document-against-itself-confirm-that-no-revisions-are.cs) | `Document`, `InvalidOperationException`, `Revisions` | Compare document against itself confirm that no revisions are |
| [compare-documents-containing-embedded-charts-ensure-tha...](./compare-documents-containing-embedded-charts-ensure-that-chart-data-changes-are.cs) | `Document`, `Revisions`, `InvalidOperationException` | Compare documents containing embedded charts ensure that chart data changes are |
| [compare-documents-different-page-orientations-verify-th...](./compare-documents-different-page-orientations-verify-that-orientation-changes-appear.cs) | `PageSetup`, `Orientation`, `Document` | Compare documents different page orientations verify that orientation changes... |
| [compared-document-as-doc-format-while-preserving-all-re...](./compared-document-as-doc-format-while-preserving-all-revision-metadata.cs) | `Document`, `DocSaveOptions`, `Original` | Compared document as doc format while preserving all revision metadata |
| [compared-document-as-docx-format-while-preserving-all-r...](./compared-document-as-docx-format-while-preserving-all-revision-metadata.cs) | `Document`, `DocumentBuilder`, `Original` | Compared document as docx format while preserving all revision metadata |
| [comparisonoptions-limit-comparison-specific-sections-do...](./comparisonoptions-limit-comparison-specific-sections-document-targeted-analysis.cs) | `Document`, `DocumentBuilder`, `BreakType` | Comparisonoptions limit comparison specific sections document targeted analysis |
| [configure-comparisonoptions-ignore-case-differences-whe...](./configure-comparisonoptions-ignore-case-differences-when-comparing-textual-content.cs) | `Document`, `CompareOptions`, `Original` | Configure comparisonoptions ignore case differences when comparing textual co... |
| [configure-comparisonoptions-ignore-formatting-changes-b...](./configure-comparisonoptions-ignore-formatting-changes-before-performing-document.cs) | `Document`, `DocumentBuilder`, `Font` | Configure comparisonoptions ignore formatting changes before performing document |
| [configure-comparisonoptions-ignore-whitespace-changes-w...](./configure-comparisonoptions-ignore-whitespace-changes-when-comparing-source-code.cs) | `Document`, `DocumentBuilder`, `CompareOptions` | Configure comparisonoptions ignore whitespace changes when comparing source code |
| [custom-logger-that-records-revision-type-author-timesta...](./custom-logger-that-records-revision-type-author-timestamp-each-detected-change.cs) | `Document`, `DocumentBuilder`, `RevisionLogger` | Custom logger that records revision type author timestamp each detected change |
| [detect-changes-table-cell-formatting-log-specific-cell-...](./detect-changes-table-cell-formatting-log-specific-cell-coordinates-each-revision.cs) | `Document`, `DocumentBuilder`, `NodeType` | Detect changes table cell formatting log specific cell coordinates each revision |
| [disposable-pattern-ensure-document-objects-are-properly...](./disposable-pattern-ensure-document-objects-are-properly-released-after-comparison.cs) | `Document`, `DocumentBuilder`, `Revisions` | Disposable pattern ensure document objects are properly released after compar... |
| [doc-file-compare-it-another-doc-result-revisions-included](./doc-file-compare-it-another-doc-result-revisions-included.cs) | `Document`, `DocumentBuilder`, `ComparisonResult` | Doc file compare it another doc result revisions included |
| [enable-detection-moved-paragraphs-setting-appropriate-f...](./enable-detection-moved-paragraphs-setting-appropriate-flags-comparisonoptions-before.cs) | `Document`, `CompareOptions`, `Original` | Enable detection moved paragraphs setting appropriate flags comparisonoptions... |
| [implement-error-handling-catch-exceptions-thrown-when-u...](./implement-error-handling-catch-exceptions-thrown-when-unsupported-file-formats-during.cs) | `LoadOptions`, `Document`, `DocumentComparer` | Implement error handling catch exceptions thrown when unsupported file format... |
| [iterate-through-doc1-revisions-collection-log-each-revi...](./iterate-through-doc1-revisions-collection-log-each-revision-type-affected-text.cs) | `Document`, `RevisionType`, `ParentStyle` | Iterate through doc1 revisions collection log each revision type affected text |
| [memorystream-documents-perform-comparison-write-result-...](./memorystream-documents-perform-comparison-write-result-byte-array.cs) | `Document`, `Revisions`, `SaveFormat` | Memorystream documents perform comparison write result byte array |
| [reject-revisions-related-header-modifications-preserve-...](./reject-revisions-related-header-modifications-preserve-footer-changes-final-output.cs) | `Document`, `HeaderRevisionCriteria`, `CompareOptions` | Reject revisions related header modifications preserve footer changes final o... |
| [set-comparisonoptions-showdeletedcontent-true-retain-de...](./set-comparisonoptions-showdeletedcontent-true-retain-deleted-text-comparison-output.cs) | `Document`, `DocumentBuilder`, `CompareOptions` | Set comparisonoptions showdeletedcontent true retain deleted text comparison... |
| [set-comparisontarget-new-document-so-that-revisions-app...](./set-comparisontarget-new-document-so-that-revisions-appear-second-file.cs) | `Document`, `DocumentBuilder`, `CompareOptions` | Set comparisontarget new document so that revisions appear second file |
| [summary-report-added-deleted-modified-paragraphs-revisi...](./summary-report-added-deleted-modified-paragraphs-revisions-collection.cs) | `RevisionType`, `Document`, `DocumentBuilder` | Summary report added deleted modified paragraphs revisions collection |
| [two-docx-files-compare-them-resulting-document-revision...](./two-docx-files-compare-them-resulting-document-revisions-applied.cs) | `Document`, `Revisions`, `Original` | Two docx files compare them resulting document revisions applied |
| [validate-that-comparison-output-matches-online-tool-res...](./validate-that-comparison-output-matches-online-tool-results-comparing-revision-counts.cs) | `Revisions`, `DocumentBuilder`, `Paragraph` | Validate that comparison output matches online tool results comparing revisio... |

## Category Statistics
- Total examples: 30

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for comparison patterns.


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
Copy-Item ..\comparison\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

### Notes for Agents and Developers

- Treat every `.cs` file in `comparison/` as a full console program, not a snippet.
- Run one file at a time by copying it to `Program.cs`.
- If a sample needs input documents, images, fonts, or data files, place them in the temporary project directory before running.
- See the root [AGENTS.md](../AGENTS.md) for repository-wide prerequisites, project file template, and testing guidance.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
