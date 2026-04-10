---
name: comparison
description: Verified C# examples for comparison scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Comparison

## Purpose

This folder is a **live, curated example set** for comparison scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free document comparison and revision-processing workflows using direct Aspose.Words APIs.

## Non-negotiable conventions

- Use `Document.Compare` and inspect actual `document.Revisions` items.
- Create meaningful local source document differences before comparison.
- Use supported `CompareOptions` members only when the task explicitly requires them.
- Bootstrap local files, folders, and MemoryStream-backed documents whenever external-style inputs are needed.
- Guard maybe-null values to avoid nullable-reference warnings such as `CS8600`, `CS8602`, and `CS8604`.

## Recommended workflow selection

- **Native comparison API workflow**: 9 examples
- **Revision-processing workflow**: 19 examples
- **Stream / batch / input-bootstrap workflow**: 2 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. Comparison scenarios must create real source differences unless the task explicitly requires a zero-difference case.
3. Revision counts, filters, acceptance, and rejection must operate on actual revision objects.
4. Exported outputs (DOCX/TXT/CSV/JSON/etc.) must actually be written when required.
5. Examples that depend on files, folders, or streams should bootstrap those inputs locally during the example run.

## File-to-task reference

- `load-two-docx-files-compare-them-and-save-the-resulting-document-with-revisions-applied.cs`
  - Task: Load two DOCX files, compare them, and save the resulting document with revisions applied.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `load-a-doc-file-compare-it-with-another-doc-and-save-the-result-with-revisions-included.cs`
  - Task: Load a DOC file, compare it with another DOC, and save the result with revisions included.
  - Workflow: revision-processing
  - Outputs: doc
  - Selected engine: verified
- `use-a-memorystream-to-load-documents-perform-comparison-and-write-the-result-to-a-byte-arr.cs`
  - Task: Use a MemoryStream to load documents, perform comparison, and write the result to a byte array.
  - Workflow: stream-batch-io
  - Outputs: docx
  - Selected engine: verified
- `batch-process-a-folder-of-document-pairs-generating-compared-versions-with-revision-tracki.cs`
  - Task: Batch process a folder of document pairs, generating compared versions with revision tracking for each pair.
  - Workflow: stream-batch-io
  - Outputs: docx
  - Selected engine: verified
- `apply-custom-author-name-and-timestamp-when-calling-compare-to-attribute-revisions-to-a-sp.cs`
  - Task: Apply custom author name and timestamp when calling Compare to attribute revisions to a specific user.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `configure-comparisonoptions-to-ignore-formatting-changes-before-performing-document-compar.cs`
  - Task: Configure ComparisonOptions to ignore formatting changes before performing document comparison on large legal contracts.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: verified
- `configure-comparisonoptions-to-ignore-case-differences-when-comparing-textual-content-acro.cs`
  - Task: Configure ComparisonOptions to ignore case differences when comparing textual content across documents.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: verified
- `configure-comparisonoptions-to-ignore-whitespace-changes-when-comparing-source-code-docume.cs`
  - Task: Configure ComparisonOptions to ignore whitespace changes when comparing source code documentation files.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: verified
- `enable-detection-of-moved-paragraphs-by-setting-appropriate-flags-in-comparisonoptions-bef.cs`
  - Task: Enable detection of moved paragraphs by setting appropriate flags in ComparisonOptions before comparison.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: verified
- `set-comparisontarget-to-new-document-so-that-revisions-appear-in-the-second-file.cs`
  - Task: Set ComparisonTarget to new document so that revisions appear in the second file.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `set-comparisonoptions-showdeletedcontent-to-true-to-retain-deleted-text-in-the-comparison.cs`
  - Task: Set ComparisonOptions.ShowDeletedContent to true to retain deleted text in the comparison output.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: verified
- `use-comparisonoptions-to-limit-comparison-to-specific-sections-of-the-document-for-targete.cs`
  - Task: Use ComparisonOptions to limit comparison to specific sections of the document for targeted analysis.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: verified
- `compare-a-doc-file-with-a-docx-file-and-verify-that-table-structure-differences-are-detect.cs`
  - Task: Compare a DOC file with a DOCX file and verify that table structure differences are detected.
  - Workflow: native-comparison-api
  - Outputs: docx, doc
  - Selected engine: verified
- `compare-documents-with-different-page-orientations-and-verify-that-orientation-changes-app.cs`
  - Task: Compare documents with different page orientations and verify that orientation changes appear as revisions.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `compare-documents-containing-embedded-charts-and-ensure-that-chart-data-changes-are-captur.cs`
  - Task: Compare documents containing embedded charts and ensure that chart data changes are captured as revisions.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `compare-a-document-against-itself-to-confirm-that-no-revisions-are-generated.cs`
  - Task: Compare a document against itself to confirm that no revisions are generated.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `iterate-through-doc1-revisions-collection-and-log-each-revision-type-and-affected-text.cs`
  - Task: Iterate through doc1.Revisions collection and log each revision type and affected text.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `detect-changes-in-table-cell-formatting-and-log-the-specific-cell-coordinates-for-each-rev.cs`
  - Task: Detect changes in table cell formatting and log the specific cell coordinates for each revision.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `create-a-custom-logger-that-records-revision-type-author-and-timestamp-for-each-detected-c.cs`
  - Task: Create a custom logger that records revision type, author, and timestamp for each detected change.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `check-doc1-revisions-count-equals-zero-to-determine-that-two-documents-are-identical-after.cs`
  - Task: Check doc1.Revisions.Count equals zero to determine that two documents are identical after comparison.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `accept-only-formatting-revisions-while-rejecting-all-content-changes-in-the-compared-docum.cs`
  - Task: Accept only formatting revisions while rejecting all content changes in the compared document.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `reject-revisions-related-to-header-modifications-and-preserve-footer-changes-in-the-final.cs`
  - Task: Reject revisions related to header modifications and preserve footer changes in the final output.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `after-comparison-programmatically-accept-all-revisions-and-save-the-cleaned-document-as-do.cs`
  - Task: After comparison, programmatically accept all revisions and save the cleaned document as DOCX.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `after-accepting-revisions-remove-all-revision-marks-by-clearing-the-revisions-collection.cs`
  - Task: After accepting revisions, remove all revision marks by clearing the Revisions collection.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `generate-a-summary-report-of-added-deleted-and-modified-paragraphs-using-the-revisions-col.cs`
  - Task: Generate a summary report of added, deleted, and modified paragraphs using the Revisions collection.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `save-the-compared-document-as-doc-format-while-preserving-all-revision-metadata.cs`
  - Task: Save the compared document as DOC format while preserving all revision metadata.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `save-the-compared-document-as-docx-format-while-preserving-all-revision-metadata.cs`
  - Task: Save the compared document as DOCX format while preserving all revision metadata.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified
- `implement-error-handling-to-catch-exceptions-thrown-when-loading-unsupported-file-formats.cs`
  - Task: Implement error handling to catch exceptions thrown when loading unsupported file formats during comparison.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: verified
- `use-a-disposable-pattern-to-ensure-document-objects-are-properly-released-after-comparison.cs`
  - Task: Use a disposable pattern to ensure Document objects are properly released after comparison operations.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: verified
- `validate-that-the-comparison-output-matches-the-online-tool-results-by-comparing-revision.cs`
  - Task: Validate that the comparison output matches the online tool results by comparing revision counts and types.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: verified

## Common failure patterns and preferred agent fixes

- **Identical document comparison**
  - Symptom: Comparison yields no revisions when differences were expected.
  - Preferred fix: Create deterministic local source documents with meaningful differences before calling Compare.

- **Invented revision reporting**
  - Symptom: Counts or revision metadata are reported without inspecting the real revisions collection.
  - Preferred fix: Use document.Revisions directly and filter actual revision items.

- **Missing local source documents**
  - Symptom: Existing DOCX/DOC/folder/stream inputs are assumed to exist.
  - Preferred fix: Bootstrap local versions, folders, or MemoryStream-backed documents in the example.

- **Incorrect post-processing assumption**
  - Symptom: Accept/reject workflows do not verify the final revision state.
  - Preferred fix: Check document.Revisions.Count after accepting or rejecting revisions.

- **Nullable warnings**
  - Symptom: CS8600, CS8602, or CS8604 around maybe-null nodes or filtered results.
  - Preferred fix: Use nullable locals and null-check before dereference or assignment.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\comparison\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer direct Aspose.Words APIs over speculative shortcuts.
