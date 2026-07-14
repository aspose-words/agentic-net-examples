---
name: comparison
description: Verified C# examples for Comparison scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Comparison

## Purpose

This folder is a live, curated example set for Comparison scenarios. Treat every `.cs` file as a standalone console application. The goal is correct, warning-free examples that use documented Aspose APIs and match the original task intent.

## Non-negotiable conventions

- Use `Document.Compare` and inspect actual `document.Revisions` items.
- Create meaningful local source document differences before comparison.
- Use supported `CompareOptions` members only when the task explicitly requires them.
- Bootstrap local files, folders, and MemoryStream-backed documents whenever external-style inputs are needed.
- Guard maybe-null values to avoid nullable-reference warnings such as `CS8600`, `CS8602`, and `CS8604`.

## Recommended workflow selection

- Native Comparison Api workflow: 9 examples
- Revision Processing workflow: 19 examples
- Stream / batch / input-bootstrap workflow: 2 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. API usage must be supported by the configured package versions.
3. Exported outputs must actually be written by the example.
4. Validation scenarios must inspect only the behavior requested by the task.
5. Examples that depend on files, folders, streams, images, or data should bootstrap those inputs locally during the example run.

## File-to-task reference

- `load-two-docx-files-compare-them-and-save-the-resulting-document-with-revisions-applied.cs`
  - Task: Load two DOCX files, compare them, and save the resulting document with revisions applied.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `load-a-doc-file-compare-it-with-another-doc-and-save-the-result-with-revisions-included.cs`
  - Task: Load a DOC file, compare it with another DOC, and save the result with revisions included.
  - Workflow: Revision Processing
  - Outputs: doc
  - Selected engine: mcp
- `use-a-memorystream-to-load-documents-perform-comparison-and-write-the-result-to-a-byte-arr.cs`
  - Task: Use a MemoryStream to load documents, perform comparison, and write the result to a byte array.
  - Workflow: Stream / batch / input-bootstrap
  - Outputs: docx
  - Selected engine: mcp
- `batch-process-a-folder-of-document-pairs-generating-compared-versions-with-revision-tracki.cs`
  - Task: Batch process a folder of document pairs, generating compared versions with revision tracking for each pair.
  - Workflow: Stream / batch / input-bootstrap
  - Outputs: docx
  - Selected engine: mcp
- `apply-custom-author-name-and-timestamp-when-calling-compare-to-attribute-revisions-to-a-sp.cs`
  - Task: Apply custom author name and timestamp when calling Compare to attribute revisions to a specific user.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `configure-comparisonoptions-to-ignore-formatting-changes-before-performing-document-compar.cs`
  - Task: Configure ComparisonOptions to ignore formatting changes before performing document comparison on large legal contracts.
  - Workflow: Native Comparison Api
  - Outputs: docx
  - Selected engine: llm
- `configure-comparisonoptions-to-ignore-case-differences-when-comparing-textual-content-acro.cs`
  - Task: Configure ComparisonOptions to ignore case differences when comparing textual content across documents.
  - Workflow: Native Comparison Api
  - Outputs: docx
  - Selected engine: mcp
- `configure-comparisonoptions-to-ignore-whitespace-changes-when-comparing-source-code-docume.cs`
  - Task: Configure ComparisonOptions to ignore whitespace changes when comparing source code documentation files.
  - Workflow: Native Comparison Api
  - Outputs: docx
  - Selected engine: llm
- `enable-detection-of-moved-paragraphs-by-setting-appropriate-flags-in-comparisonoptions-bef.cs`
  - Task: Enable detection of moved paragraphs by setting appropriate flags in ComparisonOptions before comparison.
  - Workflow: Native Comparison Api
  - Outputs: docx
  - Selected engine: llm
- `set-comparisontarget-to-new-document-so-that-revisions-appear-in-the-second-file.cs`
  - Task: Set ComparisonTarget to new document so that revisions appear in the second file.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `set-comparisonoptions-showdeletedcontent-to-true-to-retain-deleted-text-in-the-comparison.cs`
  - Task: Set ComparisonOptions.ShowDeletedContent to true to retain deleted text in the comparison output.
  - Workflow: Native Comparison Api
  - Outputs: docx
  - Selected engine: mcp
- `use-comparisonoptions-to-limit-comparison-to-specific-sections-of-the-document-for-targete.cs`
  - Task: Use ComparisonOptions to limit comparison to specific sections of the document for targeted analysis.
  - Workflow: Native Comparison Api
  - Outputs: docx
  - Selected engine: mcp
- `compare-a-doc-file-with-a-docx-file-and-verify-that-table-structure-differences-are-detect.cs`
  - Task: Compare a DOC file with a DOCX file and verify that table structure differences are detected.
  - Workflow: Native Comparison Api
  - Outputs: docx, doc
  - Selected engine: mcp
- `compare-documents-with-different-page-orientations-and-verify-that-orientation-changes-app.cs`
  - Task: Compare documents with different page orientations and verify that orientation changes appear as revisions.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `compare-documents-containing-embedded-charts-and-ensure-that-chart-data-changes-are-captur.cs`
  - Task: Compare documents containing embedded charts and ensure that chart data changes are captured as revisions.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: existing_repo
- `compare-a-document-against-itself-to-confirm-that-no-revisions-are-generated.cs`
  - Task: Compare a document against itself to confirm that no revisions are generated.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `iterate-through-doc1-revisions-collection-and-log-each-revision-type-and-affected-text.cs`
  - Task: Iterate through doc1.Revisions collection and log each revision type and affected text.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `detect-changes-in-table-cell-formatting-and-log-the-specific-cell-coordinates-for-each-rev.cs`
  - Task: Detect changes in table cell formatting and log the specific cell coordinates for each revision.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `create-a-custom-logger-that-records-revision-type-author-and-timestamp-for-each-detected-c.cs`
  - Task: Create a custom logger that records revision type, author, and timestamp for each detected change.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `check-doc1-revisions-count-equals-zero-to-determine-that-two-documents-are-identical-after.cs`
  - Task: Check doc1.Revisions.Count equals zero to determine that two documents are identical after comparison.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `accept-only-formatting-revisions-while-rejecting-all-content-changes-in-the-compared-docum.cs`
  - Task: Accept only formatting revisions while rejecting all content changes in the compared document.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `reject-revisions-related-to-header-modifications-and-preserve-footer-changes-in-the-final.cs`
  - Task: Reject revisions related to header modifications and preserve footer changes in the final output.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `after-comparison-programmatically-accept-all-revisions-and-save-the-cleaned-document-as-do.cs`
  - Task: After comparison, programmatically accept all revisions and save the cleaned document as DOCX.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `after-accepting-revisions-remove-all-revision-marks-by-clearing-the-revisions-collection.cs`
  - Task: After accepting revisions, remove all revision marks by clearing the Revisions collection.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `generate-a-summary-report-of-added-deleted-and-modified-paragraphs-using-the-revisions-col.cs`
  - Task: Generate a summary report of added, deleted, and modified paragraphs using the Revisions collection.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `save-the-compared-document-as-doc-format-while-preserving-all-revision-metadata.cs`
  - Task: Save the compared document as DOC format while preserving all revision metadata.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `save-the-compared-document-as-docx-format-while-preserving-all-revision-metadata.cs`
  - Task: Save the compared document as DOCX format while preserving all revision metadata.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp
- `implement-error-handling-to-catch-exceptions-thrown-when-loading-unsupported-file-formats.cs`
  - Task: Implement error handling to catch exceptions thrown when loading unsupported file formats during comparison.
  - Workflow: Native Comparison Api
  - Outputs: docx
  - Selected engine: mcp
- `use-a-disposable-pattern-to-ensure-document-objects-are-properly-released-after-comparison.cs`
  - Task: Use a disposable pattern to ensure Document objects are properly released after comparison operations.
  - Workflow: Native Comparison Api
  - Outputs: docx
  - Selected engine: mcp
- `validate-that-the-comparison-output-matches-the-online-tool-results-by-comparing-revision.cs`
  - Task: Validate that the comparison output matches the online tool results by comparing revision counts and types.
  - Workflow: Revision Processing
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- Identical document comparison
  - Symptom: Comparison yields no revisions when differences were expected.
  - Preferred fix: Create deterministic local source documents with meaningful differences before calling Compare.

- Invented revision reporting
  - Symptom: Counts or revision metadata are reported without inspecting the real revisions collection.
  - Preferred fix: Use document.Revisions directly and filter actual revision items.

- Missing local source documents
  - Symptom: Existing DOCX/DOC/folder/stream inputs are assumed to exist.
  - Preferred fix: Bootstrap local versions, folders, or MemoryStream-backed documents in the example.

- Incorrect post-processing assumption
  - Symptom: Accept/reject workflows do not verify the final revision state.
  - Preferred fix: Check document.Revisions.Count after accepting or rejecting revisions.

- Nullable warnings
  - Symptom: CS8600, CS8602, or CS8604 around maybe-null nodes or filtered results.
  - Preferred fix: Use nullable locals and null-check before dereference or assignment.

## Build and run contract

- Target framework: `net8.0`
- Package: `Aspose.Words` `26.5.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.5.0
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
- Keep examples standalone, bootstrap missing inputs locally, and prefer documented Aspose APIs over speculative shortcuts.
