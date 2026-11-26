---
name: comparison
description: Verified C# examples for comparison scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Comparison

## Purpose

An AI coding agent should generate correct, self-contained C# snippets that create sample source documents, configure supported CompareOptions, invoke Document.Compare, inspect the Document.Revisions collection, and persist the result while respecting nullability and resource-management guidelines.

## Non-negotiable conventions

- Do not reference Aspose.Words members that are not present in version 26.11.0.
- Do not invent comparison options, revision members, or enum values that are absent from the official API.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Revision Processing workflow**: 19 examples
- **Stream and batch I/O workflow**: 2 examples
- **Native Comparison API workflow**: 9 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Ensure that the compiled snippet uses only members that exist in Aspose.Words 26.11.0 and that revision handling is performed through the Document.Revisions collection.

## File-to-task reference

- `load-two-docx-files-compare-them-and-save-the-resulting-document-with-revisions-applied.cs`
  - Task: Load two DOCX files, compare them, and save the resulting document with revisions applied.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `load-a-doc-file-compare-it-with-another-doc-and-save-the-result-with-revisions-included.cs`
  - Task: Load a DOC file, compare it with another DOC, and save the result with revisions included.
  - Workflow: revision-processing
  - Outputs: doc
  - Selected engine: mcp
- `use-a-memorystream-to-load-documents-perform-comparison-and-write-the-result-to-a-byte-arr.cs`
  - Task: Use a MemoryStream to load documents, perform comparison, and write the result to a byte array.
  - Workflow: stream-batch-io
  - Outputs: docx
  - Selected engine: mcp
- `batch-process-a-folder-of-document-pairs-generating-compared-versions-with-revision-tracki.cs`
  - Task: Batch process a folder of document pairs, generating compared versions with revision tracking for each pair.
  - Workflow: stream-batch-io
  - Outputs: docx
  - Selected engine: mcp
- `apply-custom-author-name-and-timestamp-when-calling-compare-to-attribute-revisions-to-a-sp.cs`
  - Task: Apply custom author name and timestamp when calling Compare to attribute revisions to a specific user.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `configure-comparisonoptions-to-ignore-formatting-changes-before-performing-document-compar.cs`
  - Task: Configure ComparisonOptions to ignore formatting changes before performing document comparison on large legal contracts.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: mcp
- `configure-comparisonoptions-to-ignore-case-differences-when-comparing-textual-content-acro.cs`
  - Task: Configure ComparisonOptions to ignore case differences when comparing textual content across documents.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: mcp
- `configure-comparisonoptions-to-ignore-whitespace-changes-when-comparing-source-code-docume.cs`
  - Task: Configure ComparisonOptions to ignore whitespace changes when comparing source code documentation files.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: llm
- `enable-detection-of-moved-paragraphs-by-setting-appropriate-flags-in-comparisonoptions-bef.cs`
  - Task: Enable detection of moved paragraphs by setting appropriate flags in ComparisonOptions before comparison.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: mcp
- `set-comparisontarget-to-new-document-so-that-revisions-appear-in-the-second-file.cs`
  - Task: Set ComparisonTarget to new document so that revisions appear in the second file.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `set-comparisonoptions-showdeletedcontent-to-true-to-retain-deleted-text-in-the-comparison.cs`
  - Task: Set ComparisonOptions.ShowDeletedContent to true to retain deleted text in the comparison output.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: mcp
- `use-comparisonoptions-to-limit-comparison-to-specific-sections-of-the-document-for-targete.cs`
  - Task: Use ComparisonOptions to limit comparison to specific sections of the document for targeted analysis.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: mcp
- `compare-a-doc-file-with-a-docx-file-and-verify-that-table-structure-differences-are-detect.cs`
  - Task: Compare a DOC file with a DOCX file and verify that table structure differences are detected.
  - Workflow: native-comparison-api
  - Outputs: docx, doc
  - Selected engine: mcp
- `compare-documents-with-different-page-orientations-and-verify-that-orientation-changes-app.cs`
  - Task: Compare documents with different page orientations and verify that orientation changes appear as revisions.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: llm
- `compare-documents-containing-embedded-charts-and-ensure-that-chart-data-changes-are-captur.cs`
  - Task: Compare documents containing embedded charts and ensure that chart data changes are captured as revisions.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `compare-a-document-against-itself-to-confirm-that-no-revisions-are-generated.cs`
  - Task: Compare a document against itself to confirm that no revisions are generated.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `iterate-through-doc1-revisions-collection-and-log-each-revision-type-and-affected-text.cs`
  - Task: Iterate through doc1.Revisions collection and log each revision type and affected text.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `detect-changes-in-table-cell-formatting-and-log-the-specific-cell-coordinates-for-each-rev.cs`
  - Task: Detect changes in table cell formatting and log the specific cell coordinates for each revision.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: existing_repo
- `create-a-custom-logger-that-records-revision-type-author-and-timestamp-for-each-detected-c.cs`
  - Task: Create a custom logger that records revision type, author, and timestamp for each detected change.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `check-doc1-revisions-count-equals-zero-to-determine-that-two-documents-are-identical-after.cs`
  - Task: Check doc1.Revisions.Count equals zero to determine that two documents are identical after comparison.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `accept-only-formatting-revisions-while-rejecting-all-content-changes-in-the-compared-docum.cs`
  - Task: Accept only formatting revisions while rejecting all content changes in the compared document.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: llm
- `reject-revisions-related-to-header-modifications-and-preserve-footer-changes-in-the-final.cs`
  - Task: Reject revisions related to header modifications and preserve footer changes in the final output.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `after-comparison-programmatically-accept-all-revisions-and-save-the-cleaned-document-as-do.cs`
  - Task: After comparison, programmatically accept all revisions and save the cleaned document as DOCX.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `after-accepting-revisions-remove-all-revision-marks-by-clearing-the-revisions-collection.cs`
  - Task: After accepting revisions, remove all revision marks by clearing the Revisions collection.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `generate-a-summary-report-of-added-deleted-and-modified-paragraphs-using-the-revisions-col.cs`
  - Task: Generate a summary report of added, deleted, and modified paragraphs using the Revisions collection.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: llm
- `save-the-compared-document-as-doc-format-while-preserving-all-revision-metadata.cs`
  - Task: Save the compared document as DOC format while preserving all revision metadata.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `save-the-compared-document-as-docx-format-while-preserving-all-revision-metadata.cs`
  - Task: Save the compared document as DOCX format while preserving all revision metadata.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp
- `implement-error-handling-to-catch-exceptions-thrown-when-loading-unsupported-file-formats.cs`
  - Task: Implement error handling to catch exceptions thrown when loading unsupported file formats during comparison.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: llm
- `use-a-disposable-pattern-to-ensure-document-objects-are-properly-released-after-comparison.cs`
  - Task: Use a disposable pattern to ensure Document objects are properly released after comparison operations.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: mcp
- `validate-that-the-comparison-output-matches-the-online-tool-results-by-comparing-revision.cs`
  - Task: Validate that the comparison output matches the online tool results by comparing revision counts and types.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Use of non-existent CompareOptions members**
  - Symptom: The compiler reports CS0117 errors for properties such as IgnoreCase, IgnoreWhitespace, ShowDeletedContent, DetectMove, DetectMovedParagraphs, and ComparisonTarget.
  - Root cause: The code references CompareOptions members that were removed or never existed in the current Aspose.Words release.
  - Preferred fix: Reference only the properties documented for CompareOptions in version 26.11.0, such as IgnoreCaseChanges, IgnoreFormatting, or use reflection to set optional flags when they are available.
  - Prevention: Consult the official API reference for the exact property names and guard reflective access with version checks.

- **Invocation of undefined Revision members**
  - Symptom: CS1061 errors appear for Revision.GetText, Revision.RevisionDate, Revision.RevisionDateTime, and missing RevisionCollection.Clear.
  - Root cause: The code attempts to call members that are not part of the Revision class in the installed library.
  - Preferred fix: Use Revision.ParentNode?.GetText() to obtain affected text and Revision.DateTime for timestamps; do not call Clear on RevisionCollection because it is read-only.
  - Prevention: Access only the members listed in the Revision documentation and treat the Revisions collection as immutable.

- **Incorrect disposal of Document objects**
  - Symptom: CS1674 errors indicate that Document cannot be used in a using statement because it does not implement IDisposable.
  - Root cause: The code assumes Document implements IDisposable, which it does not in Aspose.Words.
  - Preferred fix: Create Document instances without a using block and rely on the garbage collector, or call Document.Dispose only if a future version adds it.
  - Prevention: Check the type hierarchy before applying the disposable pattern.

- **Missing exception types for loading errors**
  - Symptom: CS0234 and CS0246 errors show that LoadingException, FileFormatException, and LoadFormatException are not defined.
  - Root cause: The code imports exception types that belong to other libraries or older Aspose.Words versions.
  - Preferred fix: Catch the generic Aspose.Words.LoadFormatException or System.IO.IOException as appropriate.
  - Prevention: Reference the exception types that appear in the current API documentation.

- **Assuming revisions without inspecting the collection**
  - Symptom: Logic checks revision counts or types based on hard-coded values rather than enumerating Document.Revisions.
  - Root cause: The example code reports expected revision numbers without verifying them.
  - Preferred fix: Iterate Document.Revisions and count or filter actual items to determine the outcome.
  - Prevention: Always base decisions on the live Revisions collection.

- **Comparing identical documents without introducing differences**
  - Symptom: The comparison produces zero revisions even though the task expects changes.
  - Root cause: Both source documents contain the same content.
  - Preferred fix: Introduce deterministic differences such as altered text, formatting, or structural changes before calling Compare.
  - Prevention: Create sample inputs that deliberately differ in the aspect under test.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
```

```powershell
Copy-Item ..\comparison\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update the example code whenever a new Aspose.Words release changes the CompareOptions or Revision API surface.
