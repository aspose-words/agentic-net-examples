---
name: comparison
description: Verified C# examples for comparison scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Comparison

## Purpose

Provide concise, executable guidance for an AI coding assistant to generate reliable Aspose.Words comparison snippets that compile against version 26.10.0 and follow best-practice patterns.

## Non-negotiable conventions

- Do not reference members that are not present in Aspose.Words 26.10.0, such as CompareOptions.IgnoreWhitespaceChanges or Revision.GetText.
- Do not fabricate revision-type constants that the library does not define.
- Do not assume external files exist; generate sample inputs programmatically.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Revision Processing workflow**: 19 examples
- **Stream and batch I/O workflow**: 2 examples
- **Native Comparison API workflow**: 9 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Verify that the compiled code accesses only members defined in the referenced Aspose.Words assembly and that Document.Revisions.Count reflects the expected changes.

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
  - Selected engine: mcp
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
  - Selected engine: existing_repo
- `compare-documents-containing-embedded-charts-and-ensure-that-chart-data-changes-are-captur.cs`
  - Task: Compare documents containing embedded charts and ensure that chart data changes are captured as revisions.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: llm
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
  - Selected engine: mcp
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
  - Selected engine: mcp
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
  - Selected engine: mcp
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
  - Selected engine: mcp
- `use-a-disposable-pattern-to-ensure-document-objects-are-properly-released-after-comparison.cs`
  - Task: Use a disposable pattern to ensure Document objects are properly released after comparison operations.
  - Workflow: native-comparison-api
  - Outputs: docx
  - Selected engine: mcp
- `validate-that-the-comparison-output-matches-the-online-tool-results-by-comparing-revision.cs`
  - Task: Validate that the comparison output matches the online tool results by comparing revision counts and types.
  - Workflow: revision-processing
  - Outputs: docx
  - Selected engine: llm

## Common failure patterns and preferred agent fixes

- **Use of non-existent API members**
  - Symptom: The compiler reports CS1061 or CS0117 because a property or flag such as CompareOptions.IgnoreWhitespaceChanges does not exist.
  - Root cause: The example was written against a newer or different version of Aspose.Words where those members are defined.
  - Preferred fix: Remove the unsupported member and rely on the available options, for example set CompareOptions.IgnoreFormatting to treat whitespace as formatting.
  - Prevention: Consult the official API reference for version 26.10.0 before using any CompareOptions flags.

- **Assumed external documents**
  - Symptom: Runtime code fails with FileNotFoundException or the comparison produces zero revisions because the input files were never created.
  - Root cause: The sample assumes pre-existing DOCX/DOC files or folder structures.
  - Preferred fix: Generate the required source documents in the example code using DocumentBuilder before invoking Compare.
  - Prevention: Always bootstrap sample inputs locally; never rely on external files that are not part of the snippet.

- **Missing revision members**
  - Symptom: Compilation errors such as CS1061 for Revision.GetText or Revision.Node appear.
  - Root cause: The code accesses members that are not part of the Revision class in the target library.
  - Preferred fix: Use Revision.ParentNode to obtain the affected node and call GetText on that node instead of on Revision.
  - Prevention: Reference the Revision documentation and use only the properties and methods that exist.

- **Incorrect collection manipulation**
  - Symptom: CS1061 errors arise when calling Clear on RevisionCollection or CopyTo with an unsupported overload.
  - Root cause: The example uses collection methods that the library does not implement.
  - Preferred fix: Iterate over the collection and remove items individually if needed, or simply accept all revisions without clearing the collection.
  - Prevention: Check the RevisionCollection API for supported operations before invoking them.

- **Nullable reference misuse**
  - Symptom: Compiler warnings CS8600, CS8602, or CS8604 appear when a possibly null value is assigned or dereferenced.
  - Root cause: The code does not guard against null results from properties such as Revision.ParentNode.
  - Preferred fix: Introduce nullable locals and perform explicit null checks before accessing members.
  - Prevention: Enable nullable reference analysis and follow the nullable_warning_prevention guidelines.

- **Incorrect disposal pattern**
  - Symptom: CS1674 indicates that Document does not implement IDisposable, yet it is used in a using statement.
  - Root cause: The example assumes Document implements IDisposable, which it does not in this version.
  - Preferred fix: Do not place Document in a using block; rely on the library's internal resource management.
  - Prevention: Wrap only objects that implement IDisposable (e.g., MemoryStream) in using statements.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.10.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
```

```powershell
Copy-Item ..\comparison\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update the rules summary whenever a new Aspose.Words version introduces additional CompareOptions flags or revision members.
