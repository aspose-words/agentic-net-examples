# Comparison Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **Comparison** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Comparison**
- Slug: **comparison**
- Total examples: **30**
- Publish-ready successful examples: **30 / 30**
- Native comparison API examples: **9**
- Revision-processing examples: **19**
- Stream / batch / input-bootstrap examples: **2**

## Category rules that shaped these examples

- Use native `Document.Compare` APIs directly and operate on real `document.Revisions` items.
- Create realistic local source documents whenever the task mentions existing files, streams, folders, or multiple versions.
- Use supported `CompareOptions` properties only.
- Verify zero-difference and expected-difference scenarios explicitly through the revisions collection.
- Avoid nullable-reference warnings by null-checking maybe-null results before dereference or assignment.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
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

Replace `<example-file>` with a real file from `comparison/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# PowerShell example
Copy-Item ..\comparison\load-two-docx-files-compare-them-and-save-the-resulting-document-with-revisions-applied.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-two-docx-files-compare-them-and-save-the-resulting-document-with-revisions-applied.cs` | Load two DOCX files, compare them, and save the resulting document with revisions applied. | revision-processing | docx | verified |
| 2 | `load-a-doc-file-compare-it-with-another-doc-and-save-the-result-with-revisions-included.cs` | Load a DOC file, compare it with another DOC, and save the result with revisions included. | revision-processing | doc | verified |
| 3 | `use-a-memorystream-to-load-documents-perform-comparison-and-write-the-result-to-a-byte-arr.cs` | Use a MemoryStream to load documents, perform comparison, and write the result to a byte array. | stream-batch-io | docx | verified |
| 4 | `batch-process-a-folder-of-document-pairs-generating-compared-versions-with-revision-tracki.cs` | Batch process a folder of document pairs, generating compared versions with revision tracking for each pair. | stream-batch-io | docx | verified |
| 5 | `apply-custom-author-name-and-timestamp-when-calling-compare-to-attribute-revisions-to-a-sp.cs` | Apply custom author name and timestamp when calling Compare to attribute revisions to a specific user. | revision-processing | docx | verified |
| 6 | `configure-comparisonoptions-to-ignore-formatting-changes-before-performing-document-compar.cs` | Configure ComparisonOptions to ignore formatting changes before performing document comparison on large legal contracts. | native-comparison-api | docx | verified |
| 7 | `configure-comparisonoptions-to-ignore-case-differences-when-comparing-textual-content-acro.cs` | Configure ComparisonOptions to ignore case differences when comparing textual content across documents. | native-comparison-api | docx | verified |
| 8 | `configure-comparisonoptions-to-ignore-whitespace-changes-when-comparing-source-code-docume.cs` | Configure ComparisonOptions to ignore whitespace changes when comparing source code documentation files. | native-comparison-api | docx | verified |
| 9 | `enable-detection-of-moved-paragraphs-by-setting-appropriate-flags-in-comparisonoptions-bef.cs` | Enable detection of moved paragraphs by setting appropriate flags in ComparisonOptions before comparison. | native-comparison-api | docx | verified |
| 10 | `set-comparisontarget-to-new-document-so-that-revisions-appear-in-the-second-file.cs` | Set ComparisonTarget to new document so that revisions appear in the second file. | revision-processing | docx | verified |
| 11 | `set-comparisonoptions-showdeletedcontent-to-true-to-retain-deleted-text-in-the-comparison.cs` | Set ComparisonOptions.ShowDeletedContent to true to retain deleted text in the comparison output. | native-comparison-api | docx | verified |
| 12 | `use-comparisonoptions-to-limit-comparison-to-specific-sections-of-the-document-for-targete.cs` | Use ComparisonOptions to limit comparison to specific sections of the document for targeted analysis. | native-comparison-api | docx | verified |
| 13 | `compare-a-doc-file-with-a-docx-file-and-verify-that-table-structure-differences-are-detect.cs` | Compare a DOC file with a DOCX file and verify that table structure differences are detected. | native-comparison-api | docx, doc | verified |
| 14 | `compare-documents-with-different-page-orientations-and-verify-that-orientation-changes-app.cs` | Compare documents with different page orientations and verify that orientation changes appear as revisions. | revision-processing | docx | verified |
| 15 | `compare-documents-containing-embedded-charts-and-ensure-that-chart-data-changes-are-captur.cs` | Compare documents containing embedded charts and ensure that chart data changes are captured as revisions. | revision-processing | docx | verified |
| 16 | `compare-a-document-against-itself-to-confirm-that-no-revisions-are-generated.cs` | Compare a document against itself to confirm that no revisions are generated. | revision-processing | docx | verified |
| 17 | `iterate-through-doc1-revisions-collection-and-log-each-revision-type-and-affected-text.cs` | Iterate through doc1.Revisions collection and log each revision type and affected text. | revision-processing | docx | verified |
| 18 | `detect-changes-in-table-cell-formatting-and-log-the-specific-cell-coordinates-for-each-rev.cs` | Detect changes in table cell formatting and log the specific cell coordinates for each revision. | revision-processing | docx | verified |
| 19 | `create-a-custom-logger-that-records-revision-type-author-and-timestamp-for-each-detected-c.cs` | Create a custom logger that records revision type, author, and timestamp for each detected change. | revision-processing | docx | verified |
| 20 | `check-doc1-revisions-count-equals-zero-to-determine-that-two-documents-are-identical-after.cs` | Check doc1.Revisions.Count equals zero to determine that two documents are identical after comparison. | revision-processing | docx | verified |
| 21 | `accept-only-formatting-revisions-while-rejecting-all-content-changes-in-the-compared-docum.cs` | Accept only formatting revisions while rejecting all content changes in the compared document. | revision-processing | docx | verified |
| 22 | `reject-revisions-related-to-header-modifications-and-preserve-footer-changes-in-the-final.cs` | Reject revisions related to header modifications and preserve footer changes in the final output. | revision-processing | docx | verified |
| 23 | `after-comparison-programmatically-accept-all-revisions-and-save-the-cleaned-document-as-do.cs` | After comparison, programmatically accept all revisions and save the cleaned document as DOCX. | revision-processing | docx | verified |
| 24 | `after-accepting-revisions-remove-all-revision-marks-by-clearing-the-revisions-collection.cs` | After accepting revisions, remove all revision marks by clearing the Revisions collection. | revision-processing | docx | verified |
| 25 | `generate-a-summary-report-of-added-deleted-and-modified-paragraphs-using-the-revisions-col.cs` | Generate a summary report of added, deleted, and modified paragraphs using the Revisions collection. | revision-processing | docx | verified |
| 26 | `save-the-compared-document-as-doc-format-while-preserving-all-revision-metadata.cs` | Save the compared document as DOC format while preserving all revision metadata. | revision-processing | docx | verified |
| 27 | `save-the-compared-document-as-docx-format-while-preserving-all-revision-metadata.cs` | Save the compared document as DOCX format while preserving all revision metadata. | revision-processing | docx | verified |
| 28 | `implement-error-handling-to-catch-exceptions-thrown-when-loading-unsupported-file-formats.cs` | Implement error handling to catch exceptions thrown when loading unsupported file formats during comparison. | native-comparison-api | docx | verified |
| 29 | `use-a-disposable-pattern-to-ensure-document-objects-are-properly-released-after-comparison.cs` | Use a disposable pattern to ensure Document objects are properly released after comparison operations. | native-comparison-api | docx | verified |
| 30 | `validate-that-the-comparison-output-matches-the-online-tool-results-by-comparing-revision.cs` | Validate that the comparison output matches the online tool results by comparing revision counts and types. | revision-processing | docx | verified |

## Common failure patterns seen during generation and how they were corrected

### Identical document comparison

- Symptom: Comparison yields no revisions when differences were expected.
- Fix: Create deterministic local source documents with meaningful differences before calling Compare.

### Invented revision reporting

- Symptom: Counts or revision metadata are reported without inspecting the real revisions collection.
- Fix: Use document.Revisions directly and filter actual revision items.

### Missing local source documents

- Symptom: Existing DOCX/DOC/folder/stream inputs are assumed to exist.
- Fix: Bootstrap local versions, folders, or MemoryStream-backed documents in the example.

### Incorrect post-processing assumption

- Symptom: Accept/reject workflows do not verify the final revision state.
- Fix: Check document.Revisions.Count after accepting or rejecting revisions.

### Nullable warnings

- Symptom: CS8600, CS8602, or CS8604 around maybe-null nodes or filtered results.
- Fix: Use nullable locals and null-check before dereference or assignment.

## Notes for maintainers

- This category is now **100% publish-ready** for the current run.
- Preserve file-to-task traceability when updating this folder.
- For future updates, keep the examples standalone and continue bootstrapping local inputs inside the example whenever external sources are mentioned.
