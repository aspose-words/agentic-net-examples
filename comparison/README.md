# Comparison Examples for Aspose.Words for .NET

The Comparison category demonstrates how to use Aspose.Words native document-comparison APIs to detect, process, and report revisions across a variety of document types. Verified examples cover author attribution, case-insensitive comparison, whitespace-ignoring comparison, moved-paragraph detection, revision-target selection, deleted-content retention, orientation changes, chart revisions, revision enumeration, and table-cell formatting analysis.

## Snapshot

- Category: Comparison
- Slug: comparison
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261124_190740_1c00f5
- Revision Processing examples: 19
- Stream and batch I/O examples: 2
- Native Comparison API examples: 9

## Category rules that shaped these examples

- Do not reference Aspose.Words members that are not present in version 26.11.0.
- Do not invent comparison options, revision members, or enum values that are absent from the official API.
- Use Document.Compare with a real author name and comparison timestamp.
- Create meaningful differences between local source documents before comparing them.
- Inspect and process actual document.Revisions items for filtering, counting, acceptance, or rejection.
- Initialize all non-nullable reference-type properties to avoid CS8618 warnings.
- Guard maybe-null values with explicit null checks before dereferencing them.
- Declare locals as nullable when a value may legitimately be null and verify them before use.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.11.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0

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
dotnet add package Aspose.Words --version 26.11.0

# PowerShell example
Copy-Item ..\comparison\load-two-docx-files-compare-them-and-save-the-resulting-document-with-revisions-applied.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-two-docx-files-compare-them-and-save-the-resulting-document-with-revisions-applied.cs` | Load two DOCX files, compare them, and save the resulting document with revisions applied. | revision-processing | docx | mcp |
| 2 | `load-a-doc-file-compare-it-with-another-doc-and-save-the-result-with-revisions-included.cs` | Load a DOC file, compare it with another DOC, and save the result with revisions included. | revision-processing | doc | mcp |
| 3 | `use-a-memorystream-to-load-documents-perform-comparison-and-write-the-result-to-a-byte-arr.cs` | Use a MemoryStream to load documents, perform comparison, and write the result to a byte array. | stream-batch-io | docx | mcp |
| 4 | `batch-process-a-folder-of-document-pairs-generating-compared-versions-with-revision-tracki.cs` | Batch process a folder of document pairs, generating compared versions with revision tracking for each pair. | stream-batch-io | docx | mcp |
| 5 | `apply-custom-author-name-and-timestamp-when-calling-compare-to-attribute-revisions-to-a-sp.cs` | Apply custom author name and timestamp when calling Compare to attribute revisions to a specific user. | revision-processing | docx | mcp |
| 6 | `configure-comparisonoptions-to-ignore-formatting-changes-before-performing-document-compar.cs` | Configure ComparisonOptions to ignore formatting changes before performing document comparison on large legal contracts. | native-comparison-api | docx | mcp |
| 7 | `configure-comparisonoptions-to-ignore-case-differences-when-comparing-textual-content-acro.cs` | Configure ComparisonOptions to ignore case differences when comparing textual content across documents. | native-comparison-api | docx | mcp |
| 8 | `configure-comparisonoptions-to-ignore-whitespace-changes-when-comparing-source-code-docume.cs` | Configure ComparisonOptions to ignore whitespace changes when comparing source code documentation files. | native-comparison-api | docx | llm |
| 9 | `enable-detection-of-moved-paragraphs-by-setting-appropriate-flags-in-comparisonoptions-bef.cs` | Enable detection of moved paragraphs by setting appropriate flags in ComparisonOptions before comparison. | native-comparison-api | docx | mcp |
| 10 | `set-comparisontarget-to-new-document-so-that-revisions-appear-in-the-second-file.cs` | Set ComparisonTarget to new document so that revisions appear in the second file. | revision-processing | docx | mcp |
| 11 | `set-comparisonoptions-showdeletedcontent-to-true-to-retain-deleted-text-in-the-comparison.cs` | Set ComparisonOptions.ShowDeletedContent to true to retain deleted text in the comparison output. | native-comparison-api | docx | mcp |
| 12 | `use-comparisonoptions-to-limit-comparison-to-specific-sections-of-the-document-for-targete.cs` | Use ComparisonOptions to limit comparison to specific sections of the document for targeted analysis. | native-comparison-api | docx | mcp |
| 13 | `compare-a-doc-file-with-a-docx-file-and-verify-that-table-structure-differences-are-detect.cs` | Compare a DOC file with a DOCX file and verify that table structure differences are detected. | native-comparison-api | docx, doc | mcp |
| 14 | `compare-documents-with-different-page-orientations-and-verify-that-orientation-changes-app.cs` | Compare documents with different page orientations and verify that orientation changes appear as revisions. | revision-processing | docx | llm |
| 15 | `compare-documents-containing-embedded-charts-and-ensure-that-chart-data-changes-are-captur.cs` | Compare documents containing embedded charts and ensure that chart data changes are captured as revisions. | revision-processing | docx | mcp |
| 16 | `compare-a-document-against-itself-to-confirm-that-no-revisions-are-generated.cs` | Compare a document against itself to confirm that no revisions are generated. | revision-processing | docx | mcp |
| 17 | `iterate-through-doc1-revisions-collection-and-log-each-revision-type-and-affected-text.cs` | Iterate through doc1.Revisions collection and log each revision type and affected text. | revision-processing | docx | mcp |
| 18 | `detect-changes-in-table-cell-formatting-and-log-the-specific-cell-coordinates-for-each-rev.cs` | Detect changes in table cell formatting and log the specific cell coordinates for each revision. | revision-processing | docx | existing_repo |
| 19 | `create-a-custom-logger-that-records-revision-type-author-and-timestamp-for-each-detected-c.cs` | Create a custom logger that records revision type, author, and timestamp for each detected change. | revision-processing | docx | mcp |
| 20 | `check-doc1-revisions-count-equals-zero-to-determine-that-two-documents-are-identical-after.cs` | Check doc1.Revisions.Count equals zero to determine that two documents are identical after comparison. | revision-processing | docx | mcp |
| 21 | `accept-only-formatting-revisions-while-rejecting-all-content-changes-in-the-compared-docum.cs` | Accept only formatting revisions while rejecting all content changes in the compared document. | revision-processing | docx | llm |
| 22 | `reject-revisions-related-to-header-modifications-and-preserve-footer-changes-in-the-final.cs` | Reject revisions related to header modifications and preserve footer changes in the final output. | revision-processing | docx | mcp |
| 23 | `after-comparison-programmatically-accept-all-revisions-and-save-the-cleaned-document-as-do.cs` | After comparison, programmatically accept all revisions and save the cleaned document as DOCX. | revision-processing | docx | mcp |
| 24 | `after-accepting-revisions-remove-all-revision-marks-by-clearing-the-revisions-collection.cs` | After accepting revisions, remove all revision marks by clearing the Revisions collection. | revision-processing | docx | mcp |
| 25 | `generate-a-summary-report-of-added-deleted-and-modified-paragraphs-using-the-revisions-col.cs` | Generate a summary report of added, deleted, and modified paragraphs using the Revisions collection. | revision-processing | docx | llm |
| 26 | `save-the-compared-document-as-doc-format-while-preserving-all-revision-metadata.cs` | Save the compared document as DOC format while preserving all revision metadata. | revision-processing | docx | mcp |
| 27 | `save-the-compared-document-as-docx-format-while-preserving-all-revision-metadata.cs` | Save the compared document as DOCX format while preserving all revision metadata. | revision-processing | docx | mcp |
| 28 | `implement-error-handling-to-catch-exceptions-thrown-when-loading-unsupported-file-formats.cs` | Implement error handling to catch exceptions thrown when loading unsupported file formats during comparison. | native-comparison-api | docx | llm |
| 29 | `use-a-disposable-pattern-to-ensure-document-objects-are-properly-released-after-comparison.cs` | Use a disposable pattern to ensure Document objects are properly released after comparison operations. | native-comparison-api | docx | mcp |
| 30 | `validate-that-the-comparison-output-matches-the-online-tool-results-by-comparing-revision.cs` | Validate that the comparison output matches the online tool results by comparing revision counts and types. | revision-processing | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Use of non-existent CompareOptions members

- Symptom: The compiler reports CS0117 errors for properties such as IgnoreCase, IgnoreWhitespace, ShowDeletedContent, DetectMove, DetectMovedParagraphs, and ComparisonTarget.
- Root cause: The code references CompareOptions members that were removed or never existed in the current Aspose.Words release.
- Preferred fix: Reference only the properties documented for CompareOptions in version 26.11.0, such as IgnoreCaseChanges, IgnoreFormatting, or use reflection to set optional flags when they are available.
- Prevention: Consult the official API reference for the exact property names and guard reflective access with version checks.

### Invocation of undefined Revision members

- Symptom: CS1061 errors appear for Revision.GetText, Revision.RevisionDate, Revision.RevisionDateTime, and missing RevisionCollection.Clear.
- Root cause: The code attempts to call members that are not part of the Revision class in the installed library.
- Preferred fix: Use Revision.ParentNode?.GetText() to obtain affected text and Revision.DateTime for timestamps; do not call Clear on RevisionCollection because it is read-only.
- Prevention: Access only the members listed in the Revision documentation and treat the Revisions collection as immutable.

### Incorrect disposal of Document objects

- Symptom: CS1674 errors indicate that Document cannot be used in a using statement because it does not implement IDisposable.
- Root cause: The code assumes Document implements IDisposable, which it does not in Aspose.Words.
- Preferred fix: Create Document instances without a using block and rely on the garbage collector, or call Document.Dispose only if a future version adds it.
- Prevention: Check the type hierarchy before applying the disposable pattern.

### Missing exception types for loading errors

- Symptom: CS0234 and CS0246 errors show that LoadingException, FileFormatException, and LoadFormatException are not defined.
- Root cause: The code imports exception types that belong to other libraries or older Aspose.Words versions.
- Preferred fix: Catch the generic Aspose.Words.LoadFormatException or System.IO.IOException as appropriate.
- Prevention: Reference the exception types that appear in the current API documentation.

### Assuming revisions without inspecting the collection

- Symptom: Logic checks revision counts or types based on hard-coded values rather than enumerating Document.Revisions.
- Root cause: The example code reports expected revision numbers without verifying them.
- Preferred fix: Iterate Document.Revisions and count or filter actual items to determine the outcome.
- Prevention: Always base decisions on the live Revisions collection.

### Comparing identical documents without introducing differences

- Symptom: The comparison produces zero revisions even though the task expects changes.
- Root cause: Both source documents contain the same content.
- Preferred fix: Introduce deterministic differences such as altered text, formatting, or structural changes before calling Compare.
- Prevention: Create sample inputs that deliberately differ in the aspect under test.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Update the example code whenever a new Aspose.Words release changes the CompareOptions or Revision API surface.
