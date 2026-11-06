# Comparison Examples for Aspose.Words for .NET

The Comparison category demonstrates how to use Aspose.Words native document-comparison APIs to detect, process, and export revisions. Verified examples illustrate correct document creation, option configuration, revision inspection, and resource cleanup.

## Snapshot

- Category: Comparison
- Slug: comparison
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261102_031946_08fb5b
- Revision Processing examples: 19
- Stream and batch I/O examples: 2
- Native Comparison API examples: 9

## Category rules that shaped these examples

- Do not reference members that are not present in Aspose.Words 26.10.0, such as CompareOptions.IgnoreWhitespaceChanges or Revision.GetText.
- Do not fabricate revision-type constants that the library does not define.
- Do not assume external files exist; generate sample inputs programmatically.
- Instantiate Document and DocumentBuilder, build distinct content, then call Document.Compare with author and timestamp.
- Use the Revisions collection to count, filter, accept, or reject actual Revision objects.
- Reset MemoryStream.Position to zero before loading a document from a stream.
- Initialize all non-nullable reference-type fields before use.
- Declare locals as nullable when a value may be null and guard them with explicit null checks.
- Avoid dereferencing possibly null Revision.ParentNode without a preceding null test.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.10.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0

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
dotnet add package Aspose.Words --version 26.10.0

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
| 8 | `configure-comparisonoptions-to-ignore-whitespace-changes-when-comparing-source-code-docume.cs` | Configure ComparisonOptions to ignore whitespace changes when comparing source code documentation files. | native-comparison-api | docx | mcp |
| 9 | `enable-detection-of-moved-paragraphs-by-setting-appropriate-flags-in-comparisonoptions-bef.cs` | Enable detection of moved paragraphs by setting appropriate flags in ComparisonOptions before comparison. | native-comparison-api | docx | mcp |
| 10 | `set-comparisontarget-to-new-document-so-that-revisions-appear-in-the-second-file.cs` | Set ComparisonTarget to new document so that revisions appear in the second file. | revision-processing | docx | mcp |
| 11 | `set-comparisonoptions-showdeletedcontent-to-true-to-retain-deleted-text-in-the-comparison.cs` | Set ComparisonOptions.ShowDeletedContent to true to retain deleted text in the comparison output. | native-comparison-api | docx | mcp |
| 12 | `use-comparisonoptions-to-limit-comparison-to-specific-sections-of-the-document-for-targete.cs` | Use ComparisonOptions to limit comparison to specific sections of the document for targeted analysis. | native-comparison-api | docx | mcp |
| 13 | `compare-a-doc-file-with-a-docx-file-and-verify-that-table-structure-differences-are-detect.cs` | Compare a DOC file with a DOCX file and verify that table structure differences are detected. | native-comparison-api | docx, doc | mcp |
| 14 | `compare-documents-with-different-page-orientations-and-verify-that-orientation-changes-app.cs` | Compare documents with different page orientations and verify that orientation changes appear as revisions. | revision-processing | docx | existing_repo |
| 15 | `compare-documents-containing-embedded-charts-and-ensure-that-chart-data-changes-are-captur.cs` | Compare documents containing embedded charts and ensure that chart data changes are captured as revisions. | revision-processing | docx | llm |
| 16 | `compare-a-document-against-itself-to-confirm-that-no-revisions-are-generated.cs` | Compare a document against itself to confirm that no revisions are generated. | revision-processing | docx | mcp |
| 17 | `iterate-through-doc1-revisions-collection-and-log-each-revision-type-and-affected-text.cs` | Iterate through doc1.Revisions collection and log each revision type and affected text. | revision-processing | docx | mcp |
| 18 | `detect-changes-in-table-cell-formatting-and-log-the-specific-cell-coordinates-for-each-rev.cs` | Detect changes in table cell formatting and log the specific cell coordinates for each revision. | revision-processing | docx | mcp |
| 19 | `create-a-custom-logger-that-records-revision-type-author-and-timestamp-for-each-detected-c.cs` | Create a custom logger that records revision type, author, and timestamp for each detected change. | revision-processing | docx | mcp |
| 20 | `check-doc1-revisions-count-equals-zero-to-determine-that-two-documents-are-identical-after.cs` | Check doc1.Revisions.Count equals zero to determine that two documents are identical after comparison. | revision-processing | docx | mcp |
| 21 | `accept-only-formatting-revisions-while-rejecting-all-content-changes-in-the-compared-docum.cs` | Accept only formatting revisions while rejecting all content changes in the compared document. | revision-processing | docx | mcp |
| 22 | `reject-revisions-related-to-header-modifications-and-preserve-footer-changes-in-the-final.cs` | Reject revisions related to header modifications and preserve footer changes in the final output. | revision-processing | docx | mcp |
| 23 | `after-comparison-programmatically-accept-all-revisions-and-save-the-cleaned-document-as-do.cs` | After comparison, programmatically accept all revisions and save the cleaned document as DOCX. | revision-processing | docx | mcp |
| 24 | `after-accepting-revisions-remove-all-revision-marks-by-clearing-the-revisions-collection.cs` | After accepting revisions, remove all revision marks by clearing the Revisions collection. | revision-processing | docx | mcp |
| 25 | `generate-a-summary-report-of-added-deleted-and-modified-paragraphs-using-the-revisions-col.cs` | Generate a summary report of added, deleted, and modified paragraphs using the Revisions collection. | revision-processing | docx | mcp |
| 26 | `save-the-compared-document-as-doc-format-while-preserving-all-revision-metadata.cs` | Save the compared document as DOC format while preserving all revision metadata. | revision-processing | docx | mcp |
| 27 | `save-the-compared-document-as-docx-format-while-preserving-all-revision-metadata.cs` | Save the compared document as DOCX format while preserving all revision metadata. | revision-processing | docx | mcp |
| 28 | `implement-error-handling-to-catch-exceptions-thrown-when-loading-unsupported-file-formats.cs` | Implement error handling to catch exceptions thrown when loading unsupported file formats during comparison. | native-comparison-api | docx | mcp |
| 29 | `use-a-disposable-pattern-to-ensure-document-objects-are-properly-released-after-comparison.cs` | Use a disposable pattern to ensure Document objects are properly released after comparison operations. | native-comparison-api | docx | mcp |
| 30 | `validate-that-the-comparison-output-matches-the-online-tool-results-by-comparing-revision.cs` | Validate that the comparison output matches the online tool results by comparing revision counts and types. | revision-processing | docx | llm |

## Common failure patterns seen during generation and how they were corrected

### Use of non-existent API members

- Symptom: The compiler reports CS1061 or CS0117 because a property or flag such as CompareOptions.IgnoreWhitespaceChanges does not exist.
- Root cause: The example was written against a newer or different version of Aspose.Words where those members are defined.
- Preferred fix: Remove the unsupported member and rely on the available options, for example set CompareOptions.IgnoreFormatting to treat whitespace as formatting.
- Prevention: Consult the official API reference for version 26.10.0 before using any CompareOptions flags.

### Assumed external documents

- Symptom: Runtime code fails with FileNotFoundException or the comparison produces zero revisions because the input files were never created.
- Root cause: The sample assumes pre-existing DOCX/DOC files or folder structures.
- Preferred fix: Generate the required source documents in the example code using DocumentBuilder before invoking Compare.
- Prevention: Always bootstrap sample inputs locally; never rely on external files that are not part of the snippet.

### Missing revision members

- Symptom: Compilation errors such as CS1061 for Revision.GetText or Revision.Node appear.
- Root cause: The code accesses members that are not part of the Revision class in the target library.
- Preferred fix: Use Revision.ParentNode to obtain the affected node and call GetText on that node instead of on Revision.
- Prevention: Reference the Revision documentation and use only the properties and methods that exist.

### Incorrect collection manipulation

- Symptom: CS1061 errors arise when calling Clear on RevisionCollection or CopyTo with an unsupported overload.
- Root cause: The example uses collection methods that the library does not implement.
- Preferred fix: Iterate over the collection and remove items individually if needed, or simply accept all revisions without clearing the collection.
- Prevention: Check the RevisionCollection API for supported operations before invoking them.

### Nullable reference misuse

- Symptom: Compiler warnings CS8600, CS8602, or CS8604 appear when a possibly null value is assigned or dereferenced.
- Root cause: The code does not guard against null results from properties such as Revision.ParentNode.
- Preferred fix: Introduce nullable locals and perform explicit null checks before accessing members.
- Prevention: Enable nullable reference analysis and follow the nullable_warning_prevention guidelines.

### Incorrect disposal pattern

- Symptom: CS1674 indicates that Document does not implement IDisposable, yet it is used in a using statement.
- Root cause: The example assumes Document implements IDisposable, which it does not in this version.
- Preferred fix: Do not place Document in a using block; rely on the library's internal resource management.
- Prevention: Wrap only objects that implement IDisposable (e.g., MemoryStream) in using statements.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Update the rules summary whenever a new Aspose.Words version introduces additional CompareOptions flags or revision members.
