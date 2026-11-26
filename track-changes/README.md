# Track Changes Examples for Aspose.Words for .NET

The Track Changes category demonstrates how to use Aspose.Words 26.11.0 to create, modify, and manage revision data in Word documents. Verified examples show document creation, enabling revision tracking, performing edits such as paragraph insertion, style changes, table insertion, and then inspecting, accepting, or rejecting the generated revisions.

## Snapshot

- Category: Track Changes
- Slug: track-changes
- Total examples: 40
- Publish-ready successful examples: 40 / 40
- Source run: 20261124_190740_1c00f5
- Track Changes Workflow examples: 40

## Category rules that shaped these examples

- Do not reference members that are not present in Aspose.Words 26.11.0.
- Do not import or use fictitious namespaces or classes.
- Do not call revision-related methods that are not part of the official API surface.
- Use Document.StartTrackRevisions and Document.StopTrackRevisions to control revision capture.
- Inspect revisions through Document.Revisions and, when needed, group information via RevisionCollection.Groups.
- Accept or reject revisions with Revision.Accept, Revision.Reject, Document.AcceptAllRevisions, or Document.RejectAllRevisions.
- Use the HasRevisions property to decide whether additional actions such as watermark insertion are required.
- Create any source files, images, or folders in the sample code before loading them.
- Validate that a document was successfully loaded before accessing its nodes.
- Always call Document.Save after the requested modifications to guarantee output existence.

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
Copy-Item ..\track-changes\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `track-changes/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0

# PowerShell example
Copy-Item ..\track-changes\load-a-document-from-a-stream-start-tracking-add-a-header-and-stop-tracking-before-saving.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-document-from-a-stream-start-tracking-add-a-header-and-stop-tracking-before-saving.cs` | Load a document from a stream, start tracking, add a header, and stop tracking before saving. | track-changes-workflow | doc | mcp |
| 2 | `create-a-new-document-start-tracking-revisions-insert-a-paragraph-then-stop-tracking.cs` | Create a new document, start tracking revisions, insert a paragraph, then stop tracking. | track-changes-workflow | doc | mcp |
| 3 | `enable-revision-tracking-modify-a-table-cell-and-verify-the-revision-appears-in-the-collec.cs` | Enable revision tracking, modify a table cell, and verify the revision appears in the collection. | track-changes-workflow | docx | mcp |
| 4 | `enable-tracking-apply-a-style-change-to-multiple-paragraphs-stop-tracking-and-verify-a-sin.cs` | Enable tracking, apply a style change to multiple paragraphs, stop tracking, and verify a single revision group. | track-changes-workflow | docx | existing_repo |
| 5 | `start-tracking-revisions-insert-an-image-stop-tracking-and-verify-the-image-insertion-appe.cs` | Start tracking revisions, insert an image, stop tracking, and verify the image insertion appears as a revision. | track-changes-workflow | docx | mcp |
| 6 | `load-a-document-start-tracking-insert-a-table-stop-tracking-then-accept-the-table-insertio.cs` | Load a document, start tracking, insert a table, stop tracking, then accept the table insertion revision. | track-changes-workflow | doc | llm |
| 7 | `start-tracking-delete-a-sentence-stop-tracking-then-accept-the-deletion-revision-individua.cs` | Start tracking, delete a sentence, stop tracking, then accept the deletion revision individually. | track-changes-workflow | docx | mcp |
| 8 | `load-a-document-enable-tracking-perform-a-find-and-replace-operation-and-list-the-generate.cs` | Load a document, enable tracking, perform a find-and-replace operation, and list the generated revisions. | track-changes-workflow | doc | mcp |
| 9 | `load-a-document-enable-revision-tracking-apply-formatting-changes-and-list-resulting-revis.cs` | Load a document, enable revision tracking, apply formatting changes, and list resulting revision types. | track-changes-workflow | doc | mcp |
| 10 | `stop-tracking-revisions-then-make-additional-edits-and-confirm-no-new-revisions-are-record.cs` | Stop tracking revisions, then make additional edits and confirm no new revisions are recorded. | track-changes-workflow | docx | existing_repo |
| 11 | `determine-if-a-document-contains-any-revisions-before-further-processing-by-checking-the-h.cs` | Determine if a document contains any revisions before further processing by checking the HasRevisions property. | track-changes-workflow | doc | mcp |
| 12 | `use-the-hasrevisions-property-to-conditionally-apply-a-watermark-indicating-pending-change.cs` | Use the HasRevisions property to conditionally apply a watermark indicating pending changes. | track-changes-workflow | docx | mcp |
| 13 | `iterate-through-all-revisions-in-the-document-and-log-each-revision-s-author-and-timestamp.cs` | Iterate through all revisions in the document and log each revision's author and timestamp. | track-changes-workflow | doc | mcp |
| 14 | `iterate-over-revisions-group-them-by-type-and-output-a-summary-of-insertion-and-deletion-c.cs` | Iterate over revisions, group them by type, and output a summary of insertion and deletion counts. | track-changes-workflow | docx | mcp |
| 15 | `group-sequential-insertions-using-revisiongroup-and-accept-the-entire-group-with-a-single.cs` | Group sequential insertions using RevisionGroup and accept the entire group with a single call. | track-changes-workflow | docx | llm |
| 16 | `use-revisiongroup-to-merge-consecutive-deletions-into-a-single-revision-before-accepting-i.cs` | Use RevisionGroup to merge consecutive deletions into a single revision before accepting it. | track-changes-workflow | docx | llm |
| 17 | `accept-revisions-only-for-paragraphs-authored-by-a-specific-user-while-rejecting-others.cs` | Accept revisions only for paragraphs authored by a specific user while rejecting others. | track-changes-workflow | docx | mcp |
| 18 | `reject-a-specific-revision-based-on-author-name-then-confirm-the-change-is-no-longer-prese.cs` | Reject a specific revision based on author name, then confirm the change is no longer present. | track-changes-workflow | docx | mcp |
| 19 | `reject-all-formatting-revisions-but-keep-content-insertions-and-deletions-still-intact.cs` | Reject all formatting revisions but keep content insertions and deletions still intact. | track-changes-workflow | docx | mcp |
| 20 | `programmatically-reject-all-deletions-while-keeping-insertions-and-formatting-changes-stil.cs` | Programmatically reject all deletions while keeping insertions and formatting changes still intact. | track-changes-workflow | docx | llm |
| 21 | `reject-revisions-older-than-a-specific-date-while-preserving-newer-changes-in-the-document.cs` | Reject revisions older than a specific date while preserving newer changes in the document. | track-changes-workflow | doc | mcp |
| 22 | `accept-all-revisions-in-a-loaded-docx-file-and-save-the-cleaned-document-to-a-new-file.cs` | Accept all revisions in a loaded DOCX file and save the cleaned document to a new file. | track-changes-workflow | docx, doc | mcp |
| 23 | `accept-all-revisions-in-a-document-then-re-enable-tracking-to-capture-subsequent-changes-s.cs` | Accept all revisions in a document, then re-enable tracking to capture subsequent changes separately. | track-changes-workflow | doc | mcp |
| 24 | `after-accepting-all-revisions-lock-the-document-to-prevent-further-editing-without-trackin.cs` | After accepting all revisions, lock the document to prevent further editing without tracking. | track-changes-workflow | doc | llm |
| 25 | `reject-all-tracked-changes-in-a-document-then-verify-the-hasrevisions-property-returns-fal.cs` | Reject all tracked changes in a document, then verify the HasRevisions property returns false. | track-changes-workflow | doc | mcp |
| 26 | `compare-two-word-documents-to-generate-revision-differences-specifying-author-name-and-com.cs` | Compare two Word documents to generate revision differences, specifying author name and comparison date. | track-changes-workflow | doc | mcp |
| 27 | `use-doc-compare-to-create-a-revision-document-that-highlights-differences-between-two-vers.cs` | Use doc.Compare to create a revision document that highlights differences between two versions. | track-changes-workflow | doc | mcp |
| 28 | `programmatically-compare-a-document-against-its-previous-version-and-generate-a-revision-r.cs` | Programmatically compare a document against its previous version and generate a revision report in memory. | track-changes-workflow | doc | existing_repo |
| 29 | `compare-three-versions-of-a-contract-sequentially-generating-cumulative-revision-sets-for.cs` | Compare three versions of a contract sequentially, generating cumulative revision sets for each comparison. | track-changes-workflow | docx | mcp |
| 30 | `generate-a-report-listing-each-revision-s-type-author-and-the-paragraph-number-where-it-oc.cs` | Generate a report listing each revision's type, author, and the paragraph number where it occurs. | track-changes-workflow | docx | mcp |
| 31 | `develop-a-utility-to-export-revision-metadata-to-a-csv-file-for-external-analysis.cs` | Develop a utility to export revision metadata to a CSV file for external analysis. | track-changes-workflow | csv | mcp |
| 32 | `create-a-custom-logger-that-captures-revision-metadata-during-document-editing-for-audit-p.cs` | Create a custom logger that captures revision metadata during document editing for audit purposes. | track-changes-workflow | doc | mcp |
| 33 | `create-a-function-that-returns-true-if-any-revision-author-matches-a-specified-list-of-nam.cs` | Create a function that returns true if any revision author matches a specified list of names. | track-changes-workflow | docx | mcp |
| 34 | `create-a-utility-that-accepts-revisions-only-when-they-meet-a-minimum-word-count-threshold.cs` | Create a utility that accepts revisions only when they meet a minimum word count threshold. | track-changes-workflow | docx | llm |
| 35 | `batch-process-a-folder-of-documents-rejecting-revisions-authored-by-a-given-user-across-al.cs` | Batch process a folder of documents, rejecting revisions authored by a given user across all files. | track-changes-workflow | doc | mcp |
| 36 | `create-a-batch-process-that-opens-multiple-documents-accepts-all-revisions-and-saves-them.cs` | Create a batch process that opens multiple documents, accepts all revisions, and saves them in place. | track-changes-workflow | doc | mcp |
| 37 | `accept-revisions-in-a-specific-section-of-the-document-while-leaving-other-sections-unchan.cs` | Accept revisions in a specific section of the document while leaving other sections unchanged. | track-changes-workflow | doc | mcp |
| 38 | `after-rejecting-all-revisions-verify-that-the-document-s-original-content-matches-a-baseli.cs` | After rejecting all revisions, verify that the document's original content matches a baseline file. | track-changes-workflow | doc | mcp |
| 39 | `implement-error-handling-for-attempts-to-accept-a-revision-that-has-already-been-rejected.cs` | Implement error handling for attempts to accept a revision that has already been rejected. | track-changes-workflow | docx | mcp |
| 40 | `create-a-script-that-toggles-tracking-on-and-off-based-on-user-input-during-runtime.cs` | Create a script that toggles tracking on and off based on user input during runtime. | track-changes-workflow | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Unsupported API usage

- Symptom: Compilation fails because types such as RevisionGroups, RevisionDate, or RevisionNode are not found.
- Root cause: The code references revision members that were never implemented in the selected Aspose.Words version.
- Preferred fix: Replace the nonexistent members with the documented Revision, RevisionCollection, and RevisionType APIs.
- Prevention: Consult the official Aspose.Words 26.11.0 reference before using any revision-related type or property.

### Incorrect API target

- Symptom: Compilation errors indicate that methods like StartTrackRevisions or StopTrackRevisions are missing on DocumentBuilder.
- Root cause: Revision tracking methods were called on DocumentBuilder instead of the Document instance.
- Preferred fix: Invoke StartTrackRevisions and StopTrackRevisions on the Document object, not on DocumentBuilder.
- Prevention: Remember that tracking is a document-level feature; only Document exposes the tracking methods.

### Missing revision generation

- Symptom: Runtime checks report that zero or an unexpected number of revisions were produced.
- Root cause: Edits were performed while tracking was disabled or before StartTrackRevisions was called.
- Preferred fix: Enable tracking before making any modifications and ensure StopTrackRevisions is called only after the edits are complete.
- Prevention: Structure the code so that all tracked edits occur between StartTrackRevisions and StopTrackRevisions.

### Collection mutation during enumeration

- Symptom: Runtime exception occurs when a revision is accepted or rejected inside a foreach loop over Document.Revisions.
- Root cause: Accept/Reject modifies the RevisionCollection while it is being enumerated.
- Preferred fix: Materialize the collection into a list (e.g., .ToList()) before iterating, or iterate by index.
- Prevention: Never modify Document.Revisions directly inside a foreach; use a snapshot or index-based loop.

### Invalid protection call

- Symptom: Compilation or runtime errors arise when calling Document.Protect with ProtectionType.TrackedChanges or mismatched arguments.
- Root cause: The Protect overload was used with an unsupported ProtectionType value or incorrect parameter order.
- Preferred fix: Use a supported ProtectionType such as ProtectionType.ReadOnly and pass the correct password argument.
- Prevention: Refer to the Protect method signature in the official documentation and pass only valid enum values.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Periodically review the Aspose.Words release notes to ensure that all example code matches the current revision API surface.
