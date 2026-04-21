---
name: track-changes
description: Verified C# examples for revision and track changes workflows in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Track Changes

## Purpose

This folder is a **live, curated example set** for revision and track changes scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free use of documented Aspose.Words APIs for enabling tracking, generating revisions, inspecting revisions, and accepting or rejecting changes.

## Non-negotiable conventions

- Always use documented Aspose.Words APIs directly.
- Always create local sample source documents when a task refers to an existing file, folder, stream, template, or input asset.
- Prefer `StartTrackRevisions`, `StopTrackRevisions`, `Document.Revisions`, and documented accept/reject workflows.
- Keep validation narrow and task-specific.
- Do not invent revision helper APIs.

## Recommended workflow selection

- **Track changes workflow**: 40 examples

This category performed best with light primary rules and direct documented revision workflows.

## Validation priorities

1. The code must compile and run without manual input.
2. Required sample inputs must be bootstrapped locally inside the example.
3. Requested revision or output results must be produced successfully.
4. Validation should focus only on the exact requested revision count, type, author, or acceptance/rejection outcome.

## File-to-task reference

- `load-a-document-from-a-stream-start-tracking-add-a-header-and-stop-tracking-before-saving.cs`
  - Task: Load a document from a stream, start tracking, add a header, and stop tracking before saving.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-new-document-start-tracking-revisions-insert-a-paragraph-then-stop-tracking.cs`
  - Task: Create a new document, start tracking revisions, insert a paragraph, then stop tracking.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `enable-revision-tracking-modify-a-table-cell-and-verify-the-revision-appears-in-the-collec.cs`
  - Task: Enable revision tracking, modify a table cell, and verify the revision appears in the collection.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `enable-tracking-apply-a-style-change-to-multiple-paragraphs-stop-tracking-and-verify-a-sin.cs`
  - Task: Enable tracking, apply a style change to multiple paragraphs, stop tracking, and verify a single revision group.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `start-tracking-revisions-insert-an-image-stop-tracking-and-verify-the-image-insertion-appe.cs`
  - Task: Start tracking revisions, insert an image, stop tracking, and verify the image insertion appears as a revision.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `load-a-document-start-tracking-insert-a-table-stop-tracking-then-accept-the-table-insertio.cs`
  - Task: Load a document, start tracking, insert a table, stop tracking, then accept the table insertion revision.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `start-tracking-delete-a-sentence-stop-tracking-then-accept-the-deletion-revision-individua.cs`
  - Task: Start tracking, delete a sentence, stop tracking, then accept the deletion revision individually.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `load-a-document-enable-tracking-perform-a-find-and-replace-operation-and-list-the-generate.cs`
  - Task: Load a document, enable tracking, perform a find-and-replace operation, and list the generated revisions.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `load-a-document-enable-revision-tracking-apply-formatting-changes-and-list-resulting-revis.cs`
  - Task: Load a document, enable revision tracking, apply formatting changes, and list resulting revision types.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `stop-tracking-revisions-then-make-additional-edits-and-confirm-no-new-revisions-are-record.cs`
  - Task: Stop tracking revisions, then make additional edits and confirm no new revisions are recorded.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `determine-if-a-document-contains-any-revisions-before-further-processing-by-checking-the-h.cs`
  - Task: Determine if a document contains any revisions before further processing by checking the HasRevisions property.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-the-hasrevisions-property-to-conditionally-apply-a-watermark-indicating-pending-change.cs`
  - Task: Use the HasRevisions property to conditionally apply a watermark indicating pending changes.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `iterate-through-all-revisions-in-the-document-and-log-each-revision-s-author-and-timestamp.cs`
  - Task: Iterate through all revisions in the document and log each revision's author and timestamp.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `iterate-over-revisions-group-them-by-type-and-output-a-summary-of-insertion-and-deletion-c.cs`
  - Task: Iterate over revisions, group them by type, and output a summary of insertion and deletion counts.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `group-sequential-insertions-using-revisiongroup-and-accept-the-entire-group-with-a-single.cs`
  - Task: Group sequential insertions using RevisionGroup and accept the entire group with a single call.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `use-revisiongroup-to-merge-consecutive-deletions-into-a-single-revision-before-accepting-i.cs`
  - Task: Use RevisionGroup to merge consecutive deletions into a single revision before accepting it.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `accept-revisions-only-for-paragraphs-authored-by-a-specific-user-while-rejecting-others.cs`
  - Task: Accept revisions only for paragraphs authored by a specific user while rejecting others.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `reject-a-specific-revision-based-on-author-name-then-confirm-the-change-is-no-longer-prese.cs`
  - Task: Reject a specific revision based on author name, then confirm the change is no longer present.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `reject-all-formatting-revisions-but-keep-content-insertions-and-deletions-still-intact.cs`
  - Task: Reject all formatting revisions but keep content insertions and deletions still intact.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-reject-all-deletions-while-keeping-insertions-and-formatting-changes-stil.cs`
  - Task: Programmatically reject all deletions while keeping insertions and formatting changes still intact.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `reject-revisions-older-than-a-specific-date-while-preserving-newer-changes-in-the-document.cs`
  - Task: Reject revisions older than a specific date while preserving newer changes in the document.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `accept-all-revisions-in-a-loaded-docx-file-and-save-the-cleaned-document-to-a-new-file.cs`
  - Task: Accept all revisions in a loaded DOCX file and save the cleaned document to a new file.
  - Workflow: track-changes-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `accept-all-revisions-in-a-document-then-re-enable-tracking-to-capture-subsequent-changes-s.cs`
  - Task: Accept all revisions in a document, then re-enable tracking to capture subsequent changes separately.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `after-accepting-all-revisions-lock-the-document-to-prevent-further-editing-without-trackin.cs`
  - Task: After accepting all revisions, lock the document to prevent further editing without tracking.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `reject-all-tracked-changes-in-a-document-then-verify-the-hasrevisions-property-returns-fal.cs`
  - Task: Reject all tracked changes in a document, then verify the HasRevisions property returns false.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `compare-two-word-documents-to-generate-revision-differences-specifying-author-name-and-com.cs`
  - Task: Compare two Word documents to generate revision differences, specifying author name and comparison date.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-doc-compare-to-create-a-revision-document-that-highlights-differences-between-two-vers.cs`
  - Task: Use doc.Compare to create a revision document that highlights differences between two versions.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `programmatically-compare-a-document-against-its-previous-version-and-generate-a-revision-r.cs`
  - Task: Programmatically compare a document against its previous version and generate a revision report in memory.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `compare-three-versions-of-a-contract-sequentially-generating-cumulative-revision-sets-for.cs`
  - Task: Compare three versions of a contract sequentially, generating cumulative revision sets for each comparison.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `generate-a-report-listing-each-revision-s-type-author-and-the-paragraph-number-where-it-oc.cs`
  - Task: Generate a report listing each revision's type, author, and the paragraph number where it occurs.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `develop-a-utility-to-export-revision-metadata-to-a-csv-file-for-external-analysis.cs`
  - Task: Develop a utility to export revision metadata to a CSV file for external analysis.
  - Workflow: track-changes-workflow
  - Outputs: csv
  - Selected engine: mcp
- `create-a-custom-logger-that-captures-revision-metadata-during-document-editing-for-audit-p.cs`
  - Task: Create a custom logger that captures revision metadata during document editing for audit purposes.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-function-that-returns-true-if-any-revision-author-matches-a-specified-list-of-nam.cs`
  - Task: Create a function that returns true if any revision author matches a specified list of names.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-utility-that-accepts-revisions-only-when-they-meet-a-minimum-word-count-threshold.cs`
  - Task: Create a utility that accepts revisions only when they meet a minimum word count threshold.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `batch-process-a-folder-of-documents-rejecting-revisions-authored-by-a-given-user-across-al.cs`
  - Task: Batch process a folder of documents, rejecting revisions authored by a given user across all files.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-batch-process-that-opens-multiple-documents-accepts-all-revisions-and-saves-them.cs`
  - Task: Create a batch process that opens multiple documents, accepts all revisions, and saves them in place.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `accept-revisions-in-a-specific-section-of-the-document-while-leaving-other-sections-unchan.cs`
  - Task: Accept revisions in a specific section of the document while leaving other sections unchanged.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `after-rejecting-all-revisions-verify-that-the-document-s-original-content-matches-a-baseli.cs`
  - Task: After rejecting all revisions, verify that the document's original content matches a baseline file.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: mcp
- `implement-error-handling-for-attempts-to-accept-a-revision-that-has-already-been-rejected.cs`
  - Task: Implement error handling for attempts to accept a revision that has already been rejected.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-script-that-toggles-tracking-on-and-off-based-on-user-input-during-runtime.cs`
  - Task: Create a script that toggles tracking on and off based on user input during runtime.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **No revisions were actually generated**
  - Symptom: The example checks `Document.Revisions` but finds none.
  - Preferred fix: Enable tracking, perform real content or formatting changes, then inspect revisions.
- **Invented revision APIs**
  - Symptom: Build failures caused by fake properties, methods, or namespaces for revisions.
  - Preferred fix: Use only documented `Document`, `Revision`, `RevisionCollection`, and accept/reject workflows.
- **Over-validating unrelated structure**
  - Symptom: The revision workflow succeeds, but the example fails because of unnecessary document-structure checks.
  - Preferred fix: Validate only the exact revision behavior requested by the task.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required package

```bash
dotnet add package Aspose.Words --version 26.3.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\track-changes\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve exact file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the verified winner from the latest batch report rather than a merely compiling draft.
- Bootstrap file-based inputs locally instead of depending on machine-specific paths.
