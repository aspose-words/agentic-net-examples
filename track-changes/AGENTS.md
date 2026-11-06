---
name: track-changes
description: Verified C# examples for track-changes scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Track Changes

## Purpose

Provide an AI coding assistant with concise guidance for implementing revision workflows that start tracking, generate edits, verify revision metadata, and accept or reject changes using only supported Aspose.Words APIs.

## Non-negotiable conventions

- Do not call members that are absent from Aspose.Words 26.10.0; only use documented types such as Document, DocumentBuilder, Revision, RevisionCollection, and RevisionGroup.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Track Changes Workflow workflow**: 40 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Confirm that the revision collection contains the expected number of items before proceeding with accept/reject logic.

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
  - Selected engine: llm
- `enable-revision-tracking-modify-a-table-cell-and-verify-the-revision-appears-in-the-collec.cs`
  - Task: Enable revision tracking, modify a table cell, and verify the revision appears in the collection.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `enable-tracking-apply-a-style-change-to-multiple-paragraphs-stop-tracking-and-verify-a-sin.cs`
  - Task: Enable tracking, apply a style change to multiple paragraphs, stop tracking, and verify a single revision group.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: llm
- `start-tracking-revisions-insert-an-image-stop-tracking-and-verify-the-image-insertion-appe.cs`
  - Task: Start tracking revisions, insert an image, stop tracking, and verify the image insertion appears as a revision.
  - Workflow: track-changes-workflow
  - Outputs: docx
  - Selected engine: mcp
- `load-a-document-start-tracking-insert-a-table-stop-tracking-then-accept-the-table-insertio.cs`
  - Task: Load a document, start tracking, insert a table, stop tracking, then accept the table insertion revision.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: llm
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
  - Selected engine: llm
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
  - Selected engine: existing_repo
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
  - Selected engine: llm
- `develop-a-utility-to-export-revision-metadata-to-a-csv-file-for-external-analysis.cs`
  - Task: Develop a utility to export revision metadata to a CSV file for external analysis.
  - Workflow: track-changes-workflow
  - Outputs: csv
  - Selected engine: mcp
- `create-a-custom-logger-that-captures-revision-metadata-during-document-editing-for-audit-p.cs`
  - Task: Create a custom logger that captures revision metadata during document editing for audit purposes.
  - Workflow: track-changes-workflow
  - Outputs: doc
  - Selected engine: llm
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

- **Unsupported API usage**
  - Symptom: Compilation fails because the code references types or members that do not exist in the selected Aspose.Words version.
  - Root cause: Invented or outdated APIs were used instead of the documented members for the current package.
  - Preferred fix: Replace the nonexistent members with supported APIs such as Document.StartTrackRevisions, Revision.RevisionType, Revision.Author, and Revision.DateTime.
  - Prevention: Consult the official Aspose.Words 26.10.0 reference before writing code and avoid guessing API names.

- **Incorrect method overload**
  - Symptom: CS1501 errors indicate that a called method has no matching overload signature.
  - Root cause: Methods like InsertParagraph were invoked with arguments that the API does not accept.
  - Preferred fix: Use the correct overloads, for example DocumentBuilder.Writeln(string) or Paragraph.ParagraphFormat.Style assignment.
  - Prevention: Verify method signatures via IntelliSense or the API documentation prior to coding.

- **Revision collection mutation during enumeration**
  - Symptom: Runtime InvalidOperationException reports that the RevisionCollection was modified while being enumerated.
  - Root cause: Accepting or rejecting revisions inside a foreach loop over doc.Revisions alters the collection during iteration.
  - Preferred fix: Copy the collection to an array (e.g., doc.Revisions.Cast<Revision>().ToArray()) before iterating, or iterate by index while adjusting for removals.
  - Prevention: Never modify RevisionCollection directly inside its enumerator.

- **Missing property in current API**
  - Symptom: CS1061 errors appear for members such as RevisionDateTime or RevisionGroup.Revisions.
  - Root cause: The referenced property does not exist in Aspose.Words 26.10.0; the correct property is Revision.DateTime.
  - Preferred fix: Replace RevisionDateTime with Revision.DateTime and avoid accessing non-existent RevisionGroup members.
  - Prevention: Reference the exact property names listed in the version-specific documentation.

- **Missing local input files**
  - Symptom: Runtime failures occur when loading a document that was never created on disk.
  - Root cause: Examples assume external files, folders, or images already exist without creating them first.
  - Preferred fix: Generate deterministic sample documents in memory or write them to a temporary file before loading or processing.
  - Prevention: Always bootstrap required inputs within the sample code.

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
Copy-Item ..\track-changes\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Review and update all examples when a new Aspose.Words version deprecates or renames revision-related members.
