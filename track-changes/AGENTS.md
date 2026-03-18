---
name: track-changes
description: C# examples for track-changes using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - track-changes

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **track-changes** category.
This folder contains standalone C# examples for track-changes operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **track-changes**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using System;` (40/40 files) ← category-specific
- `using Aspose.Words;` (40/40 files)
- `using System.IO;` (8/40 files)
- `using Aspose.Words.Replacing;` (4/40 files)
- `using Aspose.Words.Comparing;` (3/40 files)
- `using System.Text;` (2/40 files)
- `using Aspose.Words.Tables;` (2/40 files)
- `using System.Collections.Generic;` (2/40 files)
- `using Aspose.Words.Drawing;` (2/40 files)
- `using System.Linq;` (2/40 files)
- `using Aspose.Words.Saving;` (1/40 files)
- `using System.Drawing;` (1/40 files)

## Common Code Pattern

Most files follow this pattern:

```csharp
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
// ... operations ...
doc.Save("output.docx");
```

## Files in this folder

| File | Key APIs | Description |
|------|----------|-------------|
| [accept-all-revisions-document-then-re-enable-tracking-c...](./accept-all-revisions-document-then-re-enable-tracking-capture-subsequent-changes.cs) | `Document`, `DocumentBuilder`, `Aspose` | Accept all revisions document then re enable tracking capture subsequent changes |
| [accept-all-revisions-loaded-docx-file-cleaned-document-...](./accept-all-revisions-loaded-docx-file-cleaned-document-new-file.cs) | `Document`, `Aspose`, `OriginalWithRevisions` | Accept all revisions loaded docx file cleaned document new file |
| [accept-revisions-only-paragraphs-authored-specific-user...](./accept-revisions-only-paragraphs-authored-specific-user-while-rejecting-others.cs) | `Document`, `AuthorCriteria`, `Revisions` | Accept revisions only paragraphs authored specific user while rejecting others |
| [accept-revisions-specific-section-document-while-leavin...](./accept-revisions-specific-section-document-while-leaving-other-sections-unchanged.cs) | `Document`, `Aspose`, `Words` | Accept revisions specific section document while leaving other sections uncha... |
| [after-accepting-all-revisions-lock-document-prevent-fur...](./after-accepting-all-revisions-lock-document-prevent-further-editing-without-tracking.cs) | `Document`, `Aspose`, `Words` | After accepting all revisions lock document prevent further editing without t... |
| [after-rejecting-all-revisions-verify-that-document-s-or...](./after-rejecting-all-revisions-verify-that-document-s-original-content-matches.cs) | `Document`, `Aspose`, `OriginalWithRevisions` | After rejecting all revisions verify that document s original content matches |
| [batch-process-folder-documents-rejecting-revisions-auth...](./batch-process-folder-documents-rejecting-revisions-authored-given-user-across-all-files.cs) | `Document`, `AuthorRevisionCriteria`, `System` | Batch process folder documents rejecting revisions authored given user across... |
| [batch-process-that-opens-multiple-documents-accepts-all...](./batch-process-that-opens-multiple-documents-accepts-all-revisions-saves-them-place.cs) | `Document`, `System`, `Aspose` | Batch process that opens multiple documents accepts all revisions saves them... |
| [compare-three-versions-contract-sequentially-generating...](./compare-three-versions-contract-sequentially-generating-cumulative-revision-sets-each.cs) | `Document`, `Revisions`, `Aspose` | Compare three versions contract sequentially generating cumulative revision s... |
| [compare-two-word-documents-revision-differences-specify...](./compare-two-word-documents-revision-differences-specifying-author-name-comparison-date.cs) | `Document`, `Aspose`, `Revisions` | Compare two word documents revision differences specifying author name compar... |
| [custom-logger-that-captures-revision-metadata-during-do...](./custom-logger-that-captures-revision-metadata-during-document-editing-audit-purposes.cs) | `StringBuilder`, `Document`, `DocumentBuilder` | Custom logger that captures revision metadata during document editing audit p... |
| [determine-if-document-contains-any-revisions-before-fur...](./determine-if-document-contains-any-revisions-before-further-processing-checking.cs) | `Document`, `Aspose` | Determine if document contains any revisions before further processing checking |
| [develop-utility-export-revision-metadata-csv-file-exter...](./develop-utility-export-revision-metadata-csv-file-external-analysis.cs) | `Document`, `StreamWriter`, `RevisionExporter` | Develop utility export revision metadata csv file external analysis |
| [doc-compare-revision-document-that-highlights-differenc...](./doc-compare-revision-document-that-highlights-differences-between-two-versions.cs) | `Document`, `Aspose`, `Revisions` | Doc compare revision document that highlights differences between two versions |
| [document-enable-revision-tracking-apply-formatting-chan...](./document-enable-revision-tracking-apply-formatting-changes-list-resulting-revision.cs) | `Document`, `DocumentBuilder`, `Aspose` | Document enable revision tracking apply formatting changes list resulting rev... |
| [document-enable-tracking-perform-find-replace-operation...](./document-enable-tracking-perform-find-replace-operation-list-revisions.cs) | `Document`, `Aspose`, `Input` | Document enable tracking perform find replace operation list revisions |
| [document-start-tracking-insert-table-stop-tracking-then...](./document-start-tracking-insert-table-stop-tracking-then-accept-table-insertion-revision.cs) | `Document`, `DocumentBuilder`, `Aspose` | Document start tracking insert table stop tracking then accept table insertio... |
| [document-stream-start-tracking-add-header-stop-tracking...](./document-stream-start-tracking-add-header-stop-tracking-before.cs) | `Document`, `DocumentBuilder`, `System` | Document stream start tracking add header stop tracking before |
| [enable-revision-tracking-modify-table-cell-verify-revis...](./enable-revision-tracking-modify-table-cell-verify-revision-appears-collection.cs) | `Document`, `DocumentBuilder`, `Aspose` | Enable revision tracking modify table cell verify revision appears collection |
| [enable-tracking-apply-style-change-multiple-paragraphs-...](./enable-tracking-apply-style-change-multiple-paragraphs-stop-tracking-verify-single.cs) | `Revisions`, `Document`, `DocumentBuilder` | Enable tracking apply style change multiple paragraphs stop tracking verify s... |
| [function-that-returns-true-if-any-revision-author-match...](./function-that-returns-true-if-any-revision-author-matches-specified-list-names.cs) | `ArgumentNullException`, `Document`, `System` | Function that returns true if any revision author matches specified list names |
| [group-sequential-insertions-revisiongroup-accept-entire...](./group-sequential-insertions-revisiongroup-accept-entire-group-single-call.cs) | `Document`, `DocumentBuilder`, `Aspose` | Group sequential insertions revisiongroup accept entire group single call |
| [hasrevisions-property-conditionally-apply-watermark-ind...](./hasrevisions-property-conditionally-apply-watermark-indicating-pending-changes.cs) | `Document`, `DocumentBuilder`, `Shape` | Hasrevisions property conditionally apply watermark indicating pending changes |
| [implement-error-handling-attempts-accept-revision-that-...](./implement-error-handling-attempts-accept-revision-that-has-already-been-rejected.cs) | `Document`, `DocumentBuilder`, `Aspose` | Implement error handling attempts accept revision that has already been rejected |
| [iterate-over-revisions-group-them-type-output-summary-i...](./iterate-over-revisions-group-them-type-output-summary-insertion-deletion-counts.cs) | `RevisionType`, `Document`, `System` | Iterate over revisions group them type output summary insertion deletion counts |
| [iterate-through-all-revisions-document-log-each-revisio...](./iterate-through-all-revisions-document-log-each-revision-s-author-timestamp.cs) | `Document`, `Aspose`, `Input` | Iterate through all revisions document log each revision s author timestamp |
| [new-document-start-tracking-revisions-insert-paragraph-...](./new-document-start-tracking-revisions-insert-paragraph-then-stop-tracking.cs) | `Document`, `DocumentBuilder`, `Aspose` | New document start tracking revisions insert paragraph then stop tracking |
| [programmatically-compare-document-against-its-previous-...](./programmatically-compare-document-against-its-previous-version-revision-report-memory.cs) | `Document`, `StringBuilder`, `System` | Programmatically compare document against its previous version revision repor... |
| [programmatically-reject-all-deletions-while-keeping-ins...](./programmatically-reject-all-deletions-while-keeping-insertions-formatting-changes.cs) | `Document`, `DeletionCriteria`, `Aspose` | Programmatically reject all deletions while keeping insertions formatting cha... |
| [reject-all-formatting-revisions-but-keep-content-insert...](./reject-all-formatting-revisions-but-keep-content-insertions-deletions-still-intact.cs) | `Document`, `System`, `Collections` | Reject all formatting revisions but keep content insertions deletions still i... |
| ... | | *and 10 more files* |

## Category Statistics
- Total examples: 40

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for track-changes patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082627`
<!-- AUTOGENERATED:END -->
