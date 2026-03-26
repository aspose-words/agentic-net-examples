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

- `using System;` (37/37 files) ← category-specific
- `using Aspose.Words;` (36/37 files)
- `using System.IO;` (14/37 files)
- `using Aspose.Words.Replacing;` (5/37 files)
- `using Aspose.Words.Comparing;` (3/37 files)
- `using System.Text;` (3/37 files)
- `using System.Collections.Generic;` (3/37 files)
- `using Aspose.Words.Saving;` (2/37 files)
- `using Aspose.Words.Tables;` (2/37 files)
- `using Aspose.Words.Drawing;` (2/37 files)
- `using System.Linq;` (2/37 files)
- `using System.Drawing;` (1/37 files)

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
| [accept-all-revisions-loaded-docx-file-cleaned-document-...](./accept-all-revisions-loaded-docx-file-cleaned-document-new-file.cs) | `Document`, `DocumentBuilder`, `CleanedDocument` | Accept all revisions loaded docx file cleaned document new file |
| [accept-revisions-only-paragraphs-authored-specific-user...](./accept-revisions-only-paragraphs-authored-specific-user-while-rejecting-others.cs) | `Document`, `AuthorCriteria`, `Revisions` | Accept revisions only paragraphs authored specific user while rejecting others |
| [accept-revisions-specific-section-document-while-leavin...](./accept-revisions-specific-section-document-while-leaving-other-sections-unchanged.cs) | `Document`, `Input`, `Sections` | Accept revisions specific section document while leaving other sections uncha... |
| [after-accepting-all-revisions-lock-document-prevent-fur...](./after-accepting-all-revisions-lock-document-prevent-further-editing-without-tracking.cs) | `Document`, `ProtectionType` | After accepting all revisions lock document prevent further editing without t... |
| [after-rejecting-all-revisions-verify-that-document-s-or...](./after-rejecting-all-revisions-verify-that-document-s-original-content-matches.cs) | `DocumentBuilder`, `Document`, `Revisions` | After rejecting all revisions verify that document s original content matches |
| [batch-process-folder-documents-rejecting-revisions-auth...](./batch-process-folder-documents-rejecting-revisions-authored-given-user-across-all-files.cs) | `Document`, `AuthorRevisionCriteria`, `AppContext` | Batch process folder documents rejecting revisions authored given user across... |
| [batch-process-that-opens-multiple-documents-accepts-all...](./batch-process-that-opens-multiple-documents-accepts-all-revisions-saves-them-place.cs) | `Document`, `SearchOption` | Batch process that opens multiple documents accepts all revisions saves them... |
| [compare-three-versions-contract-sequentially-generating...](./compare-three-versions-contract-sequentially-generating-cumulative-revision-sets-each.cs) | `Document`, `DocumentBuilder`, `Contract_v1` | Compare three versions contract sequentially generating cumulative revision s... |
| [compare-two-word-documents-revision-differences-specify...](./compare-two-word-documents-revision-differences-specifying-author-name-comparison-date.cs) | `Document`, `Revisions`, `InvalidOperationException` | Compare two word documents revision differences specifying author name compar... |
| [custom-logger-that-captures-revision-metadata-during-do...](./custom-logger-that-captures-revision-metadata-during-document-editing-audit-purposes.cs) | `StringBuilder`, `Document`, `DocumentBuilder` | Custom logger that captures revision metadata during document editing audit p... |
| [determine-if-document-contains-any-revisions-before-fur...](./determine-if-document-contains-any-revisions-before-further-processing-checking.cs) | `Document` | Determine if document contains any revisions before further processing checking |
| [develop-utility-export-revision-metadata-csv-file-exter...](./develop-utility-export-revision-metadata-csv-file-external-analysis.cs) | `Document`, `StreamWriter`, `RevisionExporter` | Develop utility export revision metadata csv file external analysis |
| [doc-compare-revision-document-that-highlights-differenc...](./doc-compare-revision-document-that-highlights-differences-between-two-versions.cs) | `Document`, `Revisions`, `Original` | Doc compare revision document that highlights differences between two versions |
| [document-enable-revision-tracking-apply-formatting-chan...](./document-enable-revision-tracking-apply-formatting-changes-list-resulting-revision.cs) | `Document`, `DocumentBuilder`, `FirstSection` | Document enable revision tracking apply formatting changes list resulting rev... |
| [document-enable-tracking-perform-find-replace-operation...](./document-enable-tracking-perform-find-replace-operation-list-revisions.cs) | `Document`, `Input`, `Range` | Document enable tracking perform find replace operation list revisions |
| [document-start-tracking-insert-table-stop-tracking-then...](./document-start-tracking-insert-table-stop-tracking-then-accept-table-insertion-revision.cs) | `Document`, `DocumentBuilder`, `Input` | Document start tracking insert table stop tracking then accept table insertio... |
| [document-stream-start-tracking-add-header-stop-tracking...](./document-stream-start-tracking-add-header-stop-tracking-before.cs) | `Document`, `DocumentBuilder`, `HeaderFooterType` | Document stream start tracking add header stop tracking before |
| [enable-revision-tracking-modify-table-cell-verify-revis...](./enable-revision-tracking-modify-table-cell-verify-revision-appears-collection.cs) | `Document`, `DocumentBuilder`, `FirstParagraph` | Enable revision tracking modify table cell verify revision appears collection |
| [enable-tracking-apply-style-change-multiple-paragraphs-...](./enable-tracking-apply-style-change-multiple-paragraphs-stop-tracking-verify-single.cs) | `Revisions`, `Document`, `DocumentBuilder` | Enable tracking apply style change multiple paragraphs stop tracking verify s... |
| [function-that-returns-true-if-any-revision-author-match...](./function-that-returns-true-if-any-revision-author-matches-specified-list-names.cs) | `ArgumentNullException`, `Document`, `Collections` | Function that returns true if any revision author matches specified list names |
| [hasrevisions-property-conditionally-apply-watermark-ind...](./hasrevisions-property-conditionally-apply-watermark-indicating-pending-changes.cs) | `Document`, `DocumentBuilder`, `Input` | Hasrevisions property conditionally apply watermark indicating pending changes |
| [iterate-over-revisions-group-them-type-output-summary-i...](./iterate-over-revisions-group-them-type-output-summary-insertion-deletion-counts.cs) | `RevisionType`, `Document`, `DocumentBuilder` | Iterate over revisions group them type output summary insertion deletion counts |
| [iterate-through-all-revisions-document-log-each-revisio...](./iterate-through-all-revisions-document-log-each-revision-s-author-timestamp.cs) | `Document`, `Input`, `Output` | Iterate through all revisions document log each revision s author timestamp |
| [new-document-start-tracking-revisions-insert-paragraph-...](./new-document-start-tracking-revisions-insert-paragraph-then-stop-tracking.cs) | `Document`, `DocumentBuilder`, `TrackedRevisions` | New document start tracking revisions insert paragraph then stop tracking |
| [programmatically-compare-document-against-its-previous-...](./programmatically-compare-document-against-its-previous-version-revision-report-memory.cs) | `Document`, `StringBuilder`, `Revisions` | Programmatically compare document against its previous version revision repor... |
| [programmatically-reject-all-deletions-while-keeping-ins...](./programmatically-reject-all-deletions-while-keeping-insertions-formatting-changes.cs) | `Document`, `DeletionCriteria`, `RevisionType` | Programmatically reject all deletions while keeping insertions formatting cha... |
| [reject-all-formatting-revisions-but-keep-content-insert...](./reject-all-formatting-revisions-but-keep-content-insertions-deletions-still-intact.cs) | `Document`, `Collections`, `Input` | Reject all formatting revisions but keep content insertions deletions still i... |
| [reject-all-tracked-changes-document-then-verify-hasrevi...](./reject-all-tracked-changes-document-then-verify-hasrevisions-property-returns-false.cs) | `Document`, `DocumentBuilder`, `RevisionCollection` | Reject all tracked changes document then verify hasrevisions property returns... |
| [reject-revisions-older-than-specific-date-while-preserv...](./reject-revisions-older-than-specific-date-while-preserving-newer-changes-document.cs) | `Document`, `DocumentBuilder`, `DateRevisionCriteria` | Reject revisions older than specific date while preserving newer changes docu... |
| [reject-specific-revision-based-author-name-then-confirm...](./reject-specific-revision-based-author-name-then-confirm-change-is-no-longer-present.cs) | `Revisions`, `Document`, `DocumentBuilder` | Reject specific revision based author name then confirm change is no longer p... |
| ... | | *and 7 more files* |

## Category Statistics
- Total examples: 37

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for track-changes patterns.


## Command Reference

### Build and Run

Files in this folder are standalone `.cs` examples. Run one example at a time by copying it into a temporary console project as `Program.cs`.

```bash
# Create a temporary console project from the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\track-changes\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

### Notes for Agents and Developers

- Treat every `.cs` file in `track-changes/` as a full console program, not a snippet.
- Run one file at a time by copying it to `Program.cs`.
- If a sample needs input documents, images, fonts, or data files, place them in the temporary project directory before running.
- See the root [AGENTS.md](../AGENTS.md) for repository-wide prerequisites, project file template, and testing guidance.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
