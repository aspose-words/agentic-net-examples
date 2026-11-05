---
name: join-documents
description: Verified C# examples for join-documents scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Join Documents

## Purpose

Provide an AI coding assistant with clear, reproducible patterns for joining documents that avoid common API misuse, ensure all source artifacts are generated locally, and validate that the merged output contains the expected content.

## Non-negotiable conventions

- Never reference members that are not part of the Aspose.Words public API, such as BookmarkCollection.Exists or non-existent InsertDocument overloads.
- Do not attempt to append or insert a node into a parent that cannot accept that node type.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Insert Document workflow**: 18 examples
- **Append Document workflow**: 12 examples
- **General Join Workflow workflow**: 1 examples
- **Mail Merge Insert workflow**: 2 examples
- **Section And Format Preservation workflow**: 3 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Validate the merged document by confirming file existence and by searching for unique text fragments from every source document.

## File-to-task reference

- `load-a-docx-template-insert-a-source-docx-at-a-bookmark-and-save-the-result-as-pdf.cs`
  - Task: Load a DOCX template, insert a source DOCX at a bookmark, and save the result as PDF.
  - Workflow: insert-document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `load-a-docx-insert-a-pdf-converted-docx-at-a-bookmark-and-save-the-merged-document-as-odt.cs`
  - Task: Load a DOCX, insert a PDF-converted DOCX at a bookmark, and save the merged document as ODT.
  - Workflow: insert-document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `load-a-password-protected-docx-insert-a-source-docx-at-a-bookmark-then-remove-protection-a.cs`
  - Task: Load a password-protected DOCX, insert a source DOCX at a bookmark, then remove protection and save as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: mcp
- `protect-a-destination-docx-with-password-before-inserting-a-source-docx-then-remove-protec.cs`
  - Task: Protect a destination DOCX with password before inserting a source DOCX, then remove protection and save as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: mcp
- `protect-a-destination-docx-with-read-only-restriction-before-appending-source-docx-then-re.cs`
  - Task: Protect a destination DOCX with read-only restriction before appending source DOCX, then remove restriction and save as PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `use-documentbuilder-movetobookmark-to-navigate-to-a-bookmark-named-content-then-insert-a-s.cs`
  - Task: Use DocumentBuilder.MoveToBookmark to navigate to a bookmark named "Content", then insert a source DOCX with KeepSourceFormatting.
  - Workflow: insert-document
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-docx-at-a-bookmark-inside-a-header-then-save-the-document-as-docx-preserving-head.cs`
  - Task: Insert a DOCX at a bookmark inside a header, then save the document as DOCX preserving header formatting.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: mcp
- `insert-a-docx-at-a-bookmark-inside-a-table-cell-then-export-document-to-odt-preserving-tab.cs`
  - Task: Insert a DOCX at a bookmark inside a table cell, then export document to ODT preserving table structure.
  - Workflow: insert-document
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-docx-at-multiple-bookmarks-named-header-and-footer-then-export-the-final-document.cs`
  - Task: Insert a DOCX at multiple bookmarks named Header and Footer, then export the final document to HTML format.
  - Workflow: insert-document
  - Outputs: docx, html
  - Selected engine: mcp
- `use-documentbuilder-insertdocument-to-insert-a-docx-at-the-current-cursor-position-with-ke.cs`
  - Task: Use DocumentBuilder.InsertDocument to insert a DOCX at the current cursor position with KeepSourceFormatting, saving as HTML.
  - Workflow: insert-document
  - Outputs: docx, html
  - Selected engine: mcp
- `use-documentbuilder-insertdocument-with-importformatmode-usedestinationstyles-to-insert-a.cs`
  - Task: Use DocumentBuilder.InsertDocument with ImportFormatMode.UseDestinationStyles to insert a DOCX into a styled template, then save as HTML.
  - Workflow: insert-document
  - Outputs: docx, html
  - Selected engine: mcp
- `insert-a-docx-during-a-find-replace-operation-that-matches-the-word-insertme-then-save-as.cs`
  - Task: Insert a DOCX during a find-replace operation that matches the word "INSERTME", then save as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: existing_repo
- `use-findreplaceoptions-to-replace-the-phrase-placeholder-with-a-docx-document-then-save-as.cs`
  - Task: Use FindReplaceOptions to replace the phrase "PLACEHOLDER" with a DOCX document, then save as ODT.
  - Workflow: general-join-workflow
  - Outputs: docx
  - Selected engine: existing_repo
- `use-findreplaceoptions-with-a-replacingcallback-to-insert-a-docx-whenever-phrase-insert-he.cs`
  - Task: Use FindReplaceOptions with a ReplacingCallback to insert a DOCX whenever phrase "INSERT_HERE" is found, then save as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: mcp
- `insert-a-docx-during-a-find-replace-operation-that-replaces-an-xml-tag-then-save-the-docum.cs`
  - Task: Insert a DOCX during a find-replace operation that replaces an XML tag, then save the document as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc, xml
  - Selected engine: existing_repo
- `set-mailmerge-insertdocument-event-to-inject-a-docx-for-each-data-row-and-generate-a-combi.cs`
  - Task: Set MailMerge.InsertDocument event to inject a DOCX for each data row, and generate a combined PDF file.
  - Workflow: insert-document
  - Outputs: docx, pdf
  - Selected engine: existing_repo
- `insert-a-docx-during-mail-merge-for-record-then-apply-importlanguageinfo-to-preserve-langu.cs`
  - Task: Insert a DOCX during mail merge for record, then apply ImportLanguageInfo to preserve language settings, save as PDF.
  - Workflow: mail-merge-insert
  - Outputs: docx, pdf
  - Selected engine: mcp
- `use-importlanguageinfo-option-to-preserve-language-settings-when-appending-a-source-docx-t.cs`
  - Task: Use ImportLanguageInfo option to preserve language settings when appending a source DOCX to a destination DOCX.
  - Workflow: append-document
  - Outputs: docx
  - Selected engine: existing_repo
- `manually-import-specific-tables-from-a-source-docx-using-nodeimporter-and-insert-them-into.cs`
  - Task: Manually import specific tables from a source DOCX using NodeImporter and insert them into a destination DOCX.
  - Workflow: insert-document
  - Outputs: docx
  - Selected engine: mcp
- `use-nodeimporter-to-import-only-paragraph-nodes-from-a-source-docx-and-insert-them-after-a.cs`
  - Task: Use NodeImporter to import only paragraph nodes from a source DOCX and insert them after a specific bookmark.
  - Workflow: insert-document
  - Outputs: docx
  - Selected engine: mcp
- `append-an-rtf-document-to-an-existing-docx-using-usedestinationstyles-then-export-the-comb.cs`
  - Task: Append an RTF document to an existing DOCX using UseDestinationStyles, then export the combined file to DOCX.
  - Workflow: append-document
  - Outputs: docx, doc
  - Selected engine: mcp
- `append-a-docx-containing-custom-styles-to-a-destination-docx-using-importformatmode-usedes.cs`
  - Task: Append a DOCX containing custom styles to a destination DOCX, using ImportFormatMode.UseDestinationStyles, then export to PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `append-a-docx-generated-from-a-mail-merge-operation-to-an-existing-pdf-converted-docx-pres.cs`
  - Task: Append a DOCX generated from a mail merge operation to an existing PDF-converted DOCX, preserving destination styles.
  - Workflow: mail-merge-insert
  - Outputs: docx, pdf
  - Selected engine: mcp
- `append-a-docx-with-video-files-to-a-destination-docx-ensuring-videos-remain-functional-aft.cs`
  - Task: Append a DOCX with video files to a destination DOCX, ensuring videos remain functional after saving as PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `append-a-docx-containing-footnotes-to-a-destination-docx-and-ensure-footnote-numbering-con.cs`
  - Task: Append a DOCX containing footnotes to a destination DOCX and ensure footnote numbering continues, then save as PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: existing_repo
- `append-a-docx-containing-a-bibliography-to-a-research-paper-docx-then-update-fields-and-ex.cs`
  - Task: Append a DOCX containing a bibliography to a research paper DOCX, then update fields and export to PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `combine-three-odt-files-into-one-docx-applying-importformatmode-keepsourceformatting-then.cs`
  - Task: Combine three ODT files into one DOCX, applying ImportFormatMode.KeepSourceFormatting, then save as DOCX.
  - Workflow: section-and-format-preservation
  - Outputs: docx, doc
  - Selected engine: mcp
- `combine-a-docx-and-an-odt-file-into-a-single-pdf-using-importformatmode-keepsourceformatti.cs`
  - Task: Combine a DOCX and an ODT file into a single PDF, using ImportFormatMode.KeepSourceFormatting for the ODT content.
  - Workflow: section-and-format-preservation
  - Outputs: docx, pdf
  - Selected engine: mcp
- `combine-multiple-docx-files-into-a-single-pdf-ensuring-each-source-document-retains-its-or.cs`
  - Task: Combine multiple DOCX files into a single PDF, ensuring each source document retains its original formatting.
  - Workflow: section-and-format-preservation
  - Outputs: docx, pdf
  - Selected engine: mcp
- `append-multiple-docx-files-in-a-loop-each-with-different-importformatmode-settings-and-out.cs`
  - Task: Append multiple DOCX files in a loop, each with different ImportFormatMode settings, and output combined file as PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `batch-process-a-directory-of-docx-files-appending-each-to-a-master-document-and-saving-res.cs`
  - Task: Batch process a directory of DOCX files, appending each to a master document and saving result as PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `batch-append-docx-files-from-multiple-subfolders-into-a-single-master-document-applying-im.cs`
  - Task: Batch append DOCX files from multiple subfolders into a single master document, applying ImportFormatMode.UseDestinationStyles, then export to PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `use-documentbuilder-to-insert-a-docx-at-end-of-sections-in-a-destination-docx-then-export.cs`
  - Task: Use DocumentBuilder to insert a DOCX at end of sections in a destination DOCX, then export to DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: mcp
- `validate-that-all-inserted-sections-contain-correct-page-numbers-by-iterating-through-dest.cs`
  - Task: Validate that all inserted sections contain correct page numbers by iterating through destination document nodes after insertion.
  - Workflow: insert-document
  - Outputs: docx
  - Selected engine: mcp
- `append-a-docx-generated-from-a-web-service-to-an-existing-docx-then-encrypt-final-pdf-with.cs`
  - Task: Append a DOCX generated from a web service to an existing DOCX, then encrypt final PDF with password.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `append-a-docx-generated-from-a-rest-api-to-an-existing-docx-then-encrypt-final-pdf-with-pa.cs`
  - Task: Append a DOCX generated from a REST API to an existing DOCX, then encrypt final PDF with password.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: llm

## Common failure patterns and preferred agent fixes

- **Nonexistent API member**
  - Symptom: Compilation fails with CS1061, CS0246, or CS0234 because a referenced type or member cannot be found.
  - Root cause: The code references classes, namespaces, or members that are not part of the Aspose.Words 26.10.0 API (e.g., BookmarkCollection.Exists, Aspose.Words.Protection, Import namespace, FindReplaceOptions.ReplacementDocument).
  - Preferred fix: Replace the invalid reference with the correct API member (e.g., use mainDoc.Range.Bookmarks["name"] to test existence, use Document.Protect for protection, and use InsertDocument on DocumentBuilder).
  - Prevention: Consult the official Aspose.Words documentation for the exact class and member names before coding and rely on IntelliSense to verify availability.

- **Invalid node insertion location**
  - Symptom: Runtime throws System.ArgumentException stating that a node cannot be inserted at the current location.
  - Root cause: InsertDocument is called on a node that cannot accept the imported content, such as inserting a whole section into a Run or inserting into a bookmark that has been removed.
  - Preferred fix: Move the DocumentBuilder to a valid container (e.g., a Paragraph or the bookmark itself) before calling InsertDocument, and ensure the bookmark remains in the document.
  - Prevention: Always verify the target node type with Node.NodeType and use DocumentBuilder.MoveToBookmark or MoveTo(Run) to position the builder correctly.

- **Improper FindReplaceOptions usage**
  - Symptom: Compilation or runtime errors occur when attempting to replace a placeholder with a document using FindReplaceOptions.
  - Root cause: The code tries to assign a Document to FindReplaceOptions.ReplacementDocument or passes a Document to Range.Replace, which expects a string.
  - Preferred fix: Implement a custom IReplacingCallback that inserts the document at the match location, and pass that callback via FindReplaceOptions.ReplacingCallback.
  - Prevention: Use the documented pattern for document insertion during find-replace: create a callback class that receives the match node and calls DocumentBuilder.InsertDocument.

- **Incorrect protection API reference**
  - Symptom: Compilation fails because Protection or InsertDocument members are missing on the MailMerge or Document classes.
  - Root cause: The sample uses outdated or non-existent protection APIs (e.g., Aspose.Words.Protection namespace) and assumes a MailMerge.InsertDocument event that does not exist.
  - Preferred fix: Apply protection with Document.Protect(ProtectionType, password) and remove it with Document.Unprotect(password); handle mail-merge insertion via MailMerge.FieldMergingCallback and DocumentBuilder.InsertDocument.
  - Prevention: Reference the current protection methods in the Document class and follow the official mail-merge insertion workflow.

- **Missing source document creation**
  - Symptom: The example assumes that source DOCX files already exist, leading to FileNotFoundException or empty merges.
  - Root cause: No code is provided to generate the prerequisite documents before the join operation.
  - Preferred fix: Add code that creates each source document with DocumentBuilder, saves it to a temporary path, and then loads it for insertion or appending.
  - Prevention: Always bootstrap required source files at the start of the sample, using clearly named temporary directories.

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
Copy-Item ..\join-documents\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Keep the sample code in sync with the Aspose.Words version used for verification and update any API changes in future releases.
