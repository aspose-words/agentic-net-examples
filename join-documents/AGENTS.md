---
name: join-documents
description: Verified C# examples for document joining scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Join Documents

## Purpose

This folder is a live, curated example set for document joining scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free document append, insert, section-preserving join, stream-based join, and mail-merge-driven document insertion using direct Aspose.Words APIs.

## Non-negotiable conventions

- Create all required local source documents before joining them.
- Use AppendDocument for whole-document append scenarios.
- Use InsertDocument or InsertDocumentInline only when the task explicitly requires insertion at a specific position or during a mail merge workflow.
- Specify ImportFormatMode explicitly when appending documents.
- Do not invent unsupported document-join APIs.
- Reset MemoryStream.Position before loading from a written stream.
- Validate that the merged output contains content from all expected sources.

## Recommended workflow selection

- **Append-document workflow**: 12 examples
- **Insert-document workflow**: 18 examples
- **Mail-merge insertion workflow**: 2 examples
- **Section / formatting preservation workflow**: 3 examples
- **General join workflow**: 1 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. Source documents, streams, and folders must be bootstrapped locally whenever the task implies an existing input.
3. Join operations must use documented Aspose.Words APIs and valid insertion locations.
4. Requested merged output files must actually be written.
5. The merged result should prove that content from all required source documents was included.

## File-to-task reference

- `load-a-docx-template-insert-a-source-docx-at-a-bookmark-and-save-the-result-as-pdf.cs`
  - Task: Load a DOCX template, insert a source DOCX at a bookmark, and save the result as PDF.
  - Workflow: insert-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `load-a-docx-insert-a-pdf-converted-docx-at-a-bookmark-and-save-the-merged-document-as-odt.cs`
  - Task: Load a DOCX, insert a PDF‑converted DOCX at a bookmark, and save the merged document as ODT.
  - Workflow: insert-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `load-a-password-protected-docx-insert-a-source-docx-at-a-bookmark-then-remove-protection-a.cs`
  - Task: Load a password‑protected DOCX, insert a source DOCX at a bookmark, then remove protection and save as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: verified
- `protect-a-destination-docx-with-password-before-inserting-a-source-docx-then-remove-protec.cs`
  - Task: Protect a destination DOCX with password before inserting a source DOCX, then remove protection and save as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: verified
- `protect-a-destination-docx-with-read-only-restriction-before-appending-source-docx-then-re.cs`
  - Task: Protect a destination DOCX with read‑only restriction before appending source DOCX, then remove restriction and save as PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `use-documentbuilder-movetobookmark-to-navigate-to-a-bookmark-named-content-then-insert-a-s.cs`
  - Task: Use DocumentBuilder.MoveToBookmark to navigate to a bookmark named “Content”, then insert a source DOCX with KeepSourceFormatting.
  - Workflow: insert-document
  - Outputs: docx
  - Selected engine: verified
- `insert-a-docx-at-a-bookmark-inside-a-header-then-save-the-document-as-docx-preserving-head.cs`
  - Task: Insert a DOCX at a bookmark inside a header, then save the document as DOCX preserving header formatting.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: verified
- `insert-a-docx-at-a-bookmark-inside-a-table-cell-then-export-document-to-odt-preserving-tab.cs`
  - Task: Insert a DOCX at a bookmark inside a table cell, then export document to ODT preserving table structure.
  - Workflow: insert-document
  - Outputs: docx
  - Selected engine: verified
- `insert-a-docx-at-multiple-bookmarks-named-header-and-footer-then-export-the-final-document.cs`
  - Task: Insert a DOCX at multiple bookmarks named Header and Footer, then export the final document to HTML format.
  - Workflow: insert-document
  - Outputs: docx, html
  - Selected engine: verified
- `use-documentbuilder-insertdocument-to-insert-a-docx-at-the-current-cursor-position-with-ke.cs`
  - Task: Use DocumentBuilder.InsertDocument to insert a DOCX at the current cursor position with KeepSourceFormatting, saving as HTML.
  - Workflow: insert-document
  - Outputs: docx, html
  - Selected engine: verified
- `use-documentbuilder-insertdocument-with-importformatmode-usedestinationstyles-to-insert-a.cs`
  - Task: Use DocumentBuilder.InsertDocument with ImportFormatMode.UseDestinationStyles to insert a DOCX into a styled template, then save as HTML.
  - Workflow: insert-document
  - Outputs: docx, html
  - Selected engine: verified
- `insert-a-docx-during-a-find-replace-operation-that-matches-the-word-insertme-then-save-as.cs`
  - Task: Insert a DOCX during a find‑replace operation that matches the word “INSERTME”, then save as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: verified
- `use-findreplaceoptions-to-replace-the-phrase-placeholder-with-a-docx-document-then-save-as.cs`
  - Task: Use FindReplaceOptions to replace the phrase “PLACEHOLDER” with a DOCX document, then save as ODT.
  - Workflow: general-join-workflow
  - Outputs: docx
  - Selected engine: verified
- `use-findreplaceoptions-with-a-replacingcallback-to-insert-a-docx-whenever-phrase-insert-he.cs`
  - Task: Use FindReplaceOptions with a ReplacingCallback to insert a DOCX whenever phrase “INSERT_HERE” is found, then save as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: verified
- `insert-a-docx-during-a-find-replace-operation-that-replaces-an-xml-tag-then-save-the-docum.cs`
  - Task: Insert a DOCX during a find‑replace operation that replaces an XML tag, then save the document as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc, xml
  - Selected engine: verified
- `set-mailmerge-insertdocument-event-to-inject-a-docx-for-each-data-row-and-generate-a-combi.cs`
  - Task: Set MailMerge.InsertDocument event to inject a DOCX for each data row, and generate a combined PDF file.
  - Workflow: insert-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `insert-a-docx-during-mail-merge-for-record-then-apply-importlanguageinfo-to-preserve-langu.cs`
  - Task: Insert a DOCX during mail merge for record, then apply ImportLanguageInfo to preserve language settings, save as PDF.
  - Workflow: mail-merge-insert
  - Outputs: docx, pdf
  - Selected engine: verified
- `use-importlanguageinfo-option-to-preserve-language-settings-when-appending-a-source-docx-t.cs`
  - Task: Use ImportLanguageInfo option to preserve language settings when appending a source DOCX to a destination DOCX.
  - Workflow: append-document
  - Outputs: docx
  - Selected engine: verified
- `manually-import-specific-tables-from-a-source-docx-using-nodeimporter-and-insert-them-into.cs`
  - Task: Manually import specific tables from a source DOCX using NodeImporter and insert them into a destination DOCX.
  - Workflow: insert-document
  - Outputs: docx
  - Selected engine: verified
- `use-nodeimporter-to-import-only-paragraph-nodes-from-a-source-docx-and-insert-them-after-a.cs`
  - Task: Use NodeImporter to import only paragraph nodes from a source DOCX and insert them after a specific bookmark.
  - Workflow: insert-document
  - Outputs: docx
  - Selected engine: verified
- `append-an-rtf-document-to-an-existing-docx-using-usedestinationstyles-then-export-the-comb.cs`
  - Task: Append an RTF document to an existing DOCX using UseDestinationStyles, then export the combined file to DOCX.
  - Workflow: append-document
  - Outputs: docx, doc
  - Selected engine: verified
- `append-a-docx-containing-custom-styles-to-a-destination-docx-using-importformatmode-usedes.cs`
  - Task: Append a DOCX containing custom styles to a destination DOCX, using ImportFormatMode.UseDestinationStyles, then export to PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `append-a-docx-generated-from-a-mail-merge-operation-to-an-existing-pdf-converted-docx-pres.cs`
  - Task: Append a DOCX generated from a mail merge operation to an existing PDF‑converted DOCX, preserving destination styles.
  - Workflow: mail-merge-insert
  - Outputs: docx, pdf
  - Selected engine: verified
- `append-a-docx-with-video-files-to-a-destination-docx-ensuring-videos-remain-functional-aft.cs`
  - Task: Append a DOCX with video files to a destination DOCX, ensuring videos remain functional after saving as PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `append-a-docx-containing-footnotes-to-a-destination-docx-and-ensure-footnote-numbering-con.cs`
  - Task: Append a DOCX containing footnotes to a destination DOCX and ensure footnote numbering continues, then save as PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `append-a-docx-containing-a-bibliography-to-a-research-paper-docx-then-update-fields-and-ex.cs`
  - Task: Append a DOCX containing a bibliography to a research paper DOCX, then update fields and export to PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `combine-three-odt-files-into-one-docx-applying-importformatmode-keepsourceformatting-then.cs`
  - Task: Combine three ODT files into one DOCX, applying ImportFormatMode.KeepSourceFormatting, then save as DOCX.
  - Workflow: section-and-format-preservation
  - Outputs: docx, doc
  - Selected engine: verified
- `combine-a-docx-and-an-odt-file-into-a-single-pdf-using-importformatmode-keepsourceformatti.cs`
  - Task: Combine a DOCX and an ODT file into a single PDF, using ImportFormatMode.KeepSourceFormatting for the ODT content.
  - Workflow: section-and-format-preservation
  - Outputs: docx, pdf
  - Selected engine: verified
- `combine-multiple-docx-files-into-a-single-pdf-ensuring-each-source-document-retains-its-or.cs`
  - Task: Combine multiple DOCX files into a single PDF, ensuring each source document retains its original formatting.
  - Workflow: section-and-format-preservation
  - Outputs: docx, pdf
  - Selected engine: verified
- `append-multiple-docx-files-in-a-loop-each-with-different-importformatmode-settings-and-out.cs`
  - Task: Append multiple DOCX files in a loop, each with different ImportFormatMode settings, and output combined file as PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `batch-process-a-directory-of-docx-files-appending-each-to-a-master-document-and-saving-res.cs`
  - Task: Batch process a directory of DOCX files, appending each to a master document and saving result as PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `batch-append-docx-files-from-multiple-subfolders-into-a-single-master-document-applying-im.cs`
  - Task: Batch append DOCX files from multiple subfolders into a single master document, applying ImportFormatMode.UseDestinationStyles, then export to PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `use-documentbuilder-to-insert-a-docx-at-end-of-sections-in-a-destination-docx-then-export.cs`
  - Task: Use DocumentBuilder to insert a DOCX at end of sections in a destination DOCX, then export to DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: verified
- `validate-that-all-inserted-sections-contain-correct-page-numbers-by-iterating-through-dest.cs`
  - Task: Validate that all inserted sections contain correct page numbers by iterating through destination document nodes after insertion.
  - Workflow: insert-document
  - Outputs: docx
  - Selected engine: verified
- `append-a-docx-generated-from-a-web-service-to-an-existing-docx-then-encrypt-final-pdf-with.cs`
  - Task: Append a DOCX generated from a web service to an existing DOCX, then encrypt final PDF with password.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: verified
- `append-a-docx-generated-from-a-rest-api-to-an-existing-docx-then-encrypt-final-pdf-with-pa.cs`
  - Task: Append a DOCX generated from a REST API to an existing DOCX, then encrypt final PDF with password.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: verified

## Common failure patterns and preferred agent fixes

- **Missing source document bootstrap**
  - Symptom: Examples assume source DOC or DOCX files already exist.
  - Preferred fix: Create all source documents locally inside the example before joining them.

- **Wrong join API choice**
  - Symptom: Code invents unsupported APIs or uses append logic where insertion is required.
  - Preferred fix: Use AppendDocument for whole-document appends and InsertDocument or InsertDocumentInline only for real insertion scenarios.

- **Invalid node insertion**
  - Symptom: Runtime failures due to inserting imported sections or nodes into an invalid parent location.
  - Preferred fix: Import content into valid containers only and never append a node to itself.

- **Mail merge insertion misuse**
  - Symptom: Document insertion during mail merge is treated like a plain append or uses invented merge helpers.
  - Preferred fix: Use the documented mail-merge event-handler insertion workflow with InsertDocument or InsertDocumentInline.

- **Weak merged-output validation**
  - Symptom: The example saves output but does not prove the result contains content from all expected sources.
  - Preferred fix: Validate using actual markers, paragraph text, section counts, or output file existence.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`
- Additional package: `Newtonsoft.Json` for reporting tasks when needed

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Newtonsoft.Json
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\join-documents\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and use only documented AppendDocument, InsertDocument, InsertDocumentInline, and mail-merge insertion workflows.
