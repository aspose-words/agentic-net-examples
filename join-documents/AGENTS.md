---
name: join-documents
description: Verified C# examples for Join Documents scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Join Documents

## Purpose

This folder is a live, curated example set for Join Documents scenarios. Treat every `.cs` file as a standalone console application. The goal is correct, warning-free examples that use documented Aspose APIs and match the original task intent.

## Non-negotiable conventions

- Create all required local source documents before joining them.
- Use AppendDocument for whole-document append scenarios.
- Use InsertDocument or InsertDocumentInline only when the task explicitly requires insertion at a specific position or during a mail merge workflow.
- Specify ImportFormatMode explicitly when appending documents.
- Do not invent unsupported document-join APIs.
- Reset MemoryStream.Position before loading from a written stream.
- Validate that the merged output contains content from all expected sources.

## Recommended workflow selection

- Append Document workflow: 12 examples
- General Join Workflow workflow: 1 examples
- Insert Document workflow: 18 examples
- Mail Merge Insert workflow: 2 examples
- Section And Format Preservation workflow: 3 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. API usage must be supported by the configured package versions.
3. Exported outputs must actually be written by the example.
4. Validation scenarios must inspect only the behavior requested by the task.
5. Examples that depend on files, folders, streams, images, or data should bootstrap those inputs locally during the example run.

## File-to-task reference

- `load-a-docx-template-insert-a-source-docx-at-a-bookmark-and-save-the-result-as-pdf.cs`
  - Task: Load a DOCX template, insert a source DOCX at a bookmark, and save the result as PDF.
  - Workflow: Insert Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `load-a-docx-insert-a-pdf-converted-docx-at-a-bookmark-and-save-the-merged-document-as-odt.cs`
  - Task: Load a DOCX, insert a PDF-converted DOCX at a bookmark, and save the merged document as ODT.
  - Workflow: Insert Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `load-a-password-protected-docx-insert-a-source-docx-at-a-bookmark-then-remove-protection-a.cs`
  - Task: Load a password-protected DOCX, insert a source DOCX at a bookmark, then remove protection and save as DOCX.
  - Workflow: Insert Document
  - Outputs: docx, doc
  - Selected engine: mcp
- `protect-a-destination-docx-with-password-before-inserting-a-source-docx-then-remove-protec.cs`
  - Task: Protect a destination DOCX with password before inserting a source DOCX, then remove protection and save as DOCX.
  - Workflow: Insert Document
  - Outputs: docx, doc
  - Selected engine: mcp
- `protect-a-destination-docx-with-read-only-restriction-before-appending-source-docx-then-re.cs`
  - Task: Protect a destination DOCX with read-only restriction before appending source DOCX, then remove restriction and save as PDF.
  - Workflow: Append Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `use-documentbuilder-movetobookmark-to-navigate-to-a-bookmark-named-content-then-insert-a-s.cs`
  - Task: Use DocumentBuilder.MoveToBookmark to navigate to a bookmark named "Content", then insert a source DOCX with KeepSourceFormatting.
  - Workflow: Insert Document
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-docx-at-a-bookmark-inside-a-header-then-save-the-document-as-docx-preserving-head.cs`
  - Task: Insert a DOCX at a bookmark inside a header, then save the document as DOCX preserving header formatting.
  - Workflow: Insert Document
  - Outputs: docx, doc
  - Selected engine: mcp
- `insert-a-docx-at-a-bookmark-inside-a-table-cell-then-export-document-to-odt-preserving-tab.cs`
  - Task: Insert a DOCX at a bookmark inside a table cell, then export document to ODT preserving table structure.
  - Workflow: Insert Document
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-docx-at-multiple-bookmarks-named-header-and-footer-then-export-the-final-document.cs`
  - Task: Insert a DOCX at multiple bookmarks named Header and Footer, then export the final document to HTML format.
  - Workflow: Insert Document
  - Outputs: docx, html
  - Selected engine: mcp
- `use-documentbuilder-insertdocument-to-insert-a-docx-at-the-current-cursor-position-with-ke.cs`
  - Task: Use DocumentBuilder.InsertDocument to insert a DOCX at the current cursor position with KeepSourceFormatting, saving as HTML.
  - Workflow: Insert Document
  - Outputs: docx, html
  - Selected engine: mcp
- `use-documentbuilder-insertdocument-with-importformatmode-usedestinationstyles-to-insert-a.cs`
  - Task: Use DocumentBuilder.InsertDocument with ImportFormatMode.UseDestinationStyles to insert a DOCX into a styled template, then save as HTML.
  - Workflow: Insert Document
  - Outputs: docx, html
  - Selected engine: mcp
- `insert-a-docx-during-a-find-replace-operation-that-matches-the-word-insertme-then-save-as.cs`
  - Task: Insert a DOCX during a find-replace operation that matches the word "INSERTME", then save as DOCX.
  - Workflow: Insert Document
  - Outputs: docx, doc
  - Selected engine: llm
- `use-findreplaceoptions-to-replace-the-phrase-placeholder-with-a-docx-document-then-save-as.cs`
  - Task: Use FindReplaceOptions to replace the phrase "PLACEHOLDER" with a DOCX document, then save as ODT.
  - Workflow: General Join Workflow
  - Outputs: docx
  - Selected engine: mcp
- `use-findreplaceoptions-with-a-replacingcallback-to-insert-a-docx-whenever-phrase-insert-he.cs`
  - Task: Use FindReplaceOptions with a ReplacingCallback to insert a DOCX whenever phrase "INSERT_HERE" is found, then save as DOCX.
  - Workflow: Insert Document
  - Outputs: docx, doc
  - Selected engine: mcp
- `insert-a-docx-during-a-find-replace-operation-that-replaces-an-xml-tag-then-save-the-docum.cs`
  - Task: Insert a DOCX during a find-replace operation that replaces an XML tag, then save the document as DOCX.
  - Workflow: Insert Document
  - Outputs: docx, doc, xml
  - Selected engine: mcp
- `set-mailmerge-insertdocument-event-to-inject-a-docx-for-each-data-row-and-generate-a-combi.cs`
  - Task: Set MailMerge.InsertDocument event to inject a DOCX for each data row, and generate a combined PDF file.
  - Workflow: Insert Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `insert-a-docx-during-mail-merge-for-record-then-apply-importlanguageinfo-to-preserve-langu.cs`
  - Task: Insert a DOCX during mail merge for record, then apply ImportLanguageInfo to preserve language settings, save as PDF.
  - Workflow: Mail Merge Insert
  - Outputs: docx, pdf
  - Selected engine: mcp
- `use-importlanguageinfo-option-to-preserve-language-settings-when-appending-a-source-docx-t.cs`
  - Task: Use ImportLanguageInfo option to preserve language settings when appending a source DOCX to a destination DOCX.
  - Workflow: Append Document
  - Outputs: docx
  - Selected engine: mcp
- `manually-import-specific-tables-from-a-source-docx-using-nodeimporter-and-insert-them-into.cs`
  - Task: Manually import specific tables from a source DOCX using NodeImporter and insert them into a destination DOCX.
  - Workflow: Insert Document
  - Outputs: docx
  - Selected engine: mcp
- `use-nodeimporter-to-import-only-paragraph-nodes-from-a-source-docx-and-insert-them-after-a.cs`
  - Task: Use NodeImporter to import only paragraph nodes from a source DOCX and insert them after a specific bookmark.
  - Workflow: Insert Document
  - Outputs: docx
  - Selected engine: mcp
- `append-an-rtf-document-to-an-existing-docx-using-usedestinationstyles-then-export-the-comb.cs`
  - Task: Append an RTF document to an existing DOCX using UseDestinationStyles, then export the combined file to DOCX.
  - Workflow: Append Document
  - Outputs: docx, doc
  - Selected engine: mcp
- `append-a-docx-containing-custom-styles-to-a-destination-docx-using-importformatmode-usedes.cs`
  - Task: Append a DOCX containing custom styles to a destination DOCX, using ImportFormatMode.UseDestinationStyles, then export to PDF.
  - Workflow: Append Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `append-a-docx-generated-from-a-mail-merge-operation-to-an-existing-pdf-converted-docx-pres.cs`
  - Task: Append a DOCX generated from a mail merge operation to an existing PDF-converted DOCX, preserving destination styles.
  - Workflow: Mail Merge Insert
  - Outputs: docx, pdf
  - Selected engine: mcp
- `append-a-docx-with-video-files-to-a-destination-docx-ensuring-videos-remain-functional-aft.cs`
  - Task: Append a DOCX with video files to a destination DOCX, ensuring videos remain functional after saving as PDF.
  - Workflow: Append Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `append-a-docx-containing-footnotes-to-a-destination-docx-and-ensure-footnote-numbering-con.cs`
  - Task: Append a DOCX containing footnotes to a destination DOCX and ensure footnote numbering continues, then save as PDF.
  - Workflow: Append Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `append-a-docx-containing-a-bibliography-to-a-research-paper-docx-then-update-fields-and-ex.cs`
  - Task: Append a DOCX containing a bibliography to a research paper DOCX, then update fields and export to PDF.
  - Workflow: Append Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `combine-three-odt-files-into-one-docx-applying-importformatmode-keepsourceformatting-then.cs`
  - Task: Combine three ODT files into one DOCX, applying ImportFormatMode.KeepSourceFormatting, then save as DOCX.
  - Workflow: Section And Format Preservation
  - Outputs: docx, doc
  - Selected engine: mcp
- `combine-a-docx-and-an-odt-file-into-a-single-pdf-using-importformatmode-keepsourceformatti.cs`
  - Task: Combine a DOCX and an ODT file into a single PDF, using ImportFormatMode.KeepSourceFormatting for the ODT content.
  - Workflow: Section And Format Preservation
  - Outputs: docx, pdf
  - Selected engine: mcp
- `combine-multiple-docx-files-into-a-single-pdf-ensuring-each-source-document-retains-its-or.cs`
  - Task: Combine multiple DOCX files into a single PDF, ensuring each source document retains its original formatting.
  - Workflow: Section And Format Preservation
  - Outputs: docx, pdf
  - Selected engine: mcp
- `append-multiple-docx-files-in-a-loop-each-with-different-importformatmode-settings-and-out.cs`
  - Task: Append multiple DOCX files in a loop, each with different ImportFormatMode settings, and output combined file as PDF.
  - Workflow: Append Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `batch-process-a-directory-of-docx-files-appending-each-to-a-master-document-and-saving-res.cs`
  - Task: Batch process a directory of DOCX files, appending each to a master document and saving result as PDF.
  - Workflow: Append Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `batch-append-docx-files-from-multiple-subfolders-into-a-single-master-document-applying-im.cs`
  - Task: Batch append DOCX files from multiple subfolders into a single master document, applying ImportFormatMode.UseDestinationStyles, then export to PDF.
  - Workflow: Append Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `use-documentbuilder-to-insert-a-docx-at-end-of-sections-in-a-destination-docx-then-export.cs`
  - Task: Use DocumentBuilder to insert a DOCX at end of sections in a destination DOCX, then export to DOCX.
  - Workflow: Insert Document
  - Outputs: docx, doc
  - Selected engine: mcp
- `validate-that-all-inserted-sections-contain-correct-page-numbers-by-iterating-through-dest.cs`
  - Task: Validate that all inserted sections contain correct page numbers by iterating through destination document nodes after insertion.
  - Workflow: Insert Document
  - Outputs: docx
  - Selected engine: mcp
- `append-a-docx-generated-from-a-web-service-to-an-existing-docx-then-encrypt-final-pdf-with.cs`
  - Task: Append a DOCX generated from a web service to an existing DOCX, then encrypt final PDF with password.
  - Workflow: Append Document
  - Outputs: docx, pdf
  - Selected engine: mcp
- `append-a-docx-generated-from-a-rest-api-to-an-existing-docx-then-encrypt-final-pdf-with-pa.cs`
  - Task: Append a DOCX generated from a REST API to an existing DOCX, then encrypt final PDF with password.
  - Workflow: Append Document
  - Outputs: docx, pdf
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- Missing source document bootstrap
  - Symptom: Examples assume source DOC or DOCX files already exist.
  - Preferred fix: Create all source documents locally inside the example before joining them.

- Wrong join API choice
  - Symptom: Code invents unsupported APIs or uses append logic where insertion is required.
  - Preferred fix: Use AppendDocument for whole-document appends and InsertDocument or InsertDocumentInline only for real insertion scenarios.

- Invalid node insertion
  - Symptom: Runtime failures due to inserting imported sections or nodes into an invalid parent location.
  - Preferred fix: Import content into valid containers only and never append a node to itself.

- Mail merge insertion misuse
  - Symptom: Document insertion during mail merge is treated like a plain append or uses invented merge helpers.
  - Preferred fix: Use the documented mail-merge event-handler insertion workflow with InsertDocument or InsertDocumentInline.

- Weak merged-output validation
  - Symptom: The example saves output but does not prove the result contains content from all expected sources.
  - Preferred fix: Validate using actual markers, paragraph text, section counts, or output file existence.

## Build and run contract

- Target framework: `net8.0`
- Package: `Aspose.Words` `26.5.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.5.0
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
- Keep examples standalone, bootstrap missing inputs locally, and prefer documented Aspose APIs over speculative shortcuts.
