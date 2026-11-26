---
name: join-documents
description: Verified C# examples for join-documents scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Join Documents

## Purpose

Provide an AI coding assistant with clear, reproducible patterns for safely joining, inserting, and appending documents in Aspose.Words, ensuring that all required resources are created locally, APIs are used correctly, and the merged output is validated for expected content.

## Non-negotiable conventions

- Never assume that source files or folders already exist; always generate them within the example.
- Do not call non-existent Aspose.Words methods such as InsertDocument on MailMerge or use invented join APIs.
- Never insert a node into an invalid parent container or append a node to itself.
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

1. After each join operation, confirm that the output file exists and that its text includes a unique marker from every source document.

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
  - Selected engine: llm
- `use-findreplaceoptions-with-a-replacingcallback-to-insert-a-docx-whenever-phrase-insert-he.cs`
  - Task: Use FindReplaceOptions with a ReplacingCallback to insert a DOCX whenever phrase "INSERT_HERE" is found, then save as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc
  - Selected engine: mcp
- `insert-a-docx-during-a-find-replace-operation-that-replaces-an-xml-tag-then-save-the-docum.cs`
  - Task: Insert a DOCX during a find-replace operation that replaces an XML tag, then save the document as DOCX.
  - Workflow: insert-document
  - Outputs: docx, doc, xml
  - Selected engine: mcp
- `set-mailmerge-insertdocument-event-to-inject-a-docx-for-each-data-row-and-generate-a-combi.cs`
  - Task: Set MailMerge.InsertDocument event to inject a DOCX for each data row, and generate a combined PDF file.
  - Workflow: insert-document
  - Outputs: docx, pdf
  - Selected engine: llm
- `insert-a-docx-during-mail-merge-for-record-then-apply-importlanguageinfo-to-preserve-langu.cs`
  - Task: Insert a DOCX during mail merge for record, then apply ImportLanguageInfo to preserve language settings, save as PDF.
  - Workflow: mail-merge-insert
  - Outputs: docx, pdf
  - Selected engine: llm
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
  - Selected engine: llm
- `append-a-docx-containing-footnotes-to-a-destination-docx-and-ensure-footnote-numbering-con.cs`
  - Task: Append a DOCX containing footnotes to a destination DOCX and ensure footnote numbering continues, then save as PDF.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: llm
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
  - Selected engine: llm
- `append-a-docx-generated-from-a-rest-api-to-an-existing-docx-then-encrypt-final-pdf-with-pa.cs`
  - Task: Append a DOCX generated from a REST API to an existing DOCX, then encrypt final PDF with password.
  - Workflow: append-document
  - Outputs: docx, pdf
  - Selected engine: llm

## Common failure patterns and preferred agent fixes

- **Missing source document bootstrap**
  - Symptom: The example fails to compile because it references source DOCX files that were never created.
  - Root cause: The code assumes external files exist without generating them in the sample.
  - Preferred fix: Create the required source documents programmatically before attempting to load or join them.
  - Prevention: Always include document creation steps at the start of the sample.

- **Wrong join API choice**
  - Symptom: Compilation errors occur when the code calls unsupported methods or uses AppendDocument where InsertDocument is required.
  - Root cause: The developer selected an API that does not match the intended operation (e.g., whole-document append vs. bookmark insertion).
  - Preferred fix: Use AppendDocument for full-document concatenation and InsertDocument (or InsertDocumentInline) only when inserting at a specific location.
  - Prevention: Map the desired join scenario to the correct Aspose.Words API before coding.

- **Invalid node insertion**
  - Symptom: Runtime exceptions indicate that a node of an unsupported type was inserted into a header, footer, or other invalid container.
  - Root cause: The code attempts to insert a Section or Body node directly into a Paragraph or Header node.
  - Preferred fix: Import only compatible node types (e.g., Paragraph, Run, Table) and insert them into a valid parent such as a Paragraph or Section.
  - Prevention: Validate the target container's NodeType before performing InsertAfter/InsertBefore operations.

- **Find-replace insertion misuse**
  - Symptom: Build failures arise from using string overloads of Range.Replace with a document object, and runtime errors occur when the callback inserts nodes incorrectly.
  - Root cause: The sample treats a document as a replacement string instead of using a ReplacingCallback that imports nodes.
  - Preferred fix: Configure FindReplaceOptions.ReplacingCallback to import nodes via NodeImporter and replace the placeholder run, then remove the placeholder paragraph.
  - Prevention: Never pass a Document to the string-based Replace method; always use a callback for document insertion.

- **MailMerge InsertDocument missing**
  - Symptom: Compilation errors show that MailMerge.InsertDocument does not exist, and runtime checks report missing sections after merge.
  - Root cause: The code references a non-existent InsertDocument event on MailMerge and does not correctly combine generated documents.
  - Preferred fix: Perform mail merge on a template, then append or insert the additional document using DocumentBuilder.InsertDocument within a FieldMergingCallback, or concatenate results with AppendDocument.
  - Prevention: Use the documented MailMerge.FieldMergingCallback pattern and combine documents with AppendDocument after each merge iteration.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
```

```powershell
Copy-Item ..\join-documents\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Keep the sample code aligned with the current Aspose.Words API surface; update namespace imports and method signatures whenever a new version is released.
