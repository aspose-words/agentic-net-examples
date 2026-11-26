# Join Documents Examples for Aspose.Words for .NET

The Join Documents category demonstrates how to combine Word processing files using Aspose.Words 26.11.0 while preserving formatting, bookmarks, and protection settings. Verified examples cover insertion at bookmarks, password-protected documents, header/footer insertion, find-replace callbacks, and mail-merge driven document injection.

## Snapshot

- Category: Join Documents
- Slug: join-documents
- Total examples: 36
- Publish-ready successful examples: 36 / 36
- Source run: 20261124_190740_1c00f5
- Insert Document examples: 18
- Append Document examples: 12
- General Join Workflow examples: 1
- Mail Merge Insert examples: 2
- Section And Format Preservation examples: 3

## Category rules that shaped these examples

- Never assume that source files or folders already exist; always generate them within the example.
- Do not call non-existent Aspose.Words methods such as InsertDocument on MailMerge or use invented join APIs.
- Never insert a node into an invalid parent container or append a node to itself.
- Use DocumentBuilder.MoveToBookmark followed by InsertDocument for precise bookmark insertion.
- Use Document.AppendDocument for whole-document concatenation and always specify ImportFormatMode.
- When performing find-replace insertion, employ FindReplaceOptions with a custom IReplacingCallback that imports nodes via NodeImporter.
- Initialize all non-nullable reference-type fields and locals to avoid CS8618 warnings.
- Guard any potentially null values with explicit null checks before dereferencing.
- Declare variables as nullable only when a null value is a legitimate state and handle that case safely.

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
Copy-Item ..\join-documents\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `join-documents/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0

# PowerShell example
Copy-Item ..\join-documents\load-a-docx-template-insert-a-source-docx-at-a-bookmark-and-save-the-result-as-pdf.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-template-insert-a-source-docx-at-a-bookmark-and-save-the-result-as-pdf.cs` | Load a DOCX template, insert a source DOCX at a bookmark, and save the result as PDF. | insert-document | docx, pdf | mcp |
| 2 | `load-a-docx-insert-a-pdf-converted-docx-at-a-bookmark-and-save-the-merged-document-as-odt.cs` | Load a DOCX, insert a PDF-converted DOCX at a bookmark, and save the merged document as ODT. | insert-document | docx, pdf | mcp |
| 3 | `load-a-password-protected-docx-insert-a-source-docx-at-a-bookmark-then-remove-protection-a.cs` | Load a password-protected DOCX, insert a source DOCX at a bookmark, then remove protection and save as DOCX. | insert-document | docx, doc | mcp |
| 4 | `protect-a-destination-docx-with-password-before-inserting-a-source-docx-then-remove-protec.cs` | Protect a destination DOCX with password before inserting a source DOCX, then remove protection and save as DOCX. | insert-document | docx, doc | mcp |
| 5 | `protect-a-destination-docx-with-read-only-restriction-before-appending-source-docx-then-re.cs` | Protect a destination DOCX with read-only restriction before appending source DOCX, then remove restriction and save as PDF. | append-document | docx, pdf | mcp |
| 6 | `use-documentbuilder-movetobookmark-to-navigate-to-a-bookmark-named-content-then-insert-a-s.cs` | Use DocumentBuilder.MoveToBookmark to navigate to a bookmark named "Content", then insert a source DOCX with KeepSourceFormatting. | insert-document | docx | mcp |
| 7 | `insert-a-docx-at-a-bookmark-inside-a-header-then-save-the-document-as-docx-preserving-head.cs` | Insert a DOCX at a bookmark inside a header, then save the document as DOCX preserving header formatting. | insert-document | docx, doc | mcp |
| 8 | `insert-a-docx-at-a-bookmark-inside-a-table-cell-then-export-document-to-odt-preserving-tab.cs` | Insert a DOCX at a bookmark inside a table cell, then export document to ODT preserving table structure. | insert-document | docx | mcp |
| 9 | `insert-a-docx-at-multiple-bookmarks-named-header-and-footer-then-export-the-final-document.cs` | Insert a DOCX at multiple bookmarks named Header and Footer, then export the final document to HTML format. | insert-document | docx, html | mcp |
| 10 | `use-documentbuilder-insertdocument-to-insert-a-docx-at-the-current-cursor-position-with-ke.cs` | Use DocumentBuilder.InsertDocument to insert a DOCX at the current cursor position with KeepSourceFormatting, saving as HTML. | insert-document | docx, html | mcp |
| 11 | `use-documentbuilder-insertdocument-with-importformatmode-usedestinationstyles-to-insert-a.cs` | Use DocumentBuilder.InsertDocument with ImportFormatMode.UseDestinationStyles to insert a DOCX into a styled template, then save as HTML. | insert-document | docx, html | mcp |
| 12 | `insert-a-docx-during-a-find-replace-operation-that-matches-the-word-insertme-then-save-as.cs` | Insert a DOCX during a find-replace operation that matches the word "INSERTME", then save as DOCX. | insert-document | docx, doc | existing_repo |
| 13 | `use-findreplaceoptions-to-replace-the-phrase-placeholder-with-a-docx-document-then-save-as.cs` | Use FindReplaceOptions to replace the phrase "PLACEHOLDER" with a DOCX document, then save as ODT. | general-join-workflow | docx | llm |
| 14 | `use-findreplaceoptions-with-a-replacingcallback-to-insert-a-docx-whenever-phrase-insert-he.cs` | Use FindReplaceOptions with a ReplacingCallback to insert a DOCX whenever phrase "INSERT_HERE" is found, then save as DOCX. | insert-document | docx, doc | mcp |
| 15 | `insert-a-docx-during-a-find-replace-operation-that-replaces-an-xml-tag-then-save-the-docum.cs` | Insert a DOCX during a find-replace operation that replaces an XML tag, then save the document as DOCX. | insert-document | docx, doc, xml | mcp |
| 16 | `set-mailmerge-insertdocument-event-to-inject-a-docx-for-each-data-row-and-generate-a-combi.cs` | Set MailMerge.InsertDocument event to inject a DOCX for each data row, and generate a combined PDF file. | insert-document | docx, pdf | llm |
| 17 | `insert-a-docx-during-mail-merge-for-record-then-apply-importlanguageinfo-to-preserve-langu.cs` | Insert a DOCX during mail merge for record, then apply ImportLanguageInfo to preserve language settings, save as PDF. | mail-merge-insert | docx, pdf | llm |
| 18 | `use-importlanguageinfo-option-to-preserve-language-settings-when-appending-a-source-docx-t.cs` | Use ImportLanguageInfo option to preserve language settings when appending a source DOCX to a destination DOCX. | append-document | docx | existing_repo |
| 19 | `manually-import-specific-tables-from-a-source-docx-using-nodeimporter-and-insert-them-into.cs` | Manually import specific tables from a source DOCX using NodeImporter and insert them into a destination DOCX. | insert-document | docx | mcp |
| 20 | `use-nodeimporter-to-import-only-paragraph-nodes-from-a-source-docx-and-insert-them-after-a.cs` | Use NodeImporter to import only paragraph nodes from a source DOCX and insert them after a specific bookmark. | insert-document | docx | mcp |
| 21 | `append-an-rtf-document-to-an-existing-docx-using-usedestinationstyles-then-export-the-comb.cs` | Append an RTF document to an existing DOCX using UseDestinationStyles, then export the combined file to DOCX. | append-document | docx, doc | mcp |
| 22 | `append-a-docx-containing-custom-styles-to-a-destination-docx-using-importformatmode-usedes.cs` | Append a DOCX containing custom styles to a destination DOCX, using ImportFormatMode.UseDestinationStyles, then export to PDF. | append-document | docx, pdf | mcp |
| 23 | `append-a-docx-generated-from-a-mail-merge-operation-to-an-existing-pdf-converted-docx-pres.cs` | Append a DOCX generated from a mail merge operation to an existing PDF-converted DOCX, preserving destination styles. | mail-merge-insert | docx, pdf | mcp |
| 24 | `append-a-docx-with-video-files-to-a-destination-docx-ensuring-videos-remain-functional-aft.cs` | Append a DOCX with video files to a destination DOCX, ensuring videos remain functional after saving as PDF. | append-document | docx, pdf | llm |
| 25 | `append-a-docx-containing-footnotes-to-a-destination-docx-and-ensure-footnote-numbering-con.cs` | Append a DOCX containing footnotes to a destination DOCX and ensure footnote numbering continues, then save as PDF. | append-document | docx, pdf | llm |
| 26 | `append-a-docx-containing-a-bibliography-to-a-research-paper-docx-then-update-fields-and-ex.cs` | Append a DOCX containing a bibliography to a research paper DOCX, then update fields and export to PDF. | append-document | docx, pdf | mcp |
| 27 | `combine-three-odt-files-into-one-docx-applying-importformatmode-keepsourceformatting-then.cs` | Combine three ODT files into one DOCX, applying ImportFormatMode.KeepSourceFormatting, then save as DOCX. | section-and-format-preservation | docx, doc | mcp |
| 28 | `combine-a-docx-and-an-odt-file-into-a-single-pdf-using-importformatmode-keepsourceformatti.cs` | Combine a DOCX and an ODT file into a single PDF, using ImportFormatMode.KeepSourceFormatting for the ODT content. | section-and-format-preservation | docx, pdf | mcp |
| 29 | `combine-multiple-docx-files-into-a-single-pdf-ensuring-each-source-document-retains-its-or.cs` | Combine multiple DOCX files into a single PDF, ensuring each source document retains its original formatting. | section-and-format-preservation | docx, pdf | mcp |
| 30 | `append-multiple-docx-files-in-a-loop-each-with-different-importformatmode-settings-and-out.cs` | Append multiple DOCX files in a loop, each with different ImportFormatMode settings, and output combined file as PDF. | append-document | docx, pdf | mcp |
| 31 | `batch-process-a-directory-of-docx-files-appending-each-to-a-master-document-and-saving-res.cs` | Batch process a directory of DOCX files, appending each to a master document and saving result as PDF. | append-document | docx, pdf | mcp |
| 32 | `batch-append-docx-files-from-multiple-subfolders-into-a-single-master-document-applying-im.cs` | Batch append DOCX files from multiple subfolders into a single master document, applying ImportFormatMode.UseDestinationStyles, then export to PDF. | append-document | docx, pdf | mcp |
| 33 | `use-documentbuilder-to-insert-a-docx-at-end-of-sections-in-a-destination-docx-then-export.cs` | Use DocumentBuilder to insert a DOCX at end of sections in a destination DOCX, then export to DOCX. | insert-document | docx, doc | mcp |
| 34 | `validate-that-all-inserted-sections-contain-correct-page-numbers-by-iterating-through-dest.cs` | Validate that all inserted sections contain correct page numbers by iterating through destination document nodes after insertion. | insert-document | docx | mcp |
| 35 | `append-a-docx-generated-from-a-web-service-to-an-existing-docx-then-encrypt-final-pdf-with.cs` | Append a DOCX generated from a web service to an existing DOCX, then encrypt final PDF with password. | append-document | docx, pdf | llm |
| 36 | `append-a-docx-generated-from-a-rest-api-to-an-existing-docx-then-encrypt-final-pdf-with-pa.cs` | Append a DOCX generated from a REST API to an existing DOCX, then encrypt final PDF with password. | append-document | docx, pdf | llm |

## Common failure patterns seen during generation and how they were corrected

### Missing source document bootstrap

- Symptom: The example fails to compile because it references source DOCX files that were never created.
- Root cause: The code assumes external files exist without generating them in the sample.
- Preferred fix: Create the required source documents programmatically before attempting to load or join them.
- Prevention: Always include document creation steps at the start of the sample.

### Wrong join API choice

- Symptom: Compilation errors occur when the code calls unsupported methods or uses AppendDocument where InsertDocument is required.
- Root cause: The developer selected an API that does not match the intended operation (e.g., whole-document append vs. bookmark insertion).
- Preferred fix: Use AppendDocument for full-document concatenation and InsertDocument (or InsertDocumentInline) only when inserting at a specific location.
- Prevention: Map the desired join scenario to the correct Aspose.Words API before coding.

### Invalid node insertion

- Symptom: Runtime exceptions indicate that a node of an unsupported type was inserted into a header, footer, or other invalid container.
- Root cause: The code attempts to insert a Section or Body node directly into a Paragraph or Header node.
- Preferred fix: Import only compatible node types (e.g., Paragraph, Run, Table) and insert them into a valid parent such as a Paragraph or Section.
- Prevention: Validate the target container's NodeType before performing InsertAfter/InsertBefore operations.

### Find-replace insertion misuse

- Symptom: Build failures arise from using string overloads of Range.Replace with a document object, and runtime errors occur when the callback inserts nodes incorrectly.
- Root cause: The sample treats a document as a replacement string instead of using a ReplacingCallback that imports nodes.
- Preferred fix: Configure FindReplaceOptions.ReplacingCallback to import nodes via NodeImporter and replace the placeholder run, then remove the placeholder paragraph.
- Prevention: Never pass a Document to the string-based Replace method; always use a callback for document insertion.

### MailMerge InsertDocument missing

- Symptom: Compilation errors show that MailMerge.InsertDocument does not exist, and runtime checks report missing sections after merge.
- Root cause: The code references a non-existent InsertDocument event on MailMerge and does not correctly combine generated documents.
- Preferred fix: Perform mail merge on a template, then append or insert the additional document using DocumentBuilder.InsertDocument within a FieldMergingCallback, or concatenate results with AppendDocument.
- Prevention: Use the documented MailMerge.FieldMergingCallback pattern and combine documents with AppendDocument after each merge iteration.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Keep the sample code aligned with the current Aspose.Words API surface; update namespace imports and method signatures whenever a new version is released.
