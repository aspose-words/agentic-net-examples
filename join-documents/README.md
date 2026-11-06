# Join Documents Examples for Aspose.Words for .NET

The Join Documents category demonstrates how to combine, insert, and append Word processing files using Aspose.Words 26.10.0. Verified examples show correct creation of source files, use of DocumentBuilder.MoveToBookmark, InsertDocument, AppendDocument, and proper saving to formats such as ODT, PDF, and HTML.

## Snapshot

- Category: Join Documents
- Slug: join-documents
- Total examples: 36
- Publish-ready successful examples: 36 / 36
- Source run: 20261102_031946_08fb5b
- Insert Document examples: 18
- Append Document examples: 12
- General Join Workflow examples: 1
- Mail Merge Insert examples: 2
- Section And Format Preservation examples: 3

## Category rules that shaped these examples

- Never reference members that are not part of the Aspose.Words public API, such as BookmarkCollection.Exists or non-existent InsertDocument overloads.
- Do not attempt to append or insert a node into a parent that cannot accept that node type.
- Generate all required DOCX or ODT files before any join operation.
- Use Document.AppendDocument for whole-document concatenation and DocumentBuilder.InsertDocument for bookmark-based insertion, always specifying ImportFormatMode.
- After merging, check that the output file exists and that its text contains known strings from each source document.
- Initialize all reference-type variables before use to satisfy the compiler's non-nullable analysis.
- Guard any potentially null objects (e.g., bookmarks, runs) with explicit null checks before dereferencing.
- Declare locals as nullable when a null value is a legitimate outcome and handle the case accordingly.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.10.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0

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
dotnet add package Aspose.Words --version 26.10.0

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
| 13 | `use-findreplaceoptions-to-replace-the-phrase-placeholder-with-a-docx-document-then-save-as.cs` | Use FindReplaceOptions to replace the phrase "PLACEHOLDER" with a DOCX document, then save as ODT. | general-join-workflow | docx | existing_repo |
| 14 | `use-findreplaceoptions-with-a-replacingcallback-to-insert-a-docx-whenever-phrase-insert-he.cs` | Use FindReplaceOptions with a ReplacingCallback to insert a DOCX whenever phrase "INSERT_HERE" is found, then save as DOCX. | insert-document | docx, doc | mcp |
| 15 | `insert-a-docx-during-a-find-replace-operation-that-replaces-an-xml-tag-then-save-the-docum.cs` | Insert a DOCX during a find-replace operation that replaces an XML tag, then save the document as DOCX. | insert-document | docx, doc, xml | existing_repo |
| 16 | `set-mailmerge-insertdocument-event-to-inject-a-docx-for-each-data-row-and-generate-a-combi.cs` | Set MailMerge.InsertDocument event to inject a DOCX for each data row, and generate a combined PDF file. | insert-document | docx, pdf | existing_repo |
| 17 | `insert-a-docx-during-mail-merge-for-record-then-apply-importlanguageinfo-to-preserve-langu.cs` | Insert a DOCX during mail merge for record, then apply ImportLanguageInfo to preserve language settings, save as PDF. | mail-merge-insert | docx, pdf | mcp |
| 18 | `use-importlanguageinfo-option-to-preserve-language-settings-when-appending-a-source-docx-t.cs` | Use ImportLanguageInfo option to preserve language settings when appending a source DOCX to a destination DOCX. | append-document | docx | existing_repo |
| 19 | `manually-import-specific-tables-from-a-source-docx-using-nodeimporter-and-insert-them-into.cs` | Manually import specific tables from a source DOCX using NodeImporter and insert them into a destination DOCX. | insert-document | docx | mcp |
| 20 | `use-nodeimporter-to-import-only-paragraph-nodes-from-a-source-docx-and-insert-them-after-a.cs` | Use NodeImporter to import only paragraph nodes from a source DOCX and insert them after a specific bookmark. | insert-document | docx | mcp |
| 21 | `append-an-rtf-document-to-an-existing-docx-using-usedestinationstyles-then-export-the-comb.cs` | Append an RTF document to an existing DOCX using UseDestinationStyles, then export the combined file to DOCX. | append-document | docx, doc | mcp |
| 22 | `append-a-docx-containing-custom-styles-to-a-destination-docx-using-importformatmode-usedes.cs` | Append a DOCX containing custom styles to a destination DOCX, using ImportFormatMode.UseDestinationStyles, then export to PDF. | append-document | docx, pdf | mcp |
| 23 | `append-a-docx-generated-from-a-mail-merge-operation-to-an-existing-pdf-converted-docx-pres.cs` | Append a DOCX generated from a mail merge operation to an existing PDF-converted DOCX, preserving destination styles. | mail-merge-insert | docx, pdf | mcp |
| 24 | `append-a-docx-with-video-files-to-a-destination-docx-ensuring-videos-remain-functional-aft.cs` | Append a DOCX with video files to a destination DOCX, ensuring videos remain functional after saving as PDF. | append-document | docx, pdf | mcp |
| 25 | `append-a-docx-containing-footnotes-to-a-destination-docx-and-ensure-footnote-numbering-con.cs` | Append a DOCX containing footnotes to a destination DOCX and ensure footnote numbering continues, then save as PDF. | append-document | docx, pdf | existing_repo |
| 26 | `append-a-docx-containing-a-bibliography-to-a-research-paper-docx-then-update-fields-and-ex.cs` | Append a DOCX containing a bibliography to a research paper DOCX, then update fields and export to PDF. | append-document | docx, pdf | mcp |
| 27 | `combine-three-odt-files-into-one-docx-applying-importformatmode-keepsourceformatting-then.cs` | Combine three ODT files into one DOCX, applying ImportFormatMode.KeepSourceFormatting, then save as DOCX. | section-and-format-preservation | docx, doc | mcp |
| 28 | `combine-a-docx-and-an-odt-file-into-a-single-pdf-using-importformatmode-keepsourceformatti.cs` | Combine a DOCX and an ODT file into a single PDF, using ImportFormatMode.KeepSourceFormatting for the ODT content. | section-and-format-preservation | docx, pdf | mcp |
| 29 | `combine-multiple-docx-files-into-a-single-pdf-ensuring-each-source-document-retains-its-or.cs` | Combine multiple DOCX files into a single PDF, ensuring each source document retains its original formatting. | section-and-format-preservation | docx, pdf | mcp |
| 30 | `append-multiple-docx-files-in-a-loop-each-with-different-importformatmode-settings-and-out.cs` | Append multiple DOCX files in a loop, each with different ImportFormatMode settings, and output combined file as PDF. | append-document | docx, pdf | mcp |
| 31 | `batch-process-a-directory-of-docx-files-appending-each-to-a-master-document-and-saving-res.cs` | Batch process a directory of DOCX files, appending each to a master document and saving result as PDF. | append-document | docx, pdf | mcp |
| 32 | `batch-append-docx-files-from-multiple-subfolders-into-a-single-master-document-applying-im.cs` | Batch append DOCX files from multiple subfolders into a single master document, applying ImportFormatMode.UseDestinationStyles, then export to PDF. | append-document | docx, pdf | mcp |
| 33 | `use-documentbuilder-to-insert-a-docx-at-end-of-sections-in-a-destination-docx-then-export.cs` | Use DocumentBuilder to insert a DOCX at end of sections in a destination DOCX, then export to DOCX. | insert-document | docx, doc | mcp |
| 34 | `validate-that-all-inserted-sections-contain-correct-page-numbers-by-iterating-through-dest.cs` | Validate that all inserted sections contain correct page numbers by iterating through destination document nodes after insertion. | insert-document | docx | mcp |
| 35 | `append-a-docx-generated-from-a-web-service-to-an-existing-docx-then-encrypt-final-pdf-with.cs` | Append a DOCX generated from a web service to an existing DOCX, then encrypt final PDF with password. | append-document | docx, pdf | mcp |
| 36 | `append-a-docx-generated-from-a-rest-api-to-an-existing-docx-then-encrypt-final-pdf-with-pa.cs` | Append a DOCX generated from a REST API to an existing DOCX, then encrypt final PDF with password. | append-document | docx, pdf | llm |

## Common failure patterns seen during generation and how they were corrected

### Nonexistent API member

- Symptom: Compilation fails with CS1061, CS0246, or CS0234 because a referenced type or member cannot be found.
- Root cause: The code references classes, namespaces, or members that are not part of the Aspose.Words 26.10.0 API (e.g., BookmarkCollection.Exists, Aspose.Words.Protection, Import namespace, FindReplaceOptions.ReplacementDocument).
- Preferred fix: Replace the invalid reference with the correct API member (e.g., use mainDoc.Range.Bookmarks["name"] to test existence, use Document.Protect for protection, and use InsertDocument on DocumentBuilder).
- Prevention: Consult the official Aspose.Words documentation for the exact class and member names before coding and rely on IntelliSense to verify availability.

### Invalid node insertion location

- Symptom: Runtime throws System.ArgumentException stating that a node cannot be inserted at the current location.
- Root cause: InsertDocument is called on a node that cannot accept the imported content, such as inserting a whole section into a Run or inserting into a bookmark that has been removed.
- Preferred fix: Move the DocumentBuilder to a valid container (e.g., a Paragraph or the bookmark itself) before calling InsertDocument, and ensure the bookmark remains in the document.
- Prevention: Always verify the target node type with Node.NodeType and use DocumentBuilder.MoveToBookmark or MoveTo(Run) to position the builder correctly.

### Improper FindReplaceOptions usage

- Symptom: Compilation or runtime errors occur when attempting to replace a placeholder with a document using FindReplaceOptions.
- Root cause: The code tries to assign a Document to FindReplaceOptions.ReplacementDocument or passes a Document to Range.Replace, which expects a string.
- Preferred fix: Implement a custom IReplacingCallback that inserts the document at the match location, and pass that callback via FindReplaceOptions.ReplacingCallback.
- Prevention: Use the documented pattern for document insertion during find-replace: create a callback class that receives the match node and calls DocumentBuilder.InsertDocument.

### Incorrect protection API reference

- Symptom: Compilation fails because Protection or InsertDocument members are missing on the MailMerge or Document classes.
- Root cause: The sample uses outdated or non-existent protection APIs (e.g., Aspose.Words.Protection namespace) and assumes a MailMerge.InsertDocument event that does not exist.
- Preferred fix: Apply protection with Document.Protect(ProtectionType, password) and remove it with Document.Unprotect(password); handle mail-merge insertion via MailMerge.FieldMergingCallback and DocumentBuilder.InsertDocument.
- Prevention: Reference the current protection methods in the Document class and follow the official mail-merge insertion workflow.

### Missing source document creation

- Symptom: The example assumes that source DOCX files already exist, leading to FileNotFoundException or empty merges.
- Root cause: No code is provided to generate the prerequisite documents before the join operation.
- Preferred fix: Add code that creates each source document with DocumentBuilder, saves it to a temporary path, and then loads it for insertion or appending.
- Prevention: Always bootstrap required source files at the start of the sample, using clearly named temporary directories.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Keep the sample code in sync with the Aspose.Words version used for verification and update any API changes in future releases.
