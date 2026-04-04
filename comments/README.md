# Comments Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **Comments** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Comments**
- Slug: **comments**
- Total examples: **30**
- Publish-ready successful examples: **30 / 30**
- Native comment API examples: **15**
- Export / report examples: **7**
- Stream / batch / input-bootstrap examples: **3**
- Rendered-output examples: **5**

## Category rules that shaped these examples

- Use native `Aspose.Words.Comment class` APIs directly and prefer simple, verifiable comment operations.
- Enumerate comments with `doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>()`; do not rely on invented `Document.Comments` APIs.
- Create realistic local sample inputs inside the example whenever the task mentions streams, folders, templates, DOC/DOCX files, XML, JSON, or database-like sources.
- Use supported save options only for PDF, XPS, and HTML scenarios.
- Recreate report/export content in destination documents rather than moving source nodes directly.
- Avoid nullable-reference warnings by null-checking maybe-null results before dereference or assignment.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\comments\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `comments/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# PowerShell example
Copy-Item ..\comments\load-a-docx-file-enumerate-all-comments-and-print-each-author-and-text-to-console.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `add-a-comment-containing-a-hyperlink-to-an-external-resource-and-verify-the-link-functions.cs` | Add a comment containing a hyperlink to an external resource and verify the link functions in rendered output. | rendered-output | pdf | verified |
| 2 | `add-a-new-comment-to-a-specific-paragraph-in-a-word-document-and-save-as-docx.cs` | Add a new comment to a specific paragraph in a word document and save as docx. | native-comment-api | docx | verified |
| 3 | `apply-a-custom-style-to-all-comment-text-blocks-within-a-document-to-match-corporate-brand.cs` | Apply a custom style to all comment text blocks within a document to match corporate branding. | native-comment-api | docx | verified |
| 4 | `compare-two-versions-of-a-document-and-list-comments-that-were-added-modified-or-deleted.cs` | Compare two versions of a document and list comments that were added modified or deleted. | native-comment-api | docx | verified |
| 5 | `convert-a-doc-file-to-pdf-while-retaining-all-comment-annotations-visible-in-the-output.cs` | Convert a doc file to pdf while retaining all comment annotations visible in the output. | rendered-output | docx, pdf, doc | verified |
| 6 | `convert-a-document-with-comments-to-xps-format-ensuring-comments-appear-as-markup-annotati.cs` | Convert a document with comments to XPS format, ensuring comments appear as markup annotations. | rendered-output | xps | verified |
| 7 | `create-a-batch-process-that-adds-a-standardized-disclaimer-comment-to-every-document-in-a.cs` | Create a batch process that adds a standardized disclaimer comment to every document in a. | stream-batch-io | docx | verified |
| 8 | `create-a-reply-to-an-existing-comment-and-ensure-the-reply-appears-nested-under-the-origin.cs` | Create a reply to an existing comment and ensure the reply appears nested under the original comment. | native-comment-api | docx | verified |
| 9 | `create-a-utility-that-reads-comment-data-from-a-database-and-inserts-corresponding-comment.cs` | Create a utility that reads comment data from a database and inserts corresponding comments into a document. | stream-batch-io | docx | verified |
| 10 | `delete-all-comments-authored-by-a-particular-user-from-the-loaded-document-before-exportin.cs` | Delete all comments authored by a particular user from the loaded document before exporting. | export-report | docx | verified |
| 11 | `export-all-comments-from-a-docx-file-to-a-csv-file-with-author-date-and-text-columns.cs` | Export all comments from a docx file to a csv file with author date and text columns. | export-report | docx, csv | verified |
| 12 | `extract-comment-metadata-author-date-and-text-and-write-it-to-a-json-file.cs` | Extract comment metadata author date and text and write it to a json file. | export-report | docx, json | verified |
| 13 | `extract-comment-text-and-embed-it-as-footnotes-within-the-same-document-for-alternative-pr.cs` | Extract comment text and embed it as footnotes within the same document for alternative presentation. | native-comment-api | docx | verified |
| 14 | `filter-comments-by-author-and-export-only-those-comments-to-a-separate-word-document-for-r.cs` | Filter comments by author and export only those comments to a separate Word document for review. | export-report | docx | verified |
| 15 | `generate-a-printable-report-listing-all-comments-with-page-numbers-and-associated-paragrap.cs` | Generate a printable report listing all comments with page numbers and associated paragraphs. | export-report | docx | verified |
| 16 | `implement-a-feature-that-hides-all-comments-in-the-document-view-without-removing-them-fro.cs` | Implement a feature that hides all comments in the document view without removing them from the document. | rendered-output | docx, pdf | verified |
| 17 | `import-comments-from-an-exported-xml-file-and-attach-them-to-appropriate-locations-in-a-ne.cs` | Import comments from an exported XML file and attach them to appropriate locations in a new document. | export-report | docx, xml | verified |
| 18 | `iterate-through-the-comment-collection-and-remove-comments-older-than-a-specified-date-thr.cs` | Iterate through the comment collection and remove comments older than a specified date threshold. | native-comment-api | docx | verified |
| 19 | `load-a-document-change-comment-author-names-to-uppercase-and-save-the-updated-file.cs` | Load a document change comment author names to uppercase and save the updated file. | native-comment-api | docx | verified |
| 20 | `load-a-document-from-a-stream-add-comments-and-save-the-modified-document-back-to-a-memory.cs` | Load a document from a stream, add comments, and save the modified document back to a memory stream. | stream-batch-io | docx, doc | verified |
| 21 | `load-a-docx-file-enumerate-all-comments-and-print-each-author-and-text-to-console.cs` | Load a docx file enumerate all comments and print each author and text to console. | native-comment-api | docx | verified |
| 22 | `load-multiple-word-documents-from-a-folder-aggregate-their-comments-and-generate-a-summary.cs` | Load multiple Word documents from a folder, aggregate their comments, and generate a summary report. | export-report | docx | verified |
| 23 | `preserve-comment-formatting-such-as-bold-and-italic-text-when-converting-a-document-to-htm.cs` | Preserve comment formatting such as bold and italic text when converting a document to HTML. | rendered-output | docx, html | verified |
| 24 | `programmatically-accept-or-reject-comments-based-on-author-name-and-generate-a-revised-doc.cs` | Programmatically accept or reject comments based on author name and generate a revised document version. | native-comment-api | docx | verified |
| 25 | `search-comments-containing-a-specific-keyword-and-highlight-the-corresponding-text-range-i.cs` | Search comments containing a specific keyword and highlight the corresponding text range in the document. | native-comment-api | docx | verified |
| 26 | `set-custom-author-name-and-initials-for-programmatically-added-comments-in-a-document.cs` | Set custom author name and initials for programmatically added comments in a document. | native-comment-api | docx | verified |
| 27 | `synchronize-comment-positions-after-document-sections-are-reordered-to-maintain-accurate-c.cs` | Synchronize comment positions after document sections are reordered to maintain accurate comment placement. | native-comment-api | docx | verified |
| 28 | `update-the-text-of-an-existing-comment-identified-by-its-index-while-preserving-original-f.cs` | Update the text of an existing comment identified by its index while preserving original formatting. | native-comment-api | docx | verified |
| 29 | `use-comment-collection-events-to-trigger-custom-logging-whenever-a-comment-is-added-or-rem.cs` | Use comment collection events or equivalent wrapper-based logging to capture add and remove operations. | native-comment-api | docx | verified |
| 30 | `validate-that-comment-reference-ids-update-correctly-after-inserting-new-paragraphs-into-t.cs` | Validate that comment reference IDs update correctly after inserting new paragraphs into the document. | native-comment-api | docx | verified |

## Common failure patterns seen during generation and how they were corrected

### Wrong Node.Text usage

- Symptom: `error CS1061: 'Node' does not contain a definition for 'Text'`
- Fix: Use `Comment.GetText()` when reading full comment content, or cast to the correct derived node type such as `Run` before accessing `.Text`.

### Invented comment collection APIs

- Symptom: code tries to use `Document.Comments` or an invented `Aspose.Words.Comments` namespace.
- Fix: Enumerate comments with `doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>()`.

### Unsupported comment-reference or save-option APIs

- Symptom: code uses `CommentReference` or `HtmlSaveOptions.ExportComments` in unsupported ways.
- Fix: Prefer simpler comment examples that do not require explicit comment-reference construction, and use only documented save options.

### Invalid live-collection mutation

- Symptom: removing comments while iterating forward through a live node collection causes unstable behavior.
- Fix: materialize a safe `.ToList()` copy first, then remove or update matching comments.

### Nullable-reference warnings and null-risk APIs

- Symptom: warnings such as `CS8600`, `CS8602`, or `CS8604` around `FirstOrDefault`, `CurrentParagraph`, `ParentNode`, `FirstParagraph`, or similar values.
- Fix: declare nullable locals when appropriate and guard maybe-null results before dereference or assignment.

## Notes for maintainers

- This category is now **100% publish-ready** for the current run.
- The one late-stage compile issue that still surfaced during verification was `Node.Text`; the corrected pattern is to use `Comment.GetText()` or cast to a concrete node type before reading text.
- Preserve file-to-task traceability when updating this folder.
- For future updates, keep the examples standalone and continue bootstrapping local inputs inside the example whenever external sources are mentioned.
