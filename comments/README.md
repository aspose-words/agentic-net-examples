# Comments Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Comments category. Each file is a standalone console example selected from the verified 26.5.0 run.

## Snapshot

- Category: Comments
- Slug: comments
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20260619_131835_59df5f
- Export Report examples: 7
- Native Comment Api examples: 15
- Rendered Output examples: 5
- Stream / batch / input-bootstrap examples: 3

## Category rules that shaped these examples

- Do not use Document.Comments. Enumerate comments with doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>().
- Do not invent an Aspose.Words.Comments namespace or any invented comment collection namespace.
- Do not invent DocumentBuilder.StartComment, DocumentBuilder.EndComment, AddComment, InsertComment, or similar builder shortcut APIs unless the exact Aspose.Words API is known and supported.
- Do not use CommentReference in this category.
- Do not use HtmlSaveOptions.ExportComments.
- Enumerate comments with doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>().
- Append at least one Paragraph and Run inside a Comment when creating comment content.
- Use safe ToList copies before deleting or bulk-updating comments.
- For report/export scenarios, extract plain data first and recreate destination content from that data.
- For stream workflows, reset MemoryStream.Position before reloading or resaving.
- Avoid CS8600, CS8602, and CS8604 by checking for null before assignment, dereference, indexing, or method calls.
- Do not assign maybe-null values directly to non-nullable variables.
- Declare nullable locals such as Paragraph? when a value may legitimately be null.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.5.0

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

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
dotnet add package Aspose.Words --version 26.5.0

# PowerShell example
Copy-Item ..\comments\load-a-docx-file-enumerate-all-comments-and-print-each-author-and-text-to-console.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-enumerate-all-comments-and-print-each-author-and-text-to-console.cs` | Load a DOCX file, enumerate all comments, and print each author and text to console. | Native Comment Api | docx | llm |
| 2 | `add-a-new-comment-to-a-specific-paragraph-in-a-word-document-and-save-as-docx.cs` | Add a new comment to a specific paragraph in a Word document and save as DOCX. | Native Comment Api | docx | mcp |
| 3 | `update-the-text-of-an-existing-comment-identified-by-its-index-while-preserving-original-f.cs` | Update the text of an existing comment identified by its index while preserving original formatting. | Native Comment Api | docx | mcp |
| 4 | `delete-all-comments-authored-by-a-particular-user-from-the-loaded-document-before-exportin.cs` | Delete all comments authored by a particular user from the loaded document before exporting. | Export Report | docx | mcp |
| 5 | `set-custom-author-name-and-initials-for-programmatically-added-comments-in-a-document.cs` | Set custom author name and initials for programmatically added comments in a document. | Native Comment Api | docx | mcp |
| 6 | `create-a-reply-to-an-existing-comment-and-ensure-the-reply-appears-nested-under-the-origin.cs` | Create a reply to an existing comment and ensure the reply appears nested under the original. | Native Comment Api | docx | mcp |
| 7 | `iterate-through-the-comment-collection-and-remove-comments-older-than-a-specified-date-thr.cs` | Iterate through the comment collection and remove comments older than a specified date threshold. | Native Comment Api | docx | mcp |
| 8 | `filter-comments-by-author-and-export-only-those-comments-to-a-separate-word-document-for-r.cs` | Filter comments by author and export only those comments to a separate Word document for review. | Export Report | docx | mcp |
| 9 | `export-all-comments-from-a-docx-file-to-a-csv-file-with-author-date-and-text-columns.cs` | Export all comments from a DOCX file to a CSV file with author, date, and text columns. | Export Report | docx, csv | mcp |
| 10 | `import-comments-from-an-exported-xml-file-and-attach-them-to-appropriate-locations-in-a-ne.cs` | Import comments from an exported XML file and attach them to appropriate locations in a new document. | Export Report | docx, xml | mcp |
| 11 | `extract-comment-metadata-author-date-and-text-and-write-it-to-a-json-file.cs` | Extract comment metadata-author, date, and text-and write it to a JSON file. | Export Report | docx, json | mcp |
| 12 | `search-comments-containing-a-specific-keyword-and-highlight-the-corresponding-text-range-i.cs` | Search comments containing a specific keyword and highlight the corresponding text range in the document. | Native Comment Api | docx | mcp |
| 13 | `load-multiple-word-documents-from-a-folder-aggregate-their-comments-and-generate-a-summary.cs` | Load multiple Word documents from a folder, aggregate their comments, and generate a summary report. | Export Report | docx | mcp |
| 14 | `generate-a-printable-report-listing-all-comments-with-page-numbers-and-associated-paragrap.cs` | Generate a printable report listing all comments with page numbers and associated paragraph text. | Export Report | docx | mcp |
| 15 | `apply-a-custom-style-to-all-comment-text-blocks-within-a-document-to-match-corporate-brand.cs` | Apply a custom style to all comment text blocks within a document to match corporate branding. | Native Comment Api | docx | mcp |
| 16 | `preserve-comment-formatting-such-as-bold-and-italic-text-when-converting-a-document-to-htm.cs` | Preserve comment formatting such as bold and italic text when converting a document to HTML format. | Rendered Output | docx, html | mcp |
| 17 | `convert-a-doc-file-to-pdf-while-retaining-all-comment-annotations-visible-in-the-output.cs` | Convert a DOC file to PDF while retaining all comment annotations visible in the output. | Rendered Output | docx, pdf, doc | mcp |
| 18 | `convert-a-document-with-comments-to-xps-format-ensuring-comments-appear-as-markup-annotati.cs` | Convert a document with comments to XPS format, ensuring comments appear as markup annotations. | Rendered Output | xps | mcp |
| 19 | `add-a-comment-containing-a-hyperlink-to-an-external-resource-and-verify-the-link-functions.cs` | Add a comment containing a hyperlink to an external resource and verify the link functions in PDF. | Rendered Output | pdf | mcp |
| 20 | `validate-that-comment-reference-ids-update-correctly-after-inserting-new-paragraphs-into-t.cs` | Validate that comment reference IDs update correctly after inserting new paragraphs into the document. | Native Comment Api | docx | mcp |
| 21 | `synchronize-comment-positions-after-document-sections-are-reordered-to-maintain-accurate-c.cs` | Synchronize comment positions after document sections are reordered to maintain accurate comment anchoring. | Native Comment Api | docx | mcp |
| 22 | `use-comment-collection-events-to-trigger-custom-logging-whenever-a-comment-is-added-or-rem.cs` | Use comment collection events to trigger custom logging whenever a comment is added or removed. | Native Comment Api | docx | mcp |
| 23 | `programmatically-accept-or-reject-comments-based-on-author-name-and-generate-a-revised-doc.cs` | Programmatically accept or reject comments based on author name and generate a revised document version. | Native Comment Api | docx | mcp |
| 24 | `load-a-document-from-a-stream-add-comments-and-save-the-modified-document-back-to-a-memory.cs` | Load a document from a stream, add comments, and save the modified document back to a memory stream. | Stream / batch / input-bootstrap | docx, doc | mcp |
| 25 | `create-a-batch-process-that-adds-a-standardized-disclaimer-comment-to-every-document-in-a.cs` | Create a batch process that adds a standardized disclaimer comment to every document in a directory. | Stream / batch / input-bootstrap | docx | mcp |
| 26 | `load-a-document-change-comment-author-names-to-uppercase-and-save-the-updated-file.cs` | Load a document, change comment author names to uppercase, and save the updated file. | Native Comment Api | docx | mcp |
| 27 | `extract-comment-text-and-embed-it-as-footnotes-within-the-same-document-for-alternative-pr.cs` | Extract comment text and embed it as footnotes within the same document for alternative presentation. | Native Comment Api | docx | llm |
| 28 | `implement-a-feature-that-hides-all-comments-in-the-document-view-without-removing-them-fro.cs` | Implement a feature that hides all comments in the document view without removing them from the file. | Rendered Output | docx, pdf | mcp |
| 29 | `compare-two-versions-of-a-document-and-list-comments-that-were-added-modified-or-deleted.cs` | Compare two versions of a document and list comments that were added, modified, or deleted. | Native Comment Api | docx | mcp |
| 30 | `create-a-utility-that-reads-comment-data-from-a-database-and-inserts-corresponding-comment.cs` | Create a utility that reads comment data from a database and inserts corresponding comments into a template document. | Stream / batch / input-bootstrap | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Wrong Node.Text usage

- Symptom: error CS1061: 'Node' does not contain a definition for 'Text'
- Fix: Use Comment.GetText() for whole-comment content, or cast to a concrete node type such as Run before reading .Text.

### Invented comment collection APIs

- Symptom: Compile failures caused by Document.Comments or Aspose.Words.Comments usage.
- Fix: Enumerate comments with doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>().

### Unsupported comment-reference or save-option APIs

- Symptom: Compile failures caused by CommentReference or undocumented save-option flags.
- Fix: Prefer simpler comment workflows and documented Aspose.Words save options only.

### Unsafe live-collection mutation

- Symptom: Unstable behavior while deleting comments during forward iteration over a live collection.
- Fix: Create a ToList copy first, then update or remove matching comments.

### Nullable-reference warnings

- Symptom: CS8600, CS8602, or CS8604 around maybe-null values such as CurrentParagraph or FirstOrDefault results.
- Fix: Use nullable locals where needed and guard maybe-null values before dereference or assignment.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.5.0 run.
- Preserve file-to-task traceability when updating this folder.
- Keep examples standalone and bootstrap local inputs inside the example whenever external sources are mentioned.
