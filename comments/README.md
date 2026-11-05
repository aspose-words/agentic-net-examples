# Comments Examples for Aspose.Words for .NET

The Comments category covers all operations that create, enumerate, modify, and export Aspose.Words comment nodes. Verified examples demonstrate reliable techniques for adding top-level comments, replying, filtering by author, exporting to CSV, highlighting, and generating printable reports.

## Snapshot

- Category: Comments
- Slug: comments
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261102_031946_08fb5b
- Native Comment API examples: 15
- Export Report examples: 7
- Rendered Output examples: 5
- Stream and batch I/O examples: 3

## Category rules that shaped these examples

- Do not reference a non-existent Document.Comments collection or an invented Aspose.Words.Comments namespace.
- Do not call DocumentBuilder.StartComment, EndComment, InsertComment, AddComment, or any similar shortcut that is not part of the official API.
- Do not use CommentReference, HtmlSaveOptions.ExportComments, or any undocumented save-option flags.
- Enumerate comments with doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>().
- When creating a comment, always add at least one Paragraph and one Run to hold visible text.
- Create a ToList() snapshot before deleting or bulk-updating comments to avoid mutating a live collection.
- Declare locals as nullable (e.g., Paragraph? para) when a value may be absent and check for null before dereferencing.
- Guard results of FirstOrDefault, CurrentParagraph, or similar queries with explicit null checks.
- Avoid assigning maybe-null expressions directly to non-nullable variables.

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
dotnet add package Aspose.Words --version 26.10.0

# PowerShell example
Copy-Item ..\comments\load-a-docx-file-enumerate-all-comments-and-print-each-author-and-text-to-console.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-enumerate-all-comments-and-print-each-author-and-text-to-console.cs` | Load a DOCX file, enumerate all comments, and print each author and text to console. | native-comment-api | docx | mcp |
| 2 | `add-a-new-comment-to-a-specific-paragraph-in-a-word-document-and-save-as-docx.cs` | Add a new comment to a specific paragraph in a Word document and save as DOCX. | native-comment-api | docx | mcp |
| 3 | `update-the-text-of-an-existing-comment-identified-by-its-index-while-preserving-original-f.cs` | Update the text of an existing comment identified by its index while preserving original formatting. | native-comment-api | docx | mcp |
| 4 | `delete-all-comments-authored-by-a-particular-user-from-the-loaded-document-before-exportin.cs` | Delete all comments authored by a particular user from the loaded document before exporting. | export-report | docx | mcp |
| 5 | `set-custom-author-name-and-initials-for-programmatically-added-comments-in-a-document.cs` | Set custom author name and initials for programmatically added comments in a document. | native-comment-api | docx | mcp |
| 6 | `create-a-reply-to-an-existing-comment-and-ensure-the-reply-appears-nested-under-the-origin.cs` | Create a reply to an existing comment and ensure the reply appears nested under the original. | native-comment-api | docx | existing_repo |
| 7 | `iterate-through-the-comment-collection-and-remove-comments-older-than-a-specified-date-thr.cs` | Iterate through the comment collection and remove comments older than a specified date threshold. | native-comment-api | docx | mcp |
| 8 | `filter-comments-by-author-and-export-only-those-comments-to-a-separate-word-document-for-r.cs` | Filter comments by author and export only those comments to a separate Word document for review. | export-report | docx | mcp |
| 9 | `export-all-comments-from-a-docx-file-to-a-csv-file-with-author-date-and-text-columns.cs` | Export all comments from a DOCX file to a CSV file with author, date, and text columns. | export-report | docx, csv | mcp |
| 10 | `import-comments-from-an-exported-xml-file-and-attach-them-to-appropriate-locations-in-a-ne.cs` | Import comments from an exported XML file and attach them to appropriate locations in a new document. | export-report | docx, xml | mcp |
| 11 | `extract-comment-metadata-author-date-and-text-and-write-it-to-a-json-file.cs` | Extract comment metadata-author, date, and text-and write it to a JSON file. | export-report | docx, json | mcp |
| 12 | `search-comments-containing-a-specific-keyword-and-highlight-the-corresponding-text-range-i.cs` | Search comments containing a specific keyword and highlight the corresponding text range in the document. | native-comment-api | docx | mcp |
| 13 | `load-multiple-word-documents-from-a-folder-aggregate-their-comments-and-generate-a-summary.cs` | Load multiple Word documents from a folder, aggregate their comments, and generate a summary report. | export-report | docx | mcp |
| 14 | `generate-a-printable-report-listing-all-comments-with-page-numbers-and-associated-paragrap.cs` | Generate a printable report listing all comments with page numbers and associated paragraph text. | export-report | docx | mcp |
| 15 | `apply-a-custom-style-to-all-comment-text-blocks-within-a-document-to-match-corporate-brand.cs` | Apply a custom style to all comment text blocks within a document to match corporate branding. | native-comment-api | docx | mcp |
| 16 | `preserve-comment-formatting-such-as-bold-and-italic-text-when-converting-a-document-to-htm.cs` | Preserve comment formatting such as bold and italic text when converting a document to HTML format. | rendered-output | docx, html | mcp |
| 17 | `convert-a-doc-file-to-pdf-while-retaining-all-comment-annotations-visible-in-the-output.cs` | Convert a DOC file to PDF while retaining all comment annotations visible in the output. | rendered-output | docx, pdf, doc | mcp |
| 18 | `convert-a-document-with-comments-to-xps-format-ensuring-comments-appear-as-markup-annotati.cs` | Convert a document with comments to XPS format, ensuring comments appear as markup annotations. | rendered-output | xps | mcp |
| 19 | `add-a-comment-containing-a-hyperlink-to-an-external-resource-and-verify-the-link-functions.cs` | Add a comment containing a hyperlink to an external resource and verify the link functions in PDF. | rendered-output | pdf | mcp |
| 20 | `validate-that-comment-reference-ids-update-correctly-after-inserting-new-paragraphs-into-t.cs` | Validate that comment reference IDs update correctly after inserting new paragraphs into the document. | native-comment-api | docx | mcp |
| 21 | `synchronize-comment-positions-after-document-sections-are-reordered-to-maintain-accurate-c.cs` | Synchronize comment positions after document sections are reordered to maintain accurate comment anchoring. | native-comment-api | docx | mcp |
| 22 | `use-comment-collection-events-to-trigger-custom-logging-whenever-a-comment-is-added-or-rem.cs` | Use comment collection events to trigger custom logging whenever a comment is added or removed. | native-comment-api | docx | llm |
| 23 | `programmatically-accept-or-reject-comments-based-on-author-name-and-generate-a-revised-doc.cs` | Programmatically accept or reject comments based on author name and generate a revised document version. | native-comment-api | docx | mcp |
| 24 | `load-a-document-from-a-stream-add-comments-and-save-the-modified-document-back-to-a-memory.cs` | Load a document from a stream, add comments, and save the modified document back to a memory stream. | stream-batch-io | docx, doc | mcp |
| 25 | `create-a-batch-process-that-adds-a-standardized-disclaimer-comment-to-every-document-in-a.cs` | Create a batch process that adds a standardized disclaimer comment to every document in a directory. | stream-batch-io | docx | mcp |
| 26 | `load-a-document-change-comment-author-names-to-uppercase-and-save-the-updated-file.cs` | Load a document, change comment author names to uppercase, and save the updated file. | native-comment-api | docx | mcp |
| 27 | `extract-comment-text-and-embed-it-as-footnotes-within-the-same-document-for-alternative-pr.cs` | Extract comment text and embed it as footnotes within the same document for alternative presentation. | native-comment-api | docx | mcp |
| 28 | `implement-a-feature-that-hides-all-comments-in-the-document-view-without-removing-them-fro.cs` | Implement a feature that hides all comments in the document view without removing them from the file. | rendered-output | docx, pdf | mcp |
| 29 | `compare-two-versions-of-a-document-and-list-comments-that-were-added-modified-or-deleted.cs` | Compare two versions of a document and list comments that were added, modified, or deleted. | native-comment-api | docx | mcp |
| 30 | `create-a-utility-that-reads-comment-data-from-a-database-and-inserts-corresponding-comment.cs` | Create a utility that reads comment data from a database and inserts corresponding comments into a template document. | stream-batch-io | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Use of non-existent comment APIs

- Symptom: Compile errors such as CS0234 or CS1061 referencing Document.Comments, Aspose.Words.Comments, DocumentBuilder.StartComment/EndComment, or HtmlSaveOptions.ExportComments.
- Root cause: The referenced members are not defined in the Aspose.Words 26.10.0 API.
- Preferred fix: Replace invented calls with the supported pattern doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>() and manipulate Comment objects directly.
- Prevention: Consult the official Aspose.Words documentation and restrict code to members that appear in IntelliSense.

### Mutating a live comment collection

- Symptom: Runtime NotSupportedException stating that adding nodes is not supported for deep node collections, or unstable behavior when deleting during forward iteration.
- Root cause: NodeCollection.Add or removal is performed on the original comment collection, which Aspose.Words treats as read-only for deep hierarchies.
- Preferred fix: Create a snapshot with .ToList() and perform all Add, Remove, or Update operations on that list.
- Prevention: Never iterate over a live NodeCollection while modifying it; always work on a copied list.

### Incorrect member access on base Node

- Symptom: Compile errors CS1061 for members like .Text, .InsertAfter, or .ParentComment on a Node instance.
- Root cause: Node is an abstract base class; those members exist only on derived types such as Run, Paragraph, or Comment.
- Preferred fix: Cast to the appropriate derived type or use Comment.GetText() for comment content; use Paragraph.InsertAfter only on Paragraph objects.
- Prevention: Inspect the concrete type before accessing members and rely on type-specific APIs.

### Nullable reference warnings

- Symptom: CS8600, CS8602, or CS8604 warnings about possible null dereference or assignment.
- Root cause: Code assigns or dereferences values that may be null without null checks.
- Preferred fix: Declare variables as nullable, add explicit null checks, and use the null-conditional operator where appropriate.
- Prevention: Enable nullable reference types and consistently validate results of queries that can return null.

### Improper reply handling

- Symptom: Compile errors for missing Comment.ParentComment property and runtime NotSupportedException when trying to add a reply via NodeCollection.Add.
- Root cause: Replies must be added using the Comment.AddReply method; treating replies as regular nodes in a deep collection is unsupported.
- Preferred fix: Create a reply with topComment.AddReply(author, initials, date, text) and avoid manual insertion into the comment collection.
- Prevention: Use the dedicated AddReply API for nested comments and do not manipulate reply nodes directly.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Periodically verify that the Aspose.Words version used in the repository matches the version documented in the rules; update the strict API bans and canonical patterns when new comment features are added.
