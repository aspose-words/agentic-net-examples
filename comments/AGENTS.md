---
name: comments
description: Verified C# examples for comments scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Comments

## Purpose

Provide an AI coding agent with clear, safe patterns for working with Aspose.Words comments, ensuring compilable code, correct comment hierarchy handling, and robust export or transformation workflows.

## Non-negotiable conventions

- Do not reference a non-existent Document.Comments collection or an invented Aspose.Words.Comments namespace.
- Do not call DocumentBuilder.StartComment, EndComment, InsertComment, AddComment, or any similar shortcut that is not part of the official API.
- Do not use CommentReference, HtmlSaveOptions.ExportComments, or any undocumented save-option flags.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Native Comment API workflow**: 15 examples
- **Export Report workflow**: 7 examples
- **Rendered Output workflow**: 5 examples
- **Stream and batch I/O workflow**: 3 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Ensure the code compiles against Aspose.Words 26.10.0 and runs without NotSupportedException or null-reference failures.

## File-to-task reference

- `load-a-docx-file-enumerate-all-comments-and-print-each-author-and-text-to-console.cs`
  - Task: Load a DOCX file, enumerate all comments, and print each author and text to console.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `add-a-new-comment-to-a-specific-paragraph-in-a-word-document-and-save-as-docx.cs`
  - Task: Add a new comment to a specific paragraph in a Word document and save as DOCX.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `update-the-text-of-an-existing-comment-identified-by-its-index-while-preserving-original-f.cs`
  - Task: Update the text of an existing comment identified by its index while preserving original formatting.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `delete-all-comments-authored-by-a-particular-user-from-the-loaded-document-before-exportin.cs`
  - Task: Delete all comments authored by a particular user from the loaded document before exporting.
  - Workflow: export-report
  - Outputs: docx
  - Selected engine: mcp
- `set-custom-author-name-and-initials-for-programmatically-added-comments-in-a-document.cs`
  - Task: Set custom author name and initials for programmatically added comments in a document.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `create-a-reply-to-an-existing-comment-and-ensure-the-reply-appears-nested-under-the-origin.cs`
  - Task: Create a reply to an existing comment and ensure the reply appears nested under the original.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: existing_repo
- `iterate-through-the-comment-collection-and-remove-comments-older-than-a-specified-date-thr.cs`
  - Task: Iterate through the comment collection and remove comments older than a specified date threshold.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `filter-comments-by-author-and-export-only-those-comments-to-a-separate-word-document-for-r.cs`
  - Task: Filter comments by author and export only those comments to a separate Word document for review.
  - Workflow: export-report
  - Outputs: docx
  - Selected engine: mcp
- `export-all-comments-from-a-docx-file-to-a-csv-file-with-author-date-and-text-columns.cs`
  - Task: Export all comments from a DOCX file to a CSV file with author, date, and text columns.
  - Workflow: export-report
  - Outputs: docx, csv
  - Selected engine: mcp
- `import-comments-from-an-exported-xml-file-and-attach-them-to-appropriate-locations-in-a-ne.cs`
  - Task: Import comments from an exported XML file and attach them to appropriate locations in a new document.
  - Workflow: export-report
  - Outputs: docx, xml
  - Selected engine: mcp
- `extract-comment-metadata-author-date-and-text-and-write-it-to-a-json-file.cs`
  - Task: Extract comment metadata-author, date, and text-and write it to a JSON file.
  - Workflow: export-report
  - Outputs: docx, json
  - Selected engine: mcp
- `search-comments-containing-a-specific-keyword-and-highlight-the-corresponding-text-range-i.cs`
  - Task: Search comments containing a specific keyword and highlight the corresponding text range in the document.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `load-multiple-word-documents-from-a-folder-aggregate-their-comments-and-generate-a-summary.cs`
  - Task: Load multiple Word documents from a folder, aggregate their comments, and generate a summary report.
  - Workflow: export-report
  - Outputs: docx
  - Selected engine: mcp
- `generate-a-printable-report-listing-all-comments-with-page-numbers-and-associated-paragrap.cs`
  - Task: Generate a printable report listing all comments with page numbers and associated paragraph text.
  - Workflow: export-report
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-custom-style-to-all-comment-text-blocks-within-a-document-to-match-corporate-brand.cs`
  - Task: Apply a custom style to all comment text blocks within a document to match corporate branding.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `preserve-comment-formatting-such-as-bold-and-italic-text-when-converting-a-document-to-htm.cs`
  - Task: Preserve comment formatting such as bold and italic text when converting a document to HTML format.
  - Workflow: rendered-output
  - Outputs: docx, html
  - Selected engine: mcp
- `convert-a-doc-file-to-pdf-while-retaining-all-comment-annotations-visible-in-the-output.cs`
  - Task: Convert a DOC file to PDF while retaining all comment annotations visible in the output.
  - Workflow: rendered-output
  - Outputs: docx, pdf, doc
  - Selected engine: mcp
- `convert-a-document-with-comments-to-xps-format-ensuring-comments-appear-as-markup-annotati.cs`
  - Task: Convert a document with comments to XPS format, ensuring comments appear as markup annotations.
  - Workflow: rendered-output
  - Outputs: xps
  - Selected engine: mcp
- `add-a-comment-containing-a-hyperlink-to-an-external-resource-and-verify-the-link-functions.cs`
  - Task: Add a comment containing a hyperlink to an external resource and verify the link functions in PDF.
  - Workflow: rendered-output
  - Outputs: pdf
  - Selected engine: mcp
- `validate-that-comment-reference-ids-update-correctly-after-inserting-new-paragraphs-into-t.cs`
  - Task: Validate that comment reference IDs update correctly after inserting new paragraphs into the document.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `synchronize-comment-positions-after-document-sections-are-reordered-to-maintain-accurate-c.cs`
  - Task: Synchronize comment positions after document sections are reordered to maintain accurate comment anchoring.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `use-comment-collection-events-to-trigger-custom-logging-whenever-a-comment-is-added-or-rem.cs`
  - Task: Use comment collection events to trigger custom logging whenever a comment is added or removed.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: llm
- `programmatically-accept-or-reject-comments-based-on-author-name-and-generate-a-revised-doc.cs`
  - Task: Programmatically accept or reject comments based on author name and generate a revised document version.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `load-a-document-from-a-stream-add-comments-and-save-the-modified-document-back-to-a-memory.cs`
  - Task: Load a document from a stream, add comments, and save the modified document back to a memory stream.
  - Workflow: stream-batch-io
  - Outputs: docx, doc
  - Selected engine: mcp
- `create-a-batch-process-that-adds-a-standardized-disclaimer-comment-to-every-document-in-a.cs`
  - Task: Create a batch process that adds a standardized disclaimer comment to every document in a directory.
  - Workflow: stream-batch-io
  - Outputs: docx
  - Selected engine: mcp
- `load-a-document-change-comment-author-names-to-uppercase-and-save-the-updated-file.cs`
  - Task: Load a document, change comment author names to uppercase, and save the updated file.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `extract-comment-text-and-embed-it-as-footnotes-within-the-same-document-for-alternative-pr.cs`
  - Task: Extract comment text and embed it as footnotes within the same document for alternative presentation.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `implement-a-feature-that-hides-all-comments-in-the-document-view-without-removing-them-fro.cs`
  - Task: Implement a feature that hides all comments in the document view without removing them from the file.
  - Workflow: rendered-output
  - Outputs: docx, pdf
  - Selected engine: mcp
- `compare-two-versions-of-a-document-and-list-comments-that-were-added-modified-or-deleted.cs`
  - Task: Compare two versions of a document and list comments that were added, modified, or deleted.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `create-a-utility-that-reads-comment-data-from-a-database-and-inserts-corresponding-comment.cs`
  - Task: Create a utility that reads comment data from a database and inserts corresponding comments into a template document.
  - Workflow: stream-batch-io
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Use of non-existent comment APIs**
  - Symptom: Compile errors such as CS0234 or CS1061 referencing Document.Comments, Aspose.Words.Comments, DocumentBuilder.StartComment/EndComment, or HtmlSaveOptions.ExportComments.
  - Root cause: The referenced members are not defined in the Aspose.Words 26.10.0 API.
  - Preferred fix: Replace invented calls with the supported pattern doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>() and manipulate Comment objects directly.
  - Prevention: Consult the official Aspose.Words documentation and restrict code to members that appear in IntelliSense.

- **Mutating a live comment collection**
  - Symptom: Runtime NotSupportedException stating that adding nodes is not supported for deep node collections, or unstable behavior when deleting during forward iteration.
  - Root cause: NodeCollection.Add or removal is performed on the original comment collection, which Aspose.Words treats as read-only for deep hierarchies.
  - Preferred fix: Create a snapshot with .ToList() and perform all Add, Remove, or Update operations on that list.
  - Prevention: Never iterate over a live NodeCollection while modifying it; always work on a copied list.

- **Incorrect member access on base Node**
  - Symptom: Compile errors CS1061 for members like .Text, .InsertAfter, or .ParentComment on a Node instance.
  - Root cause: Node is an abstract base class; those members exist only on derived types such as Run, Paragraph, or Comment.
  - Preferred fix: Cast to the appropriate derived type or use Comment.GetText() for comment content; use Paragraph.InsertAfter only on Paragraph objects.
  - Prevention: Inspect the concrete type before accessing members and rely on type-specific APIs.

- **Nullable reference warnings**
  - Symptom: CS8600, CS8602, or CS8604 warnings about possible null dereference or assignment.
  - Root cause: Code assigns or dereferences values that may be null without null checks.
  - Preferred fix: Declare variables as nullable, add explicit null checks, and use the null-conditional operator where appropriate.
  - Prevention: Enable nullable reference types and consistently validate results of queries that can return null.

- **Improper reply handling**
  - Symptom: Compile errors for missing Comment.ParentComment property and runtime NotSupportedException when trying to add a reply via NodeCollection.Add.
  - Root cause: Replies must be added using the Comment.AddReply method; treating replies as regular nodes in a deep collection is unsupported.
  - Preferred fix: Create a reply with topComment.AddReply(author, initials, date, text) and avoid manual insertion into the comment collection.
  - Prevention: Use the dedicated AddReply API for nested comments and do not manipulate reply nodes directly.

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
Copy-Item ..\comments\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Periodically verify that the Aspose.Words version used in the repository matches the version documented in the rules; update the strict API bans and canonical patterns when new comment features are added.
