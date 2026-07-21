---
name: comments
description: Verified C# examples for comments scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Comments

## Purpose

This folder is a live, curated example set for Comments scenarios. Treat every `.cs` file as a standalone console application. The goal is correct, warning-free examples that use documented Aspose APIs and match the original task intent.

## Non-negotiable conventions

- Use `Aspose.Words.Comment class` for comment APIs directly.
- Enumerate comments with `doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>()`.
- Do not use invented `Document.Comments` or `Aspose.Words.Comments` APIs.
- Do not use undocumented save-option flags such as `HtmlSaveOptions.ExportComments`.
- Prefer simple, verifiable workflows over speculative markup tricks.
- Create realistic local sample inputs whenever the task mentions streams, files, folders, DOC/DOCX inputs, XML, JSON, or database-like sources.
- Guard maybe-null values to avoid nullable-reference warnings such as `CS8600`, `CS8602`, and `CS8604`.

## Recommended workflow selection

- Native Comment Api workflow: 15 examples
- Export Report workflow: 7 examples
- Rendered Output workflow: 5 examples
- Stream Batch IO workflow: 3 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. API usage must be supported by the configured package versions.
3. Exported outputs must actually be written by the example.
4. Validation scenarios must inspect only the behavior requested by the task.
5. Examples that depend on files, folders, streams, images, or data should bootstrap those inputs locally during the example run.

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
  - Selected engine: mcp
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
  - Selected engine: existing_repo
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
  - Selected engine: existing_repo
- `synchronize-comment-positions-after-document-sections-are-reordered-to-maintain-accurate-c.cs`
  - Task: Synchronize comment positions after document sections are reordered to maintain accurate comment anchoring.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
- `use-comment-collection-events-to-trigger-custom-logging-whenever-a-comment-is-added-or-rem.cs`
  - Task: Use comment collection events to trigger custom logging whenever a comment is added or removed.
  - Workflow: native-comment-api
  - Outputs: docx
  - Selected engine: mcp
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

- **Wrong Node.Text usage**
  - Symptom: error CS1061: 'Node' does not contain a definition for 'Text'
  - Preferred fix: Use Comment.GetText() for whole-comment content, or cast to a concrete node type such as Run before reading .Text.

- **Invented comment collection APIs**
  - Symptom: Compile failures caused by Document.Comments or Aspose.Words.Comments usage.
  - Preferred fix: Enumerate comments with doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>().

- **Unsupported comment-reference or save-option APIs**
  - Symptom: Compile failures caused by CommentReference or undocumented save-option flags.
  - Preferred fix: Prefer simpler comment workflows and documented Aspose.Words save options only.

- **Unsafe live-collection mutation**
  - Symptom: Unstable behavior while deleting comments during forward iteration over a live collection.
  - Preferred fix: Create a ToList copy first, then update or remove matching comments.

- **Nullable-reference warnings**
  - Symptom: CS8600, CS8602, or CS8604 around maybe-null values such as CurrentParagraph or FirstOrDefault results.
  - Preferred fix: Use nullable locals where needed and guard maybe-null values before dereference or assignment.

## Build and run contract

- Target framework: `net8.0`
- Package: `Aspose.Words` `26.6.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.6.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\comments\<example-file>.cs .\Program.cs
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
