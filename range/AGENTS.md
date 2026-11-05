---
name: range
description: Verified C# examples for range scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Range

## Purpose

Provide an AI coding agent with clear guidance to generate reliable Aspose.Words snippets that manipulate document ranges without invoking unsupported members, ambiguous types, or external resources.

## Non-negotiable conventions

- Do not reference members that are read-only or absent in Aspose.Words 26.10.0.
- Never use System.Range in place of Aspose.Words.Range.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Range Workflow workflow**: 30 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Confirm that the generated code compiles against Aspose.Words 26.10.0 and that the expected output file is created.

## File-to-task reference

- `extract-plain-unformatted-text-from-a-document-using-the-range-text-property.cs`
  - Task: Extract plain unformatted text from a document using the Range.Text property.
  - Workflow: range-workflow
  - Outputs: doc
  - Selected engine: mcp
- `extract-plain-text-of-each-section-via-each-section-s-range-text-property.cs`
  - Task: Extract plain text of each section via each section's Range.Text property.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: llm
- `use-the-range-object-to-extract-plain-text-from-a-header-and-footer-for-indexing.cs`
  - Task: Use the Range object to extract plain text from a header and footer for indexing.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `generate-a-plain-text-version-of-a-document-by-extracting-each-section-s-range-text.cs`
  - Task: Generate a plain-text version of a document by extracting each section's Range.Text.
  - Workflow: range-workflow
  - Outputs: doc
  - Selected engine: llm
- `retrieve-the-count-of-bookmarks-within-a-range.cs`
  - Task: Retrieve the count of bookmarks within a range.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-the-count-of-form-fields-within-a-range.cs`
  - Task: Retrieve the count of form fields within a range.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `iterate-through-each-bookmark-in-a-range-and-output-its-name.cs`
  - Task: Iterate through each bookmark in a range and output its name.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `iterate-over-bookmarks-in-a-range-and-modify-their-names.cs`
  - Task: Iterate over bookmarks in a range and modify their names.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `add-a-new-bookmark-at-the-start-of-the-document-using-range-bookmarks.cs`
  - Task: Add a new bookmark at the start of the document using Range.Bookmarks.
  - Workflow: range-workflow
  - Outputs: doc
  - Selected engine: mcp
- `remove-a-specific-bookmark-by-locating-its-range-and-calling-remove.cs`
  - Task: Remove a specific bookmark by locating its range and calling Remove.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `clear-the-text-of-a-specific-bookmark-s-range-without-deleting-the-bookmark.cs`
  - Task: Clear the text of a specific bookmark's range without deleting the bookmark.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `replace-text-within-a-range-by-assigning-a-new-string-to-the-range-text-property.cs`
  - Task: Replace text within a range by assigning a new string to the Range.Text property.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `search-for-a-specific-phrase-within-a-range-and-replace-it-with-another-string.cs`
  - Task: Search for a specific phrase within a range and replace it with another string.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `perform-a-case-insensitive-search-within-a-range-and-collect-matching-paragraph-indices.cs`
  - Task: Perform a case-insensitive search within a range and collect matching paragraph indices.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `insert-new-text-at-the-beginning-of-a-range.cs`
  - Task: Insert new text at the beginning of a range.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: llm
- `append-text-to-the-end-of-a-range.cs`
  - Task: Append text to the end of a range.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `copy-paragraph-s-range-content-into-a-string-variable-for-further-processing.cs`
  - Task: Copy paragraph's range content into a string variable for further processing.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `delete-all-characters-in-a-document-s-body-by-calling-delete-on-doc-range.cs`
  - Task: Delete all characters in a document's body by calling Delete on doc.Range.
  - Workflow: range-workflow
  - Outputs: doc
  - Selected engine: mcp
- `save-the-document-after-removing-all-content-from-its-range-to-create-an-empty-template.cs`
  - Task: Save the document after removing all content from its range to create an empty template.
  - Workflow: range-workflow
  - Outputs: doc
  - Selected engine: mcp
- `implement-a-batch-process-that-clears-the-content-of-multiple-documents-using-doc-range-de.cs`
  - Task: Implement a batch process that clears the content of multiple documents using doc.Range.Delete.
  - Workflow: range-workflow
  - Outputs: doc
  - Selected engine: mcp
- `iterate-over-form-fields-in-a-range-and-list-their-names-and-types.cs`
  - Task: Iterate over form fields in a range and list their names and types.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-checkbox-form-field-inside-a-specific-range-and-set-its-default-state.cs`
  - Task: Create a checkbox form field inside a specific range and set its default state.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `update-the-value-of-a-text-input-form-field-located-in-a-given-range.cs`
  - Task: Update the value of a text input form field located in a given range.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `delete-all-form-fields-within-a-document-by-iterating-over-range-formfields-and-calling-re.cs`
  - Task: Delete all form fields within a document by iterating over Range.FormFields and calling Remove.
  - Workflow: range-workflow
  - Outputs: doc
  - Selected engine: mcp
- `validate-that-a-range-contains-no-form-fields-before-performing-a-text-replacement-operati.cs`
  - Task: Validate that a range contains no form fields before performing a text replacement operation.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `generate-a-summary-report-of-bookmark-names-and-their-corresponding-text-extracted-from-ra.cs`
  - Task: Generate a summary report of bookmark names and their corresponding text extracted from ranges.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `replace-placeholder-text-inside-a-range-with-dynamic-data-retrieved-from-a-database.cs`
  - Task: Replace placeholder text inside a range with dynamic data retrieved from a database.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `implement-a-script-that-removes-all-bookmarks-and-form-fields-from-a-document-range-before.cs`
  - Task: Implement a script that removes all bookmarks and form fields from a document range before publishing.
  - Workflow: range-workflow
  - Outputs: doc
  - Selected engine: llm
- `log-the-names-of-all-bookmarks-found-in-a-range-for-debugging.cs`
  - Task: Log the names of all bookmarks found in a range for debugging.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `export-the-extracted-plain-text-from-a-range-to-a-txt-file-while-preserving-line-breaks.cs`
  - Task: Export the extracted plain text from a range to a .txt file while preserving line breaks.
  - Workflow: range-workflow
  - Outputs: txt
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Read-only member assignment**
  - Symptom: Compilation fails with CS0200 because Section.Body or Range.Text is assigned a value.
  - Root cause: Attempting to set properties that are defined as read-only in the Aspose.Words API.
  - Preferred fix: Use the appropriate insertion methods (e.g., DocumentBuilder, InsertBefore) instead of assigning to read-only members.
  - Prevention: Consult the API reference to verify mutability before writing to a property.

- **Ambiguous Range reference**
  - Symptom: Compilation fails with CS0104 due to a conflict between System.Range and Aspose.Words.Range.
  - Root cause: The code imports both namespaces without aliasing, causing the compiler to treat Range as ambiguous.
  - Preferred fix: Alias Aspose.Words.Range (e.g., using AsposeRange = Aspose.Words.Range) or fully qualify the type.
  - Prevention: Avoid importing System.Range or always qualify the Aspose.Words.Range type.

- **Non-existent member usage**
  - Symptom: CS1061 errors appear for methods such as BookmarkCollection.Add, Range.InsertBefore, Range.Search, or FormField.StartNode.
  - Root cause: The code calls members that do not exist in the selected Aspose.Words version.
  - Preferred fix: Replace the call with the documented alternative (e.g., DocumentBuilder.StartBookmark, parentNode.InsertBefore, doc.Range.Replace).
  - Prevention: Reference the official Aspose.Words documentation for the exact method signatures.

- **Missing overload arguments**
  - Symptom: CS7036 errors indicate required parameters (e.g., maxLength for InsertTextInput) are omitted.
  - Root cause: The code supplies an overload that does not match any available signature.
  - Preferred fix: Provide all required arguments as defined in the API, using named parameters when helpful.
  - Prevention: Review method overloads in the API reference before coding.

- **Assumed external resources**
  - Symptom: Runtime NullReferenceException occurs because a source file or document node was never created.
  - Root cause: The example expects an existing file or document structure that was not bootstrapped locally.
  - Preferred fix: Create deterministic sample documents, sections, and files within the code before processing them.
  - Prevention: Always include local input creation steps as part of the workflow.

- **Over-broad validation**
  - Symptom: Runtime checks examine unrelated document internals, causing failures unrelated to the task.
  - Root cause: Validation logic is too generic and touches nodes that are not part of the requested operation.
  - Preferred fix: Validate only the existence of the output file and the specific range result required by the task.
  - Prevention: Scope assertions to the exact behavior being demonstrated.

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
Copy-Item ..\range\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Periodically verify that all referenced Aspose.Words members remain unchanged in future package releases.
