---
name: range
description: Verified C# examples for range scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Range

## Purpose

Provide an AI coding assistant with concise, reproducible patterns for safe Range operations, ensuring that all inputs are created locally, only documented APIs are used, and output files are persisted.

## Non-negotiable conventions

- Do not reference members that are not present in Aspose.Words 26.11.0.
- Do not import or use System.Range when Aspose.Words.Range is required.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Range Workflow workflow**: 30 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Confirm that the generated code compiles against Aspose.Words 26.11.0 and that the expected output file exists after execution.

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
  - Selected engine: mcp
- `use-the-range-object-to-extract-plain-text-from-a-header-and-footer-for-indexing.cs`
  - Task: Use the Range object to extract plain text from a header and footer for indexing.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: mcp
- `generate-a-plain-text-version-of-a-document-by-extracting-each-section-s-range-text.cs`
  - Task: Generate a plain-text version of a document by extracting each section's Range.Text.
  - Workflow: range-workflow
  - Outputs: doc
  - Selected engine: mcp
- `retrieve-the-count-of-bookmarks-within-a-range.cs`
  - Task: Retrieve the count of bookmarks within a range.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: llm
- `retrieve-the-count-of-form-fields-within-a-range.cs`
  - Task: Retrieve the count of form fields within a range.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: existing_repo
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
  - Selected engine: mcp
- `append-text-to-the-end-of-a-range.cs`
  - Task: Append text to the end of a range.
  - Workflow: range-workflow
  - Outputs: docx
  - Selected engine: llm
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
  - Selected engine: llm
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
  - Selected engine: mcp
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

- **Ambiguous Range reference**
  - Symptom: The compiler reports "'Range' is an ambiguous reference between 'Aspose.Words.Range' and 'System.Range'".
  - Root cause: Both System.Range (C# 8) and Aspose.Words.Range are imported, and the code does not qualify the type.
  - Preferred fix: Fully qualify the Aspose.Words.Range type or add an alias (e.g., using AwRange = Aspose.Words.Range).
  - Prevention: Always import Aspose.Words.Range with an explicit alias when System.Range is also in scope.

- **Invocation of non-existent members**
  - Symptom: Compilation fails with CS1061 or CS1729 indicating missing methods such as HeaderFooter.IsFooter, BookmarkCollection.Add, Range.InsertBefore, or assigning to Range.Text.
  - Root cause: The code assumes APIs that were never part of the documented library version.
  - Preferred fix: Replace the call with the correct documented alternative (e.g., use HeaderFooter.HeaderFooterType to identify footers, use DocumentBuilder.StartBookmark/EndBookmark instead of BookmarkCollection.Add, use DocumentBuilder.Write or InsertAfter on a Paragraph for insertion, and use Range.Replace for text changes).
  - Prevention: Consult the official Aspose.Words API reference for the exact member signatures before coding.

- **Incorrect enum or missing namespace**
  - Symptom: CS0234 or CS0103 errors for types like TextFormFieldType, Builders, or BuildingBlocks.
  - Root cause: The code references enums or namespaces that either belong to a different version or were misspelled.
  - Preferred fix: Import Aspose.Words.Fields and use the correct enum values (e.g., TextFormFieldType.Regular) and correct namespaces (Aspose.Words.BuildingBlocks).
  - Prevention: Verify enum names and namespaces against the 26.11.0 documentation and enable IntelliSense to avoid typographical errors.

- **Missing local bootstrap inputs**
  - Symptom: Runtime NullReferenceException or compile-time errors when loading a document that has not been saved first.
  - Root cause: The example attempts to load or process a file that was never created in the sandbox.
  - Preferred fix: Create the required document, sections, or other resources programmatically before loading or processing them.
  - Prevention: Always generate a deterministic sample document at the start of the program and reference that file for subsequent operations.

- **Over-broad validation**
  - Symptom: The program checks unrelated document internals, causing failures unrelated to the task.
  - Root cause: Validation logic is not scoped to the specific operation being demonstrated.
  - Preferred fix: Limit assertions to the existence of the output file and to the direct result of the performed Range operation.
  - Prevention: Design validation steps that verify only the behavior described in the task.

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
Copy-Item ..\range\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Periodically run the verified examples against the latest Aspose.Words package to detect API deprecations and update the canonical patterns accordingly.
