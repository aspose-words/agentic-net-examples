# Range Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Range category. Each file is a standalone console example selected from the verified 26.6.0 run.

## Snapshot

- Category: Range
- Slug: range
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20260711_192617_b9179d
- Range Workflow examples: 30

## Category rules that shaped these examples

- Use native Aspose.Words APIs directly.
- Create local sample source documents when a task refers to an existing file, folder, stream, or template.
- Do not assume external files or folders already exist.
- Prefer `Document.Range` and node-scoped `Range` APIs for search, replace, extraction, bookmarks, and fields.
- Keep validation narrow and task-specific.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.6.0

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.6.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\range\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `range/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.6.0

# PowerShell example
Copy-Item ..\range\extract-plain-unformatted-text-from-a-document-using-the-range-text-property.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `extract-plain-unformatted-text-from-a-document-using-the-range-text-property.cs` | Extract plain unformatted text from a document using the Range.Text property. | range-workflow | doc | mcp |
| 2 | `extract-plain-text-of-each-section-via-each-section-s-range-text-property.cs` | Extract plain text of each section via each section's Range.Text property. | range-workflow | docx | mcp |
| 3 | `use-the-range-object-to-extract-plain-text-from-a-header-and-footer-for-indexing.cs` | Use the Range object to extract plain text from a header and footer for indexing. | range-workflow | docx | mcp |
| 4 | `generate-a-plain-text-version-of-a-document-by-extracting-each-section-s-range-text.cs` | Generate a plain-text version of a document by extracting each section's Range.Text. | range-workflow | doc | mcp |
| 5 | `retrieve-the-count-of-bookmarks-within-a-range.cs` | Retrieve the count of bookmarks within a range. | range-workflow | docx | mcp |
| 6 | `retrieve-the-count-of-form-fields-within-a-range.cs` | Retrieve the count of form fields within a range. | range-workflow | docx | mcp |
| 7 | `iterate-through-each-bookmark-in-a-range-and-output-its-name.cs` | Iterate through each bookmark in a range and output its name. | range-workflow | docx | mcp |
| 8 | `iterate-over-bookmarks-in-a-range-and-modify-their-names.cs` | Iterate over bookmarks in a range and modify their names. | range-workflow | docx | mcp |
| 9 | `add-a-new-bookmark-at-the-start-of-the-document-using-range-bookmarks.cs` | Add a new bookmark at the start of the document using Range.Bookmarks. | range-workflow | doc | mcp |
| 10 | `remove-a-specific-bookmark-by-locating-its-range-and-calling-remove.cs` | Remove a specific bookmark by locating its range and calling Remove. | range-workflow | docx | mcp |
| 11 | `clear-the-text-of-a-specific-bookmark-s-range-without-deleting-the-bookmark.cs` | Clear the text of a specific bookmark's range without deleting the bookmark. | range-workflow | docx | mcp |
| 12 | `replace-text-within-a-range-by-assigning-a-new-string-to-the-range-text-property.cs` | Replace text within a range by assigning a new string to the Range.Text property. | range-workflow | docx | mcp |
| 13 | `search-for-a-specific-phrase-within-a-range-and-replace-it-with-another-string.cs` | Search for a specific phrase within a range and replace it with another string. | range-workflow | docx | mcp |
| 14 | `perform-a-case-insensitive-search-within-a-range-and-collect-matching-paragraph-indices.cs` | Perform a case-insensitive search within a range and collect matching paragraph indices. | range-workflow | docx | mcp |
| 15 | `insert-new-text-at-the-beginning-of-a-range.cs` | Insert new text at the beginning of a range. | range-workflow | docx | mcp |
| 16 | `append-text-to-the-end-of-a-range.cs` | Append text to the end of a range. | range-workflow | docx | mcp |
| 17 | `copy-paragraph-s-range-content-into-a-string-variable-for-further-processing.cs` | Copy paragraph's range content into a string variable for further processing. | range-workflow | docx | mcp |
| 18 | `delete-all-characters-in-a-document-s-body-by-calling-delete-on-doc-range.cs` | Delete all characters in a document's body by calling Delete on doc.Range. | range-workflow | doc | mcp |
| 19 | `save-the-document-after-removing-all-content-from-its-range-to-create-an-empty-template.cs` | Save the document after removing all content from its range to create an empty template. | range-workflow | doc | mcp |
| 20 | `implement-a-batch-process-that-clears-the-content-of-multiple-documents-using-doc-range-de.cs` | Implement a batch process that clears the content of multiple documents using doc.Range.Delete. | range-workflow | doc | mcp |
| 21 | `iterate-over-form-fields-in-a-range-and-list-their-names-and-types.cs` | Iterate over form fields in a range and list their names and types. | range-workflow | docx | mcp |
| 22 | `create-a-checkbox-form-field-inside-a-specific-range-and-set-its-default-state.cs` | Create a checkbox form field inside a specific range and set its default state. | range-workflow | docx | mcp |
| 23 | `update-the-value-of-a-text-input-form-field-located-in-a-given-range.cs` | Update the value of a text input form field located in a given range. | range-workflow | docx | mcp |
| 24 | `delete-all-form-fields-within-a-document-by-iterating-over-range-formfields-and-calling-re.cs` | Delete all form fields within a document by iterating over Range.FormFields and calling Remove. | range-workflow | doc | mcp |
| 25 | `validate-that-a-range-contains-no-form-fields-before-performing-a-text-replacement-operati.cs` | Validate that a range contains no form fields before performing a text replacement operation. | range-workflow | docx | mcp |
| 26 | `generate-a-summary-report-of-bookmark-names-and-their-corresponding-text-extracted-from-ra.cs` | Generate a summary report of bookmark names and their corresponding text extracted from ranges. | range-workflow | docx | mcp |
| 27 | `replace-placeholder-text-inside-a-range-with-dynamic-data-retrieved-from-a-database.cs` | Replace placeholder text inside a range with dynamic data retrieved from a database. | range-workflow | docx | mcp |
| 28 | `implement-a-script-that-removes-all-bookmarks-and-form-fields-from-a-document-range-before.cs` | Implement a script that removes all bookmarks and form fields from a document range before publishing. | range-workflow | doc | mcp |
| 29 | `log-the-names-of-all-bookmarks-found-in-a-range-for-debugging.cs` | Log the names of all bookmarks found in a range for debugging. | range-workflow | docx | mcp |
| 30 | `export-the-extracted-plain-text-from-a-range-to-a-txt-file-while-preserving-line-breaks.cs` | Export the extracted plain text from a range to a .txt file while preserving line breaks. | range-workflow | txt | mcp |

## Common failure patterns seen during generation and how they were corrected

### Unsupported API invention

- Symptom: Generated code references members that do not exist in the selected package version.
- Fix: Replace invented members with documented Aspose.Words APIs already proven in this category.

### Missing local bootstrap inputs

- Symptom: The example assumes source files, folders, images, or data already exist.
- Fix: Create deterministic local inputs before loading, processing, or validating them.

### Over-broad validation

- Symptom: The example fails at runtime while checking unrelated document internals.
- Fix: Validate only the requested behavior and the existence of expected outputs.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.6.0 run.
- Preserve file-to-task traceability when updating this folder.
- Keep examples standalone and bootstrap local inputs inside the example whenever external sources are mentioned.
