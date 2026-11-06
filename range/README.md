# Range Examples for Aspose.Words for .NET

The Range category focuses on using Aspose.Words.Range and node-scoped Range objects for text extraction, search-replace, bookmark, and form-field operations. Verified examples demonstrate correct document bootstrapping, section-level text retrieval, form-field counting, bookmark enumeration, and safe insertion of text.

## Snapshot

- Category: Range
- Slug: range
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261102_031946_08fb5b
- Range Workflow examples: 30

## Category rules that shaped these examples

- Do not reference members that are read-only or absent in Aspose.Words 26.10.0.
- Never use System.Range in place of Aspose.Words.Range.
- Use DocumentBuilder for inserting bookmarks, form fields, or new paragraphs.
- Use Range.Text for read-only extraction; use Range.Replace with FindReplaceOptions for modifications.
- Iterate over range.Bookmarks or range.FormFields to enumerate names and types.
- Always ensure a node (e.g., first paragraph) exists before accessing its ParentNode or inserting before it.
- Create required sections, bodies, and paragraphs locally before processing.

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
dotnet add package Aspose.Words --version 26.10.0

# PowerShell example
Copy-Item ..\range\extract-plain-unformatted-text-from-a-document-using-the-range-text-property.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `extract-plain-unformatted-text-from-a-document-using-the-range-text-property.cs` | Extract plain unformatted text from a document using the Range.Text property. | range-workflow | doc | mcp |
| 2 | `extract-plain-text-of-each-section-via-each-section-s-range-text-property.cs` | Extract plain text of each section via each section's Range.Text property. | range-workflow | docx | llm |
| 3 | `use-the-range-object-to-extract-plain-text-from-a-header-and-footer-for-indexing.cs` | Use the Range object to extract plain text from a header and footer for indexing. | range-workflow | docx | mcp |
| 4 | `generate-a-plain-text-version-of-a-document-by-extracting-each-section-s-range-text.cs` | Generate a plain-text version of a document by extracting each section's Range.Text. | range-workflow | doc | llm |
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
| 15 | `insert-new-text-at-the-beginning-of-a-range.cs` | Insert new text at the beginning of a range. | range-workflow | docx | llm |
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
| 28 | `implement-a-script-that-removes-all-bookmarks-and-form-fields-from-a-document-range-before.cs` | Implement a script that removes all bookmarks and form fields from a document range before publishing. | range-workflow | doc | llm |
| 29 | `log-the-names-of-all-bookmarks-found-in-a-range-for-debugging.cs` | Log the names of all bookmarks found in a range for debugging. | range-workflow | docx | mcp |
| 30 | `export-the-extracted-plain-text-from-a-range-to-a-txt-file-while-preserving-line-breaks.cs` | Export the extracted plain text from a range to a .txt file while preserving line breaks. | range-workflow | txt | mcp |

## Common failure patterns seen during generation and how they were corrected

### Read-only member assignment

- Symptom: Compilation fails with CS0200 because Section.Body or Range.Text is assigned a value.
- Root cause: Attempting to set properties that are defined as read-only in the Aspose.Words API.
- Preferred fix: Use the appropriate insertion methods (e.g., DocumentBuilder, InsertBefore) instead of assigning to read-only members.
- Prevention: Consult the API reference to verify mutability before writing to a property.

### Ambiguous Range reference

- Symptom: Compilation fails with CS0104 due to a conflict between System.Range and Aspose.Words.Range.
- Root cause: The code imports both namespaces without aliasing, causing the compiler to treat Range as ambiguous.
- Preferred fix: Alias Aspose.Words.Range (e.g., using AsposeRange = Aspose.Words.Range) or fully qualify the type.
- Prevention: Avoid importing System.Range or always qualify the Aspose.Words.Range type.

### Non-existent member usage

- Symptom: CS1061 errors appear for methods such as BookmarkCollection.Add, Range.InsertBefore, Range.Search, or FormField.StartNode.
- Root cause: The code calls members that do not exist in the selected Aspose.Words version.
- Preferred fix: Replace the call with the documented alternative (e.g., DocumentBuilder.StartBookmark, parentNode.InsertBefore, doc.Range.Replace).
- Prevention: Reference the official Aspose.Words documentation for the exact method signatures.

### Missing overload arguments

- Symptom: CS7036 errors indicate required parameters (e.g., maxLength for InsertTextInput) are omitted.
- Root cause: The code supplies an overload that does not match any available signature.
- Preferred fix: Provide all required arguments as defined in the API, using named parameters when helpful.
- Prevention: Review method overloads in the API reference before coding.

### Assumed external resources

- Symptom: Runtime NullReferenceException occurs because a source file or document node was never created.
- Root cause: The example expects an existing file or document structure that was not bootstrapped locally.
- Preferred fix: Create deterministic sample documents, sections, and files within the code before processing them.
- Prevention: Always include local input creation steps as part of the workflow.

### Over-broad validation

- Symptom: Runtime checks examine unrelated document internals, causing failures unrelated to the task.
- Root cause: Validation logic is too generic and touches nodes that are not part of the requested operation.
- Preferred fix: Validate only the existence of the output file and the specific range result required by the task.
- Prevention: Scope assertions to the exact behavior being demonstrated.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Periodically verify that all referenced Aspose.Words members remain unchanged in future package releases.
