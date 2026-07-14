# Mail Merge Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Mail Merge category. Each file is a standalone console example selected from the verified 26.5.0 run.

## Snapshot

- Category: Mail Merge
- Slug: mail-merge
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20260619_131835_59df5f
- Image Mail Merge examples: 4
- Input Bootstrap examples: 1
- Region Mail Merge examples: 7
- Simple Mail Merge examples: 17
- Table Mail Merge examples: 1

## Category rules that shaped these examples

- Do not assume template files already exist.
- Do not mismatch merge field names and data field names.
- Do not invent unsupported mail-merge APIs.
- Do not leave merge fields unresolved unless explicitly required by the task.
- Create the template locally with DocumentBuilder.
- Insert merge fields explicitly with builder.InsertField("MERGEFIELD FieldName").
- Use MailMerge.Execute for simple merges and ExecuteWithRegions for region-based merges.
- Use MailMergeCleanupOptions when the task requires a clean merged result.
- Validate arrays, tables, and lookup results before executing mail merge.
- Avoid maybe-null dereference when reading field names, table values, or merge results.

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
Copy-Item ..\mail-merge\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `mail-merge/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

# PowerShell example
Copy-Item ..\mail-merge\create-a-mail-merge-template-programmatically-using-documentbuilder-and-add-static-header.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `create-a-mail-merge-template-programmatically-using-documentbuilder-and-add-static-header.cs` | Create a mail merge template programmatically using DocumentBuilder and add static header text. | Simple Mail Merge | docx | mcp |
| 2 | `insert-merge-fields-for-customer-name-and-address-into-the-template-with-documentbuilder.cs` | Insert merge fields for customer name and address into the template with DocumentBuilder. | Simple Mail Merge | docx | mcp |
| 3 | `define-a-mail-merge-region-for-order-items-by-inserting-start-and-end-merge-fields.cs` | Define a mail merge region for order items by inserting start and end merge fields. | Region Mail Merge | docx | mcp |
| 4 | `insert-a-table-placeholder-and-define-a-mail-merge-region-for-table-rows-using-documentbui.cs` | Insert a table placeholder and define a mail merge region for table rows using DocumentBuilder. | Region Mail Merge | docx | mcp |
| 5 | `handle-the-imagefieldmerging-event-to-customize-image-insertion-using-imagefieldmergingarg.cs` | Handle the ImageFieldMerging event to customize image insertion using ImageFieldMergingArgs. | Image Mail Merge | docx | mcp |
| 6 | `insert-a-company-logo-into-each-merged-document-by-handling-imagefieldmerging-with-a-stati.cs` | Insert a company logo into each merged document by handling ImageFieldMerging with a static image path. | Image Mail Merge | docx | mcp |
| 7 | `apply-conditional-logic-in-imagefieldmerging-to-select-different-images-based-on-field-nam.cs` | Apply conditional logic in ImageFieldMerging to select different images based on field name. | Image Mail Merge | docx | llm |
| 8 | `set-the-text-property-of-a-merge-field-to-apply-bold-formatting-to-inserted-names.cs` | Set the Text property of a merge field to apply bold formatting to inserted names. | Simple Mail Merge | docx | mcp |
| 9 | `set-the-text-property-to-insert-formatted-dates-using-a-specific-culture-format-in-merge-f.cs` | Set the Text property to insert formatted dates using a specific culture format in merge fields. | Simple Mail Merge | docx | mcp |
| 10 | `load-xml-data-into-a-dataset-using-the-readxml-method-for-mail-merge-source.cs` | Load XML data into a DataSet using the ReadXml method for mail merge source. | Simple Mail Merge | xml | mcp |
| 11 | `execute-a-simple-mail-merge-with-a-single-data-object-and-save-the-result-as-docx.cs` | Execute a simple mail merge with a single data object and save the result as DOCX. | Simple Mail Merge | docx, doc | mcp |
| 12 | `execute-a-simple-mail-merge-for-multiple-records-and-generate-separate-pdf-files-for-each.cs` | Execute a simple mail merge for multiple records and generate separate PDF files for each record. | Simple Mail Merge | pdf | mcp |
| 13 | `use-mailmerge-execute-with-a-collection-of-objects-to-create-a-batch-of-merged-documents.cs` | Use MailMerge.Execute with a collection of objects to create a batch of merged documents. | Input Bootstrap | docx | mcp |
| 14 | `execute-a-mail-merge-with-regions-to-repeat-a-product-list-for-each-order-record.cs` | Execute a mail merge with regions to repeat a product list for each order record. | Region Mail Merge | docx | mcp |
| 15 | `use-mailmerge-executewithregions-to-merge-data-into-multiple-nested-regions-within-the-tem.cs` | Use MailMerge.ExecuteWithRegions to merge data into multiple nested regions within the template. | Region Mail Merge | docx | mcp |
| 16 | `clone-the-template-document-after-each-merge-to-produce-independent-output-files.cs` | Clone the template document after each merge to produce independent output files. | Simple Mail Merge | docx | mcp |
| 17 | `retrieve-mail-merge-region-metadata-using-mailmergeregioninfo-to-verify-start-and-end-posi.cs` | Retrieve mail merge region metadata using MailMergeRegionInfo to verify start and end positions. | Region Mail Merge | docx | mcp |
| 18 | `save-the-merged-document-as-pdf-using-document-save-with-saveformat-pdf-after-mail-merge.cs` | Save the merged document as PDF using Document.Save with SaveFormat.Pdf after mail merge. | Simple Mail Merge | pdf | mcp |
| 19 | `save-the-merged-document-as-docx-using-document-save-with-saveformat-docx-after-mail-merge.cs` | Save the merged document as DOCX using Document.Save with SaveFormat.Docx after mail merge. | Simple Mail Merge | docx, doc | mcp |
| 20 | `perform-mail-merge-using-xml-data-source-loaded-into-a-dataset-and-generate-docx-output.cs` | Perform mail merge using XML data source loaded into a DataSet and generate DOCX output. | Simple Mail Merge | docx, xml | mcp |
| 21 | `add-static-footer-text-to-the-template-using-documentbuilder-before-executing-mail-merge.cs` | Add static footer text to the template using DocumentBuilder before executing mail merge. | Simple Mail Merge | docx | mcp |
| 22 | `insert-a-page-break-field-before-each-region-to-start-new-pages-for-each-repeat.cs` | Insert a PAGE_BREAK field before each region to start new pages for each repeat. | Region Mail Merge | docx | mcp |
| 23 | `handle-missingfieldevent-to-implement-error-handling-for-absent-merge-fields-before-execut.cs` | Handle MissingFieldEvent to implement error handling for absent merge fields before execution. | Simple Mail Merge | docx | mcp |
| 24 | `adjust-image-size-during-insertion-by-modifying-imagescale-property-in-imagefieldmergingar.cs` | Adjust image size during insertion by modifying ImageScale property in ImageFieldMergingArgs. | Image Mail Merge | docx | mcp |
| 25 | `use-mailmergeregioninfo-to-obtain-region-start-and-end-positions-for-validation-purposes.cs` | Use MailMergeRegionInfo to obtain region start and end positions for validation purposes. | Region Mail Merge | docx | mcp |
| 26 | `perform-mail-merge-with-xml-data-source-by-loading-xml-schema-and-data-into-a-dataset.cs` | Perform mail merge with XML data source by loading XML schema and data into a DataSet. | Simple Mail Merge | xml | mcp |
| 27 | `execute-simple-mail-merge-for-a-collection-of-records-and-clone-template-to-create-separat.cs` | Execute simple mail merge for a collection of records and clone template to create separate documents. | Simple Mail Merge | docx | mcp |
| 28 | `customize-text-insertion-by-setting-the-text-property-with-formatted-strings-for-each-merg.cs` | Customize text insertion by setting the Text property with formatted strings for each merge field. | Simple Mail Merge | docx | mcp |
| 29 | `generate-multiple-merged-documents-by-cloning-the-template-after-each-mail-merge-operation.cs` | Generate multiple merged documents by cloning the template after each mail merge operation. | Simple Mail Merge | docx | mcp |
| 30 | `use-documentbuilder-to-add-a-static-table-of-contents-that-updates-after-mail-merge-execut.cs` | Use DocumentBuilder to add a static table of contents that updates after mail merge execution. | Table Mail Merge | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Template and data mismatch

- Symptom: Merge fields remain unresolved because field names do not match provided data keys or columns.
- Fix: Ensure merge field names exactly match the field-name array or DataTable column names.

### Missing local template bootstrap

- Symptom: Examples assume an existing mail merge template document already exists.
- Fix: Create the mail merge template locally using DocumentBuilder before execution.

### Wrong API choice for regions

- Symptom: Region merges are attempted with simple Execute instead of ExecuteWithRegions.
- Fix: Use TableStart/TableEnd fields and ExecuteWithRegions with a matching DataTable.

### Cleanup options omitted

- Symptom: Unused merge fields or regions remain in the output unexpectedly.
- Fix: Apply MailMergeCleanupOptions when the task requires unused fields or empty regions to be removed.

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
