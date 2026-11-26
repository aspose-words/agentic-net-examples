# Mail Merge Examples for Aspose.Words for .NET

Aspose.Words MailMerge enables programmatic creation of merge templates, insertion of MERGEFIELD tags, and execution of simple or region-based merges. The verified examples demonstrate building templates with DocumentBuilder, handling image fields via callbacks, and generating DOCX or PDF outputs.

## Snapshot

- Category: Mail Merge
- Slug: mail-merge
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261124_190740_1c00f5
- Simple Mail Merge examples: 17
- Region Mail Merge examples: 7
- Image Mail Merge examples: 4
- Input Bootstrap examples: 1
- Table Mail Merge examples: 1

## Category rules that shaped these examples

- Do not assume external template files exist; always generate the template programmatically.
- Do not reference nonexistent APIs such as MailMerge.ImageFieldMerging or MailMerge.GetRegions.
- Do not leave merge fields unresolved unless the task explicitly requires it.
- Use DocumentBuilder to construct the template and insert MERGEFIELD tags with builder.InsertField("MERGEFIELD FieldName").
- For simple merges call MailMerge.Execute; for region merges call MailMerge.ExecuteWithRegions after inserting TableStart and TableEnd fields.
- Implement IFieldMergingCallback to handle both text and image fields, supplying image data via ImageFieldMergingArgs.ImageStream.
- Validate that all field name arrays, DataTable columns, and DataSet tables are non-null and contain rows before invoking mail merge.
- Check that ImageFieldMergingArgs.ImageStream is assigned before the merge to avoid null dereference.
- Confirm that region start and end fields are present and correctly paired before executing region merges.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.11.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0

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
dotnet add package Aspose.Words --version 26.11.0

# PowerShell example
Copy-Item ..\mail-merge\create-a-mail-merge-template-programmatically-using-documentbuilder-and-add-static-header.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `create-a-mail-merge-template-programmatically-using-documentbuilder-and-add-static-header.cs` | Create a mail merge template programmatically using DocumentBuilder and add static header text. | simple-mail-merge | docx | mcp |
| 2 | `insert-merge-fields-for-customer-name-and-address-into-the-template-with-documentbuilder.cs` | Insert merge fields for customer name and address into the template with DocumentBuilder. | simple-mail-merge | docx | mcp |
| 3 | `define-a-mail-merge-region-for-order-items-by-inserting-start-and-end-merge-fields.cs` | Define a mail merge region for order items by inserting start and end merge fields. | region-mail-merge | docx | llm |
| 4 | `insert-a-table-placeholder-and-define-a-mail-merge-region-for-table-rows-using-documentbui.cs` | Insert a table placeholder and define a mail merge region for table rows using DocumentBuilder. | region-mail-merge | docx | llm |
| 5 | `handle-the-imagefieldmerging-event-to-customize-image-insertion-using-imagefieldmergingarg.cs` | Handle the ImageFieldMerging event to customize image insertion using ImageFieldMergingArgs. | image-mail-merge | docx | mcp |
| 6 | `insert-a-company-logo-into-each-merged-document-by-handling-imagefieldmerging-with-a-stati.cs` | Insert a company logo into each merged document by handling ImageFieldMerging with a static image path. | image-mail-merge | docx | llm |
| 7 | `apply-conditional-logic-in-imagefieldmerging-to-select-different-images-based-on-field-nam.cs` | Apply conditional logic in ImageFieldMerging to select different images based on field name. | image-mail-merge | docx | existing_repo |
| 8 | `set-the-text-property-of-a-merge-field-to-apply-bold-formatting-to-inserted-names.cs` | Set the Text property of a merge field to apply bold formatting to inserted names. | simple-mail-merge | docx | llm |
| 9 | `set-the-text-property-to-insert-formatted-dates-using-a-specific-culture-format-in-merge-f.cs` | Set the Text property to insert formatted dates using a specific culture format in merge fields. | simple-mail-merge | docx | mcp |
| 10 | `load-xml-data-into-a-dataset-using-the-readxml-method-for-mail-merge-source.cs` | Load XML data into a DataSet using the ReadXml method for mail merge source. | simple-mail-merge | xml | llm |
| 11 | `execute-a-simple-mail-merge-with-a-single-data-object-and-save-the-result-as-docx.cs` | Execute a simple mail merge with a single data object and save the result as DOCX. | simple-mail-merge | docx, doc | mcp |
| 12 | `execute-a-simple-mail-merge-for-multiple-records-and-generate-separate-pdf-files-for-each.cs` | Execute a simple mail merge for multiple records and generate separate PDF files for each record. | simple-mail-merge | pdf | mcp |
| 13 | `use-mailmerge-execute-with-a-collection-of-objects-to-create-a-batch-of-merged-documents.cs` | Use MailMerge.Execute with a collection of objects to create a batch of merged documents. | input-bootstrap | docx | mcp |
| 14 | `execute-a-mail-merge-with-regions-to-repeat-a-product-list-for-each-order-record.cs` | Execute a mail merge with regions to repeat a product list for each order record. | region-mail-merge | docx | llm |
| 15 | `use-mailmerge-executewithregions-to-merge-data-into-multiple-nested-regions-within-the-tem.cs` | Use MailMerge.ExecuteWithRegions to merge data into multiple nested regions within the template. | region-mail-merge | docx | mcp |
| 16 | `clone-the-template-document-after-each-merge-to-produce-independent-output-files.cs` | Clone the template document after each merge to produce independent output files. | simple-mail-merge | docx | mcp |
| 17 | `retrieve-mail-merge-region-metadata-using-mailmergeregioninfo-to-verify-start-and-end-posi.cs` | Retrieve mail merge region metadata using MailMergeRegionInfo to verify start and end positions. | region-mail-merge | docx | mcp |
| 18 | `save-the-merged-document-as-pdf-using-document-save-with-saveformat-pdf-after-mail-merge.cs` | Save the merged document as PDF using Document.Save with SaveFormat.Pdf after mail merge. | simple-mail-merge | pdf | mcp |
| 19 | `save-the-merged-document-as-docx-using-document-save-with-saveformat-docx-after-mail-merge.cs` | Save the merged document as DOCX using Document.Save with SaveFormat.Docx after mail merge. | simple-mail-merge | docx, doc | mcp |
| 20 | `perform-mail-merge-using-xml-data-source-loaded-into-a-dataset-and-generate-docx-output.cs` | Perform mail merge using XML data source loaded into a DataSet and generate DOCX output. | simple-mail-merge | docx, xml | mcp |
| 21 | `add-static-footer-text-to-the-template-using-documentbuilder-before-executing-mail-merge.cs` | Add static footer text to the template using DocumentBuilder before executing mail merge. | simple-mail-merge | docx | mcp |
| 22 | `insert-a-page-break-field-before-each-region-to-start-new-pages-for-each-repeat.cs` | Insert a PAGE_BREAK field before each region to start new pages for each repeat. | region-mail-merge | docx | mcp |
| 23 | `handle-missingfieldevent-to-implement-error-handling-for-absent-merge-fields-before-execut.cs` | Handle MissingFieldEvent to implement error handling for absent merge fields before execution. | simple-mail-merge | docx | mcp |
| 24 | `adjust-image-size-during-insertion-by-modifying-imagescale-property-in-imagefieldmergingar.cs` | Adjust image size during insertion by modifying ImageScale property in ImageFieldMergingArgs. | image-mail-merge | docx | existing_repo |
| 25 | `use-mailmergeregioninfo-to-obtain-region-start-and-end-positions-for-validation-purposes.cs` | Use MailMergeRegionInfo to obtain region start and end positions for validation purposes. | region-mail-merge | docx | existing_repo |
| 26 | `perform-mail-merge-with-xml-data-source-by-loading-xml-schema-and-data-into-a-dataset.cs` | Perform mail merge with XML data source by loading XML schema and data into a DataSet. | simple-mail-merge | xml | mcp |
| 27 | `execute-simple-mail-merge-for-a-collection-of-records-and-clone-template-to-create-separat.cs` | Execute simple mail merge for a collection of records and clone template to create separate documents. | simple-mail-merge | docx | mcp |
| 28 | `customize-text-insertion-by-setting-the-text-property-with-formatted-strings-for-each-merg.cs` | Customize text insertion by setting the Text property with formatted strings for each merge field. | simple-mail-merge | docx | mcp |
| 29 | `generate-multiple-merged-documents-by-cloning-the-template-after-each-mail-merge-operation.cs` | Generate multiple merged documents by cloning the template after each mail merge operation. | simple-mail-merge | docx | mcp |
| 30 | `use-documentbuilder-to-add-a-static-table-of-contents-that-updates-after-mail-merge-execut.cs` | Use DocumentBuilder to add a static table of contents that updates after mail merge execution. | table-mail-merge | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Incorrect region field placement

- Symptom: Mail merge throws InvalidOperationException indicating missing item names or badly formed region.
- Root cause: TableStart and TableEnd fields are not inserted in the same section, row, or cell, or the region name does not match the DataTable name.
- Preferred fix: Insert TableStart and TableEnd fields with identical region names inside the same table structure before executing ExecuteWithRegions.
- Prevention: Always build the region template first, then verify the field codes using doc.Range.Fields to ensure correct pairing.

### Use of obsolete or missing API members

- Symptom: Compilation fails with CS1061 or CS0246 for members like ImageFieldMerging, GetRegions, or MailMergeCleanupOptions.
- Root cause: The code references APIs that were renamed or removed in the current Aspose.Words version.
- Preferred fix: Replace obsolete calls with the supported properties: set doc.MailMerge.FieldMergingCallback, use MailMergeCleanupOptions enum, and access region metadata via doc.MailMerge.GetRegions() when available.
- Prevention: Consult the Aspose.Words 26.11.0 API reference and avoid hard-coding member names from older versions.

### Missing System.Drawing dependencies

- Symptom: Compilation errors CS1069 for Bitmap, Graphics, or ImageFormat types.
- Root cause: The sample attempts to use System.Drawing types, which are not included in the default .NET 9.0 runtime.
- Preferred fix: Provide image data as a byte array or stream (e.g., MemoryStream) and avoid System.Drawing classes entirely.
- Prevention: Use base64-encoded PNG data or file streams for images and rely on ImageFieldMergingArgs.ImageStream.

### Template-data field mismatch

- Symptom: Unresolved MERGEFIELD tags remain in the output or runtime throws InvalidOperationException for missing values.
- Root cause: Field names in the template do not exactly match the keys or column names of the supplied data source.
- Preferred fix: Ensure that every MERGEFIELD tag name matches a column in the DataTable or an entry in the field-name array, respecting case sensitivity.
- Prevention: Generate the field-name list programmatically from the data source and compare it to the template before merging.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Update the code whenever a new Aspose.Words version changes API signatures, and run the verification suite after each upgrade.
