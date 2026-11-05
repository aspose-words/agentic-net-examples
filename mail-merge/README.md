# Mail Merge Examples for Aspose.Words for .NET

Mail Merge in Aspose.Words enables programmatic creation of merge templates, insertion of MERGEFIELD tags, and execution of simple or region-based merges. The verified examples demonstrate correct template bootstrapping, image field handling via IFieldMergingCallback, and generation of separate PDF outputs for each record.

## Snapshot

- Category: Mail Merge
- Slug: mail-merge
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261102_031946_08fb5b
- Simple Mail Merge examples: 17
- Region Mail Merge examples: 7
- Image Mail Merge examples: 4
- Input Bootstrap examples: 1
- Table Mail Merge examples: 1

## Category rules that shaped these examples

- Do not reference types such as Table, FieldType, MailMergeCleanupOptions, GetRegions, or ImageScale that are not present in the referenced Aspose.Words version.
- Do not assume that a template file exists on disk; always construct the template programmatically or load a known file.
- Use DocumentBuilder to build the template and insert MERGEFIELD tags explicitly.
- Implement IFieldMergingCallback for image fields and assign it to doc.MailMerge.FieldMergingCallback.
- Choose MailMerge.Execute for flat data and MailMerge.ExecuteWithRegions for table-based regions.
- Validate that all field-name arrays, DataTable columns, and image byte arrays are non-null before calling Execute.
- Guard against missing values in the data source by providing default text in the FieldMerging callback.

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
dotnet add package Aspose.Words --version 26.10.0

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
| 4 | `insert-a-table-placeholder-and-define-a-mail-merge-region-for-table-rows-using-documentbui.cs` | Insert a table placeholder and define a mail merge region for table rows using DocumentBuilder. | region-mail-merge | docx | mcp |
| 5 | `handle-the-imagefieldmerging-event-to-customize-image-insertion-using-imagefieldmergingarg.cs` | Handle the ImageFieldMerging event to customize image insertion using ImageFieldMergingArgs. | image-mail-merge | docx | llm |
| 6 | `insert-a-company-logo-into-each-merged-document-by-handling-imagefieldmerging-with-a-stati.cs` | Insert a company logo into each merged document by handling ImageFieldMerging with a static image path. | image-mail-merge | docx | mcp |
| 7 | `apply-conditional-logic-in-imagefieldmerging-to-select-different-images-based-on-field-nam.cs` | Apply conditional logic in ImageFieldMerging to select different images based on field name. | image-mail-merge | docx | llm |
| 8 | `set-the-text-property-of-a-merge-field-to-apply-bold-formatting-to-inserted-names.cs` | Set the Text property of a merge field to apply bold formatting to inserted names. | simple-mail-merge | docx | mcp |
| 9 | `set-the-text-property-to-insert-formatted-dates-using-a-specific-culture-format-in-merge-f.cs` | Set the Text property to insert formatted dates using a specific culture format in merge fields. | simple-mail-merge | docx | mcp |
| 10 | `load-xml-data-into-a-dataset-using-the-readxml-method-for-mail-merge-source.cs` | Load XML data into a DataSet using the ReadXml method for mail merge source. | simple-mail-merge | xml | mcp |
| 11 | `execute-a-simple-mail-merge-with-a-single-data-object-and-save-the-result-as-docx.cs` | Execute a simple mail merge with a single data object and save the result as DOCX. | simple-mail-merge | docx, doc | mcp |
| 12 | `execute-a-simple-mail-merge-for-multiple-records-and-generate-separate-pdf-files-for-each.cs` | Execute a simple mail merge for multiple records and generate separate PDF files for each record. | simple-mail-merge | pdf | mcp |
| 13 | `use-mailmerge-execute-with-a-collection-of-objects-to-create-a-batch-of-merged-documents.cs` | Use MailMerge.Execute with a collection of objects to create a batch of merged documents. | input-bootstrap | docx | mcp |
| 14 | `execute-a-mail-merge-with-regions-to-repeat-a-product-list-for-each-order-record.cs` | Execute a mail merge with regions to repeat a product list for each order record. | region-mail-merge | docx | mcp |
| 15 | `use-mailmerge-executewithregions-to-merge-data-into-multiple-nested-regions-within-the-tem.cs` | Use MailMerge.ExecuteWithRegions to merge data into multiple nested regions within the template. | region-mail-merge | docx | mcp |
| 16 | `clone-the-template-document-after-each-merge-to-produce-independent-output-files.cs` | Clone the template document after each merge to produce independent output files. | simple-mail-merge | docx | mcp |
| 17 | `retrieve-mail-merge-region-metadata-using-mailmergeregioninfo-to-verify-start-and-end-posi.cs` | Retrieve mail merge region metadata using MailMergeRegionInfo to verify start and end positions. | region-mail-merge | docx | llm |
| 18 | `save-the-merged-document-as-pdf-using-document-save-with-saveformat-pdf-after-mail-merge.cs` | Save the merged document as PDF using Document.Save with SaveFormat.Pdf after mail merge. | simple-mail-merge | pdf | mcp |
| 19 | `save-the-merged-document-as-docx-using-document-save-with-saveformat-docx-after-mail-merge.cs` | Save the merged document as DOCX using Document.Save with SaveFormat.Docx after mail merge. | simple-mail-merge | docx, doc | mcp |
| 20 | `perform-mail-merge-using-xml-data-source-loaded-into-a-dataset-and-generate-docx-output.cs` | Perform mail merge using XML data source loaded into a DataSet and generate DOCX output. | simple-mail-merge | docx, xml | mcp |
| 21 | `add-static-footer-text-to-the-template-using-documentbuilder-before-executing-mail-merge.cs` | Add static footer text to the template using DocumentBuilder before executing mail merge. | simple-mail-merge | docx | mcp |
| 22 | `insert-a-page-break-field-before-each-region-to-start-new-pages-for-each-repeat.cs` | Insert a PAGE_BREAK field before each region to start new pages for each repeat. | region-mail-merge | docx | existing_repo |
| 23 | `handle-missingfieldevent-to-implement-error-handling-for-absent-merge-fields-before-execut.cs` | Handle MissingFieldEvent to implement error handling for absent merge fields before execution. | simple-mail-merge | docx | mcp |
| 24 | `adjust-image-size-during-insertion-by-modifying-imagescale-property-in-imagefieldmergingar.cs` | Adjust image size during insertion by modifying ImageScale property in ImageFieldMergingArgs. | image-mail-merge | docx | existing_repo |
| 25 | `use-mailmergeregioninfo-to-obtain-region-start-and-end-positions-for-validation-purposes.cs` | Use MailMergeRegionInfo to obtain region start and end positions for validation purposes. | region-mail-merge | docx | existing_repo |
| 26 | `perform-mail-merge-with-xml-data-source-by-loading-xml-schema-and-data-into-a-dataset.cs` | Perform mail merge with XML data source by loading XML schema and data into a DataSet. | simple-mail-merge | xml | mcp |
| 27 | `execute-simple-mail-merge-for-a-collection-of-records-and-clone-template-to-create-separat.cs` | Execute simple mail merge for a collection of records and clone template to create separate documents. | simple-mail-merge | docx | mcp |
| 28 | `customize-text-insertion-by-setting-the-text-property-with-formatted-strings-for-each-merg.cs` | Customize text insertion by setting the Text property with formatted strings for each merge field. | simple-mail-merge | docx | mcp |
| 29 | `generate-multiple-merged-documents-by-cloning-the-template-after-each-mail-merge-operation.cs` | Generate multiple merged documents by cloning the template after each mail merge operation. | simple-mail-merge | docx | mcp |
| 30 | `use-documentbuilder-to-add-a-static-table-of-contents-that-updates-after-mail-merge-execut.cs` | Use DocumentBuilder to add a static table of contents that updates after mail merge execution. | table-mail-merge | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Missing API members

- Symptom: Compilation fails with CS0246, CS1061, or CS0234 because types like Table, FieldType, MailMergeCleanupOptions, GetRegions, or ImageScale cannot be resolved.
- Root cause: The code references members that are not part of Aspose.Words 26.10.0 or were removed in newer releases.
- Preferred fix: Replace the missing members with supported alternatives: use DocumentBuilder.StartTable/EndTable instead of Table, use FieldType.FieldMergeField for field insertion, omit MailMergeCleanupOptions or use the available CleanupOptions enum, and retrieve region information via doc.MailMerge.GetRegionsHierarchy when GetRegions is unavailable.
- Prevention: Consult the Aspose.Words 26.10.0 API reference and limit code to members listed in the official documentation.

### Template bootstrap omission

- Symptom: Runtime validation reports unresolved merge fields because the document contains no MERGEFIELD tags.
- Root cause: The example assumes an external template file that was never created.
- Preferred fix: Create the template in code with DocumentBuilder before performing any mail merge operations.
- Prevention: Always include template construction steps in the sample or verify that the referenced file exists in the test environment.

### Field-name and data mismatch

- Symptom: Unmerged placeholders remain in the output or validation throws an InvalidOperationException indicating missing data.
- Root cause: The names of MERGEFIELD tags do not exactly match the keys in the string[] fieldNames array or the column names of the DataTable.
- Preferred fix: Ensure that every MERGEFIELD tag uses the exact identifier defined in the data source, respecting case sensitivity.
- Prevention: Generate field names programmatically from the data source or perform a pre-merge check that compares template fields with data columns.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Update the sample code whenever Aspose.Words releases a new major version to reflect any API changes and keep the documentation aligned with the current library.
