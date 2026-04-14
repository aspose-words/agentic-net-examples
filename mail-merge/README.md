# Mail Merge Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Mail Merge category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Mail Merge**
- Slug: **mail-merge**
- Total examples: **30**
- Publish-ready successful examples: **30 / 30**
- Simple mail merge workflow: **17**
- Region mail merge workflow: **7**
- Table/DataTable mail merge workflow: **1**
- Image mail merge workflow: **4**
- Input-bootstrap workflow: **1**

## Category rules that shaped these examples

- Create the mail merge template locally with DocumentBuilder.
- Insert merge fields explicitly using MERGEFIELD.
- Use MailMerge.Execute for simple merges and ExecuteWithRegions for region merges.
- Ensure field names match data exactly and apply cleanup options when required.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`
- Newtonsoft.Json for reporting tasks when needed

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Newtonsoft.Json
Copy-Item ..\mail-merge\<example-file>.cs .\Program.cs
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```
## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
```

### PowerShell example

```powershell
Copy-Item ..\mail-merge\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `create-a-mail-merge-template-programmatically-using-documentbuilder-and-add-static-header.cs` | Create a mail merge template programmatically using DocumentBuilder and add static header text. | simple-mail-merge | docx | verified |
| 2 | `insert-merge-fields-for-customer-name-and-address-into-the-template-with-documentbuilder.cs` | Insert merge fields for customer name and address into the template with DocumentBuilder. | simple-mail-merge | docx | verified |
| 3 | `define-a-mail-merge-region-for-order-items-by-inserting-start-and-end-merge-fields.cs` | Define a mail merge region for order items by inserting start and end merge fields. | region-mail-merge | docx | verified |
| 4 | `insert-a-table-placeholder-and-define-a-mail-merge-region-for-table-rows-using-documentbui.cs` | Insert a table placeholder and define a mail merge region for table rows using DocumentBuilder. | region-mail-merge | docx | verified |
| 5 | `handle-the-imagefieldmerging-event-to-customize-image-insertion-using-imagefieldmergingarg.cs` | Handle the ImageFieldMerging event to customize image insertion using ImageFieldMergingArgs. | image-mail-merge | docx | verified |
| 6 | `insert-a-company-logo-into-each-merged-document-by-handling-imagefieldmerging-with-a-stati.cs` | Insert a company logo into each merged document by handling ImageFieldMerging with a static image path. | image-mail-merge | docx | verified |
| 7 | `apply-conditional-logic-in-imagefieldmerging-to-select-different-images-based-on-field-nam.cs` | Apply conditional logic in ImageFieldMerging to select different images based on field name. | image-mail-merge | docx | verified |
| 8 | `set-the-text-property-of-a-merge-field-to-apply-bold-formatting-to-inserted-names.cs` | Set the Text property of a merge field to apply bold formatting to inserted names. | simple-mail-merge | docx | verified |
| 9 | `set-the-text-property-to-insert-formatted-dates-using-a-specific-culture-format-in-merge-f.cs` | Set the Text property to insert formatted dates using a specific culture format in merge fields. | simple-mail-merge | docx | verified |
| 10 | `load-xml-data-into-a-dataset-using-the-readxml-method-for-mail-merge-source.cs` | Load XML data into a DataSet using the ReadXml method for mail merge source. | simple-mail-merge | xml | verified |
| 11 | `execute-a-simple-mail-merge-with-a-single-data-object-and-save-the-result-as-docx.cs` | Execute a simple mail merge with a single data object and save the result as DOCX. | simple-mail-merge | docx, doc | verified |
| 12 | `execute-a-simple-mail-merge-for-multiple-records-and-generate-separate-pdf-files-for-each.cs` | Execute a simple mail merge for multiple records and generate separate PDF files for each record. | simple-mail-merge | pdf | verified |
| 13 | `use-mailmerge-execute-with-a-collection-of-objects-to-create-a-batch-of-merged-documents.cs` | Use MailMerge.Execute with a collection of objects to create a batch of merged documents. | input-bootstrap | docx | verified |
| 14 | `execute-a-mail-merge-with-regions-to-repeat-a-product-list-for-each-order-record.cs` | Execute a mail merge with regions to repeat a product list for each order record. | region-mail-merge | docx | verified |
| 15 | `use-mailmerge-executewithregions-to-merge-data-into-multiple-nested-regions-within-the-tem.cs` | Use MailMerge.ExecuteWithRegions to merge data into multiple nested regions within the template. | region-mail-merge | docx | verified |
| 16 | `clone-the-template-document-after-each-merge-to-produce-independent-output-files.cs` | Clone the template document after each merge to produce independent output files. | simple-mail-merge | docx | verified |
| 17 | `retrieve-mail-merge-region-metadata-using-mailmergeregioninfo-to-verify-start-and-end-posi.cs` | Retrieve mail merge region metadata using MailMergeRegionInfo to verify start and end positions. | region-mail-merge | docx | verified |
| 18 | `save-the-merged-document-as-pdf-using-document-save-with-saveformat-pdf-after-mail-merge.cs` | Save the merged document as PDF using Document.Save with SaveFormat.Pdf after mail merge. | simple-mail-merge | pdf | verified |
| 19 | `save-the-merged-document-as-docx-using-document-save-with-saveformat-docx-after-mail-merge.cs` | Save the merged document as DOCX using Document.Save with SaveFormat.Docx after mail merge. | simple-mail-merge | docx, doc | verified |
| 20 | `perform-mail-merge-using-xml-data-source-loaded-into-a-dataset-and-generate-docx-output.cs` | Perform mail merge using XML data source loaded into a DataSet and generate DOCX output. | simple-mail-merge | docx, xml | verified |
| 21 | `add-static-footer-text-to-the-template-using-documentbuilder-before-executing-mail-merge.cs` | Add static footer text to the template using DocumentBuilder before executing mail merge. | simple-mail-merge | docx | verified |
| 22 | `insert-a-page-break-field-before-each-region-to-start-new-pages-for-each-repeat.cs` | Insert a PAGE_BREAK field before each region to start new pages for each repeat. | region-mail-merge | docx | verified |
| 23 | `handle-missingfieldevent-to-implement-error-handling-for-absent-merge-fields-before-execut.cs` | Handle MissingFieldEvent to implement error handling for absent merge fields before execution. | simple-mail-merge | docx | verified |
| 24 | `adjust-image-size-during-insertion-by-modifying-imagescale-property-in-imagefieldmergingar.cs` | Adjust image size during insertion by modifying ImageScale property in ImageFieldMergingArgs. | image-mail-merge | docx | verified |
| 25 | `use-mailmergeregioninfo-to-obtain-region-start-and-end-positions-for-validation-purposes.cs` | Use MailMergeRegionInfo to obtain region start and end positions for validation purposes. | region-mail-merge | docx | verified |
| 26 | `perform-mail-merge-with-xml-data-source-by-loading-xml-schema-and-data-into-a-dataset.cs` | Perform mail merge with XML data source by loading XML schema and data into a DataSet. | simple-mail-merge | xml | verified |
| 27 | `execute-simple-mail-merge-for-a-collection-of-records-and-clone-template-to-create-separat.cs` | Execute simple mail merge for a collection of records and clone template to create separate documents. | simple-mail-merge | docx | verified |
| 28 | `customize-text-insertion-by-setting-the-text-property-with-formatted-strings-for-each-merg.cs` | Customize text insertion by setting the Text property with formatted strings for each merge field. | simple-mail-merge | docx | verified |
| 29 | `generate-multiple-merged-documents-by-cloning-the-template-after-each-mail-merge-operation.cs` | Generate multiple merged documents by cloning the template after each mail merge operation. | simple-mail-merge | docx | verified |
| 30 | `use-documentbuilder-to-add-a-static-table-of-contents-that-updates-after-mail-merge-execut.cs` | Use DocumentBuilder to add a static table of contents that updates after mail merge execution. | table-mail-merge | docx | verified |

## Notes for maintainers

- This category is now **100% publish-ready** for the current run.
