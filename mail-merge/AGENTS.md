---
name: mail-merge
description: Verified C# examples for Mail Merge scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Mail Merge

## Purpose

This folder is a live, curated example set for Mail Merge scenarios. Treat every `.cs` file as a standalone console application. The goal is correct, warning-free examples that use documented Aspose APIs and match the original task intent.

## Non-negotiable conventions

- Always create the mail merge template locally using DocumentBuilder.
- Insert merge fields explicitly with builder.InsertField("MERGEFIELD FieldName").
- Use MailMerge.Execute for simple merges.
- Use MailMerge.ExecuteWithRegions for region-based merges.
- Ensure merge field names exactly match provided data field names or DataTable column names.
- Apply MailMergeCleanupOptions when the task requires cleanup of unused fields or regions.

## Recommended workflow selection

- Image Mail Merge workflow: 4 examples
- Input Bootstrap workflow: 1 examples
- Region Mail Merge workflow: 7 examples
- Simple Mail Merge workflow: 17 examples
- Table Mail Merge workflow: 1 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. API usage must be supported by the configured package versions.
3. Exported outputs must actually be written by the example.
4. Validation scenarios must inspect only the behavior requested by the task.
5. Examples that depend on files, folders, streams, images, or data should bootstrap those inputs locally during the example run.

## File-to-task reference

- `create-a-mail-merge-template-programmatically-using-documentbuilder-and-add-static-header.cs`
  - Task: Create a mail merge template programmatically using DocumentBuilder and add static header text.
  - Workflow: Simple Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `insert-merge-fields-for-customer-name-and-address-into-the-template-with-documentbuilder.cs`
  - Task: Insert merge fields for customer name and address into the template with DocumentBuilder.
  - Workflow: Simple Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `define-a-mail-merge-region-for-order-items-by-inserting-start-and-end-merge-fields.cs`
  - Task: Define a mail merge region for order items by inserting start and end merge fields.
  - Workflow: Region Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-table-placeholder-and-define-a-mail-merge-region-for-table-rows-using-documentbui.cs`
  - Task: Insert a table placeholder and define a mail merge region for table rows using DocumentBuilder.
  - Workflow: Region Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `handle-the-imagefieldmerging-event-to-customize-image-insertion-using-imagefieldmergingarg.cs`
  - Task: Handle the ImageFieldMerging event to customize image insertion using ImageFieldMergingArgs.
  - Workflow: Image Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-company-logo-into-each-merged-document-by-handling-imagefieldmerging-with-a-stati.cs`
  - Task: Insert a company logo into each merged document by handling ImageFieldMerging with a static image path.
  - Workflow: Image Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `apply-conditional-logic-in-imagefieldmerging-to-select-different-images-based-on-field-nam.cs`
  - Task: Apply conditional logic in ImageFieldMerging to select different images based on field name.
  - Workflow: Image Mail Merge
  - Outputs: docx
  - Selected engine: llm
- `set-the-text-property-of-a-merge-field-to-apply-bold-formatting-to-inserted-names.cs`
  - Task: Set the Text property of a merge field to apply bold formatting to inserted names.
  - Workflow: Simple Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `set-the-text-property-to-insert-formatted-dates-using-a-specific-culture-format-in-merge-f.cs`
  - Task: Set the Text property to insert formatted dates using a specific culture format in merge fields.
  - Workflow: Simple Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `load-xml-data-into-a-dataset-using-the-readxml-method-for-mail-merge-source.cs`
  - Task: Load XML data into a DataSet using the ReadXml method for mail merge source.
  - Workflow: Simple Mail Merge
  - Outputs: xml
  - Selected engine: mcp
- `execute-a-simple-mail-merge-with-a-single-data-object-and-save-the-result-as-docx.cs`
  - Task: Execute a simple mail merge with a single data object and save the result as DOCX.
  - Workflow: Simple Mail Merge
  - Outputs: docx, doc
  - Selected engine: mcp
- `execute-a-simple-mail-merge-for-multiple-records-and-generate-separate-pdf-files-for-each.cs`
  - Task: Execute a simple mail merge for multiple records and generate separate PDF files for each record.
  - Workflow: Simple Mail Merge
  - Outputs: pdf
  - Selected engine: mcp
- `use-mailmerge-execute-with-a-collection-of-objects-to-create-a-batch-of-merged-documents.cs`
  - Task: Use MailMerge.Execute with a collection of objects to create a batch of merged documents.
  - Workflow: Input Bootstrap
  - Outputs: docx
  - Selected engine: mcp
- `execute-a-mail-merge-with-regions-to-repeat-a-product-list-for-each-order-record.cs`
  - Task: Execute a mail merge with regions to repeat a product list for each order record.
  - Workflow: Region Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `use-mailmerge-executewithregions-to-merge-data-into-multiple-nested-regions-within-the-tem.cs`
  - Task: Use MailMerge.ExecuteWithRegions to merge data into multiple nested regions within the template.
  - Workflow: Region Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `clone-the-template-document-after-each-merge-to-produce-independent-output-files.cs`
  - Task: Clone the template document after each merge to produce independent output files.
  - Workflow: Simple Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-mail-merge-region-metadata-using-mailmergeregioninfo-to-verify-start-and-end-posi.cs`
  - Task: Retrieve mail merge region metadata using MailMergeRegionInfo to verify start and end positions.
  - Workflow: Region Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `save-the-merged-document-as-pdf-using-document-save-with-saveformat-pdf-after-mail-merge.cs`
  - Task: Save the merged document as PDF using Document.Save with SaveFormat.Pdf after mail merge.
  - Workflow: Simple Mail Merge
  - Outputs: pdf
  - Selected engine: mcp
- `save-the-merged-document-as-docx-using-document-save-with-saveformat-docx-after-mail-merge.cs`
  - Task: Save the merged document as DOCX using Document.Save with SaveFormat.Docx after mail merge.
  - Workflow: Simple Mail Merge
  - Outputs: docx, doc
  - Selected engine: mcp
- `perform-mail-merge-using-xml-data-source-loaded-into-a-dataset-and-generate-docx-output.cs`
  - Task: Perform mail merge using XML data source loaded into a DataSet and generate DOCX output.
  - Workflow: Simple Mail Merge
  - Outputs: docx, xml
  - Selected engine: mcp
- `add-static-footer-text-to-the-template-using-documentbuilder-before-executing-mail-merge.cs`
  - Task: Add static footer text to the template using DocumentBuilder before executing mail merge.
  - Workflow: Simple Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-page-break-field-before-each-region-to-start-new-pages-for-each-repeat.cs`
  - Task: Insert a PAGE_BREAK field before each region to start new pages for each repeat.
  - Workflow: Region Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `handle-missingfieldevent-to-implement-error-handling-for-absent-merge-fields-before-execut.cs`
  - Task: Handle MissingFieldEvent to implement error handling for absent merge fields before execution.
  - Workflow: Simple Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `adjust-image-size-during-insertion-by-modifying-imagescale-property-in-imagefieldmergingar.cs`
  - Task: Adjust image size during insertion by modifying ImageScale property in ImageFieldMergingArgs.
  - Workflow: Image Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `use-mailmergeregioninfo-to-obtain-region-start-and-end-positions-for-validation-purposes.cs`
  - Task: Use MailMergeRegionInfo to obtain region start and end positions for validation purposes.
  - Workflow: Region Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `perform-mail-merge-with-xml-data-source-by-loading-xml-schema-and-data-into-a-dataset.cs`
  - Task: Perform mail merge with XML data source by loading XML schema and data into a DataSet.
  - Workflow: Simple Mail Merge
  - Outputs: xml
  - Selected engine: mcp
- `execute-simple-mail-merge-for-a-collection-of-records-and-clone-template-to-create-separat.cs`
  - Task: Execute simple mail merge for a collection of records and clone template to create separate documents.
  - Workflow: Simple Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `customize-text-insertion-by-setting-the-text-property-with-formatted-strings-for-each-merg.cs`
  - Task: Customize text insertion by setting the Text property with formatted strings for each merge field.
  - Workflow: Simple Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `generate-multiple-merged-documents-by-cloning-the-template-after-each-mail-merge-operation.cs`
  - Task: Generate multiple merged documents by cloning the template after each mail merge operation.
  - Workflow: Simple Mail Merge
  - Outputs: docx
  - Selected engine: mcp
- `use-documentbuilder-to-add-a-static-table-of-contents-that-updates-after-mail-merge-execut.cs`
  - Task: Use DocumentBuilder to add a static table of contents that updates after mail merge execution.
  - Workflow: Table Mail Merge
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- Template and data mismatch
  - Symptom: Merge fields remain unresolved because field names do not match provided data keys or columns.
  - Preferred fix: Ensure merge field names exactly match the field-name array or DataTable column names.

- Missing local template bootstrap
  - Symptom: Examples assume an existing mail merge template document already exists.
  - Preferred fix: Create the mail merge template locally using DocumentBuilder before execution.

- Wrong API choice for regions
  - Symptom: Region merges are attempted with simple Execute instead of ExecuteWithRegions.
  - Preferred fix: Use TableStart/TableEnd fields and ExecuteWithRegions with a matching DataTable.

- Cleanup options omitted
  - Symptom: Unused merge fields or regions remain in the output unexpectedly.
  - Preferred fix: Apply MailMergeCleanupOptions when the task requires unused fields or empty regions to be removed.

## Build and run contract

- Target framework: `net8.0`
- Package: `Aspose.Words` `26.5.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.5.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\mail-merge\<example-file>.cs .\Program.cs
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
