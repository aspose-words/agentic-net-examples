---
name: mail-merge
description: Verified C# examples for mail-merge scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Mail Merge

## Purpose

The AI coding agent should generate reliable Aspose.Words mail-merge programs that build the template in code, insert exact MERGEFIELD names, configure the appropriate MailMerge API (Execute or ExecuteWithRegions), apply cleanup options when required, and handle image insertion through the IFieldMergingCallback interface without referencing unavailable members.

## Non-negotiable conventions

- Do not reference types such as Table, FieldType, MailMergeCleanupOptions, GetRegions, or ImageScale that are not present in the referenced Aspose.Words version.
- Do not assume that a template file exists on disk; always construct the template programmatically or load a known file.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Simple Mail Merge workflow**: 17 examples
- **Region Mail Merge workflow**: 7 examples
- **Image Mail Merge workflow**: 4 examples
- **Input Bootstrap workflow**: 1 examples
- **Table Mail Merge workflow**: 1 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Verify that every MERGEFIELD name in the template appears in the supplied field-name collection or DataTable before executing the merge.

## File-to-task reference

- `create-a-mail-merge-template-programmatically-using-documentbuilder-and-add-static-header.cs`
  - Task: Create a mail merge template programmatically using DocumentBuilder and add static header text.
  - Workflow: simple-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `insert-merge-fields-for-customer-name-and-address-into-the-template-with-documentbuilder.cs`
  - Task: Insert merge fields for customer name and address into the template with DocumentBuilder.
  - Workflow: simple-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `define-a-mail-merge-region-for-order-items-by-inserting-start-and-end-merge-fields.cs`
  - Task: Define a mail merge region for order items by inserting start and end merge fields.
  - Workflow: region-mail-merge
  - Outputs: docx
  - Selected engine: llm
- `insert-a-table-placeholder-and-define-a-mail-merge-region-for-table-rows-using-documentbui.cs`
  - Task: Insert a table placeholder and define a mail merge region for table rows using DocumentBuilder.
  - Workflow: region-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `handle-the-imagefieldmerging-event-to-customize-image-insertion-using-imagefieldmergingarg.cs`
  - Task: Handle the ImageFieldMerging event to customize image insertion using ImageFieldMergingArgs.
  - Workflow: image-mail-merge
  - Outputs: docx
  - Selected engine: llm
- `insert-a-company-logo-into-each-merged-document-by-handling-imagefieldmerging-with-a-stati.cs`
  - Task: Insert a company logo into each merged document by handling ImageFieldMerging with a static image path.
  - Workflow: image-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `apply-conditional-logic-in-imagefieldmerging-to-select-different-images-based-on-field-nam.cs`
  - Task: Apply conditional logic in ImageFieldMerging to select different images based on field name.
  - Workflow: image-mail-merge
  - Outputs: docx
  - Selected engine: llm
- `set-the-text-property-of-a-merge-field-to-apply-bold-formatting-to-inserted-names.cs`
  - Task: Set the Text property of a merge field to apply bold formatting to inserted names.
  - Workflow: simple-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `set-the-text-property-to-insert-formatted-dates-using-a-specific-culture-format-in-merge-f.cs`
  - Task: Set the Text property to insert formatted dates using a specific culture format in merge fields.
  - Workflow: simple-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `load-xml-data-into-a-dataset-using-the-readxml-method-for-mail-merge-source.cs`
  - Task: Load XML data into a DataSet using the ReadXml method for mail merge source.
  - Workflow: simple-mail-merge
  - Outputs: xml
  - Selected engine: mcp
- `execute-a-simple-mail-merge-with-a-single-data-object-and-save-the-result-as-docx.cs`
  - Task: Execute a simple mail merge with a single data object and save the result as DOCX.
  - Workflow: simple-mail-merge
  - Outputs: docx, doc
  - Selected engine: mcp
- `execute-a-simple-mail-merge-for-multiple-records-and-generate-separate-pdf-files-for-each.cs`
  - Task: Execute a simple mail merge for multiple records and generate separate PDF files for each record.
  - Workflow: simple-mail-merge
  - Outputs: pdf
  - Selected engine: mcp
- `use-mailmerge-execute-with-a-collection-of-objects-to-create-a-batch-of-merged-documents.cs`
  - Task: Use MailMerge.Execute with a collection of objects to create a batch of merged documents.
  - Workflow: input-bootstrap
  - Outputs: docx
  - Selected engine: mcp
- `execute-a-mail-merge-with-regions-to-repeat-a-product-list-for-each-order-record.cs`
  - Task: Execute a mail merge with regions to repeat a product list for each order record.
  - Workflow: region-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `use-mailmerge-executewithregions-to-merge-data-into-multiple-nested-regions-within-the-tem.cs`
  - Task: Use MailMerge.ExecuteWithRegions to merge data into multiple nested regions within the template.
  - Workflow: region-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `clone-the-template-document-after-each-merge-to-produce-independent-output-files.cs`
  - Task: Clone the template document after each merge to produce independent output files.
  - Workflow: simple-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-mail-merge-region-metadata-using-mailmergeregioninfo-to-verify-start-and-end-posi.cs`
  - Task: Retrieve mail merge region metadata using MailMergeRegionInfo to verify start and end positions.
  - Workflow: region-mail-merge
  - Outputs: docx
  - Selected engine: llm
- `save-the-merged-document-as-pdf-using-document-save-with-saveformat-pdf-after-mail-merge.cs`
  - Task: Save the merged document as PDF using Document.Save with SaveFormat.Pdf after mail merge.
  - Workflow: simple-mail-merge
  - Outputs: pdf
  - Selected engine: mcp
- `save-the-merged-document-as-docx-using-document-save-with-saveformat-docx-after-mail-merge.cs`
  - Task: Save the merged document as DOCX using Document.Save with SaveFormat.Docx after mail merge.
  - Workflow: simple-mail-merge
  - Outputs: docx, doc
  - Selected engine: mcp
- `perform-mail-merge-using-xml-data-source-loaded-into-a-dataset-and-generate-docx-output.cs`
  - Task: Perform mail merge using XML data source loaded into a DataSet and generate DOCX output.
  - Workflow: simple-mail-merge
  - Outputs: docx, xml
  - Selected engine: mcp
- `add-static-footer-text-to-the-template-using-documentbuilder-before-executing-mail-merge.cs`
  - Task: Add static footer text to the template using DocumentBuilder before executing mail merge.
  - Workflow: simple-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-page-break-field-before-each-region-to-start-new-pages-for-each-repeat.cs`
  - Task: Insert a PAGE_BREAK field before each region to start new pages for each repeat.
  - Workflow: region-mail-merge
  - Outputs: docx
  - Selected engine: existing_repo
- `handle-missingfieldevent-to-implement-error-handling-for-absent-merge-fields-before-execut.cs`
  - Task: Handle MissingFieldEvent to implement error handling for absent merge fields before execution.
  - Workflow: simple-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `adjust-image-size-during-insertion-by-modifying-imagescale-property-in-imagefieldmergingar.cs`
  - Task: Adjust image size during insertion by modifying ImageScale property in ImageFieldMergingArgs.
  - Workflow: image-mail-merge
  - Outputs: docx
  - Selected engine: existing_repo
- `use-mailmergeregioninfo-to-obtain-region-start-and-end-positions-for-validation-purposes.cs`
  - Task: Use MailMergeRegionInfo to obtain region start and end positions for validation purposes.
  - Workflow: region-mail-merge
  - Outputs: docx
  - Selected engine: existing_repo
- `perform-mail-merge-with-xml-data-source-by-loading-xml-schema-and-data-into-a-dataset.cs`
  - Task: Perform mail merge with XML data source by loading XML schema and data into a DataSet.
  - Workflow: simple-mail-merge
  - Outputs: xml
  - Selected engine: mcp
- `execute-simple-mail-merge-for-a-collection-of-records-and-clone-template-to-create-separat.cs`
  - Task: Execute simple mail merge for a collection of records and clone template to create separate documents.
  - Workflow: simple-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `customize-text-insertion-by-setting-the-text-property-with-formatted-strings-for-each-merg.cs`
  - Task: Customize text insertion by setting the Text property with formatted strings for each merge field.
  - Workflow: simple-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `generate-multiple-merged-documents-by-cloning-the-template-after-each-mail-merge-operation.cs`
  - Task: Generate multiple merged documents by cloning the template after each mail merge operation.
  - Workflow: simple-mail-merge
  - Outputs: docx
  - Selected engine: mcp
- `use-documentbuilder-to-add-a-static-table-of-contents-that-updates-after-mail-merge-execut.cs`
  - Task: Use DocumentBuilder to add a static table of contents that updates after mail merge execution.
  - Workflow: table-mail-merge
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Missing API members**
  - Symptom: Compilation fails with CS0246, CS1061, or CS0234 because types like Table, FieldType, MailMergeCleanupOptions, GetRegions, or ImageScale cannot be resolved.
  - Root cause: The code references members that are not part of Aspose.Words 26.10.0 or were removed in newer releases.
  - Preferred fix: Replace the missing members with supported alternatives: use DocumentBuilder.StartTable/EndTable instead of Table, use FieldType.FieldMergeField for field insertion, omit MailMergeCleanupOptions or use the available CleanupOptions enum, and retrieve region information via doc.MailMerge.GetRegionsHierarchy when GetRegions is unavailable.
  - Prevention: Consult the Aspose.Words 26.10.0 API reference and limit code to members listed in the official documentation.

- **Template bootstrap omission**
  - Symptom: Runtime validation reports unresolved merge fields because the document contains no MERGEFIELD tags.
  - Root cause: The example assumes an external template file that was never created.
  - Preferred fix: Create the template in code with DocumentBuilder before performing any mail merge operations.
  - Prevention: Always include template construction steps in the sample or verify that the referenced file exists in the test environment.

- **Field-name and data mismatch**
  - Symptom: Unmerged placeholders remain in the output or validation throws an InvalidOperationException indicating missing data.
  - Root cause: The names of MERGEFIELD tags do not exactly match the keys in the string[] fieldNames array or the column names of the DataTable.
  - Preferred fix: Ensure that every MERGEFIELD tag uses the exact identifier defined in the data source, respecting case sensitivity.
  - Prevention: Generate field names programmatically from the data source or perform a pre-merge check that compares template fields with data columns.

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
Copy-Item ..\mail-merge\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update the sample code whenever Aspose.Words releases a new major version to reflect any API changes and keep the documentation aligned with the current library.
