---
name: mail-merge
description: C# examples for mail-merge using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - mail-merge

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **mail-merge** category.
This folder contains standalone C# examples for mail-merge operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **mail-merge**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using System;` (29/29 files) ← category-specific
- `using Aspose.Words;` (29/29 files)
- `using Aspose.Words.MailMerging;` (19/29 files)
- `using System.Data;` (18/29 files)
- `using System.IO;` (7/29 files)
- `using System.Collections.Generic;` (7/29 files)
- `using Aspose.Words.Fields;` (6/29 files)
- `using Aspose.Words.Tables;` (3/29 files)
- `using System.Collections;` (2/29 files)
- `using Aspose.Words.Saving;` (2/29 files)
- `using Aspose.Words.Settings;` (1/29 files)
- `using System.Globalization;` (1/29 files)

## Common Code Pattern

Most files follow this pattern:

```csharp
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
// ... operations ...
doc.Save("output.docx");
```

## Files in this folder

| File | Key APIs | Description |
|------|----------|-------------|
| [add-static-footer-text-template-documentbuilder-before-...](./add-static-footer-text-template-documentbuilder-before-executing-mail-merge.cs) | `Document`, `DocumentBuilder`, `Template` | Add static footer text template documentbuilder before executing mail merge |
| [adjust-image-size-during-insertion-modifying-imagescale...](./adjust-image-size-during-insertion-modifying-imagescale-property-imagefieldmergingargs.cs) | `Document`, `DocumentBuilder`, `DataTable` | Adjust image size during insertion modifying imagescale property imagefieldme... |
| [apply-conditional-logic-imagefieldmerging-select-differ...](./apply-conditional-logic-imagefieldmerging-select-different-images-based-field-name.cs) | `IFieldMergingCallback`, `Columns`, `MailMerge` | Apply conditional logic imagefieldmerging select different images based field... |
| [clone-template-document-after-each-merge-produce-indepe...](./clone-template-document-after-each-merge-produce-independent-output-files.cs) | `Columns`, `Rows`, `Document` | Clone template document after each merge produce independent output files |
| [customize-text-insertion-setting-text-property-formatte...](./customize-text-insertion-setting-text-property-formatted-strings-each-merge-field.cs) | `Columns`, `Document`, `DocumentBuilder` | Customize text insertion setting text property formatted strings each merge f... |
| [define-mail-merge-region-order-items-inserting-start-en...](./define-mail-merge-region-order-items-inserting-start-end-merge-fields.cs) | `Rows`, `Document`, `DocumentBuilder` | Define mail merge region order items inserting start end merge fields |
| [documentbuilder-add-static-table-contents-that-updates-...](./documentbuilder-add-static-table-contents-that-updates-after-mail-merge-execution.cs) | `ParagraphFormat`, `StyleIdentifier`, `Rows` | Documentbuilder add static table contents that updates after mail merge execu... |
| [execute-mail-merge-regions-repeat-product-list-each-ord...](./execute-mail-merge-regions-repeat-product-list-each-order-record.cs) | `Rows`, `DataTable`, `Columns` | Execute mail merge regions repeat product list each order record |
| [execute-simple-mail-merge-collection-records-clone-temp...](./execute-simple-mail-merge-collection-records-clone-template-separate-documents.cs) | `Customer`, `Document`, `DocumentBuilder` | Execute simple mail merge collection records clone template separate documents |
| [execute-simple-mail-merge-multiple-records-separate-pdf...](./execute-simple-mail-merge-multiple-records-separate-pdf-files-each-record.cs) | `Rows`, `Document`, `DocumentBuilder` | Execute simple mail merge multiple records separate pdf files each record |
| [execute-simple-mail-merge-single-data-object-result-as-...](./execute-simple-mail-merge-single-data-object-result-as-docx.cs) | `Document`, `DocumentBuilder`, `MailMerge` | Execute simple mail merge single data object result as docx |
| [handle-imagefieldmerging-event-customize-image-insertio...](./handle-imagefieldmerging-event-customize-image-insertion-imagefieldmergingargs.cs) | `Document`, `DataTable`, `IFieldMergingCallback` | Handle imagefieldmerging event customize image insertion imagefieldmergingargs |
| [handle-missingfieldevent-implement-error-handling-absen...](./handle-missingfieldevent-implement-error-handling-absent-merge-fields-before-execution.cs) | `Document`, `DocumentBuilder`, `DataTable` | Handle missingfieldevent implement error handling absent merge fields before... |
| [insert-merge-fields-customer-name-address-template-docu...](./insert-merge-fields-customer-name-address-template-documentbuilder.cs) | `Document`, `DocumentBuilder`, `CustomerTemplate` | Insert merge fields customer name address template documentbuilder |
| [insert-page-break-field-before-each-region-start-new-pa...](./insert-page-break-field-before-each-region-start-new-pages-each-repeat.cs) | `DocumentBuilder`, `Document`, `Collections` | Insert page break field before each region start new pages each repeat |
| [insert-table-placeholder-define-mail-merge-region-table...](./insert-table-placeholder-define-mail-merge-region-table-rows-documentbuilder.cs) | `DataTable`, `Rows`, `Document` | Insert table placeholder define mail merge region table rows documentbuilder |
| [mail-merge-template-programmatically-documentbuilder-ad...](./mail-merge-template-programmatically-documentbuilder-add-static-header-text.cs) | `Font`, `Document`, `DocumentBuilder` | Mail merge template programmatically documentbuilder add static header text |
| [mailmerge-execute-collection-objects-batch-merged-docum...](./mailmerge-execute-collection-objects-batch-merged-documents.cs) | `Customer`, `Document`, `DocumentBuilder` | Mailmerge execute collection objects batch merged documents |
| [mailmerge-executewithregions-merge-data-multiple-nested...](./mailmerge-executewithregions-merge-data-multiple-nested-regions-within-template.cs) | `Order`, `Orders`, `CustomerList` | Mailmerge executewithregions merge data multiple nested regions within template |
| [mailmergeregioninfo-obtain-region-start-end-positions-v...](./mailmergeregioninfo-obtain-region-start-end-positions-validation-purposes.cs) | `Document`, `DocumentBuilder`, `Collections` | Mailmergeregioninfo obtain region start end positions validation purposes |
| [merged-document-as-docx-document-saveformat-docx-after-...](./merged-document-as-docx-document-saveformat-docx-after-mail-merge.cs) | `Document`, `DocumentBuilder`, `MailMerge` | Merged document as docx document saveformat docx after mail merge |
| [merged-document-as-pdf-document-saveformat-pdf-after-ma...](./merged-document-as-pdf-document-saveformat-pdf-after-mail-merge.cs) | `Document`, `Columns`, `DataTable` | Merged document as pdf document saveformat pdf after mail merge |
| [multiple-merged-documents-cloning-template-after-each-m...](./multiple-merged-documents-cloning-template-after-each-mail-merge-operation.cs) | `Document`, `Columns`, `Rows` | Multiple merged documents cloning template after each mail merge operation |
| [perform-mail-merge-xml-data-source-loaded-dataset-docx-...](./perform-mail-merge-xml-data-source-loaded-dataset-docx-output.cs) | `Document`, `DocumentBuilder`, `DataSet` | Perform mail merge xml data source loaded dataset docx output |
| [perform-mail-merge-xml-data-source-xml-schema-data-dataset](./perform-mail-merge-xml-data-source-xml-schema-data-dataset.cs) | `Document`, `DocumentBuilder`, `MailMerge` | Perform mail merge xml data source xml schema data dataset |
| [retrieve-mail-merge-region-metadata-mailmergeregioninfo...](./retrieve-mail-merge-region-metadata-mailmergeregioninfo-verify-start-end-positions.cs) | `Document`, `Collections`, `MailMergeRegions` | Retrieve mail merge region metadata mailmergeregioninfo verify start end posi... |
| [set-text-property-insert-formatted-dates-specific-cultu...](./set-text-property-insert-formatted-dates-specific-culture-format-merge-fields.cs) | `Document`, `DocumentBuilder`, `DataTable` | Set text property insert formatted dates specific culture format merge fields |
| [set-text-property-merge-field-apply-bold-formatting-ins...](./set-text-property-merge-field-apply-bold-formatting-inserted-names.cs) | `DocumentBuilder`, `Document`, `DataTable` | Set text property merge field apply bold formatting inserted names |
| [xml-data-dataset-readxml-method-mail-merge-source](./xml-data-dataset-readxml-method-mail-merge-source.cs) | `DataSet`, `Employees`, `StringReader` | Xml data dataset readxml method mail merge source |

## Category Statistics
- Total examples: 29

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for mail-merge patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
