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

- `using System;` (30/30 files) ← category-specific
- `using Aspose.Words;` (30/30 files)
- `using System.Data;` (20/30 files)
- `using Aspose.Words.MailMerging;` (19/30 files)
- `using System.Collections.Generic;` (8/30 files)
- `using Aspose.Words.Fields;` (6/30 files)
- `using Aspose.Words.Tables;` (3/30 files)
- `using System.Drawing;` (2/30 files)
- `using System.Collections;` (2/30 files)
- `using Aspose.Words.Saving;` (2/30 files)
- `using System.IO;` (1/30 files)
- `using Aspose.Words.Drawing;` (1/30 files)

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
| [add-static-footer-text-template-documentbuilder-before-...](./add-static-footer-text-template-documentbuilder-before-executing-mail-merge.cs) | `Aspose`, `Document`, `DocumentBuilder` | Add static footer text template documentbuilder before executing mail merge |
| [adjust-image-size-during-insertion-modifying-imagescale...](./adjust-image-size-during-insertion-modifying-imagescale-property-imagefieldmergingargs.cs) | `Aspose`, `Document`, `DocumentBuilder` | Adjust image size during insertion modifying imagescale property imagefieldme... |
| [apply-conditional-logic-imagefieldmerging-select-differ...](./apply-conditional-logic-imagefieldmerging-select-different-images-based-field-name.cs) | `System`, `Document`, `DocumentBuilder` | Apply conditional logic imagefieldmerging select different images based field... |
| [clone-template-document-after-each-merge-produce-indepe...](./clone-template-document-after-each-merge-produce-independent-output-files.cs) | `Columns`, `Rows`, `Document` | Clone template document after each merge produce independent output files |
| [customize-text-insertion-setting-text-property-formatte...](./customize-text-insertion-setting-text-property-formatted-strings-each-merge-field.cs) | `Columns`, `Document`, `DocumentBuilder` | Customize text insertion setting text property formatted strings each merge f... |
| [define-mail-merge-region-order-items-inserting-start-en...](./define-mail-merge-region-order-items-inserting-start-end-merge-fields.cs) | `Rows`, `Document`, `DocumentBuilder` | Define mail merge region order items inserting start end merge fields |
| [documentbuilder-add-static-table-contents-that-updates-...](./documentbuilder-add-static-table-contents-that-updates-after-mail-merge-execution.cs) | `ParagraphFormat`, `StyleIdentifier`, `Rows` | Documentbuilder add static table contents that updates after mail merge execu... |
| [execute-mail-merge-regions-repeat-product-list-each-ord...](./execute-mail-merge-regions-repeat-product-list-each-order-record.cs) | `Rows`, `DataTable`, `DataSet` | Execute mail merge regions repeat product list each order record |
| [execute-simple-mail-merge-collection-records-clone-temp...](./execute-simple-mail-merge-collection-records-clone-template-separate-documents.cs) | `Document`, `Customer`, `Aspose` | Execute simple mail merge collection records clone template separate documents |
| [execute-simple-mail-merge-multiple-records-separate-pdf...](./execute-simple-mail-merge-multiple-records-separate-pdf-files-each-record.cs) | `Rows`, `Document`, `DocumentBuilder` | Execute simple mail merge multiple records separate pdf files each record |
| [execute-simple-mail-merge-single-data-object-result-as-...](./execute-simple-mail-merge-single-data-object-result-as-docx.cs) | `Document`, `DocumentBuilder`, `Aspose` | Execute simple mail merge single data object result as docx |
| [handle-imagefieldmerging-event-customize-image-insertio...](./handle-imagefieldmerging-event-customize-image-insertion-imagefieldmergingargs.cs) | `System`, `Aspose`, `Document` | Handle imagefieldmerging event customize image insertion imagefieldmergingargs |
| [handle-missingfieldevent-implement-error-handling-absen...](./handle-missingfieldevent-implement-error-handling-absent-merge-fields-before-execution.cs) | `Document`, `DocumentBuilder`, `DataTable` | Handle missingfieldevent implement error handling absent merge fields before... |
| [insert-company-logo-each-merged-document-handling-image...](./insert-company-logo-each-merged-document-handling-imagefieldmerging-static-image-path.cs) | `MergeFieldImageDimension`, `Document`, `DataTable` | Insert company logo each merged document handling imagefieldmerging static im... |
| [insert-merge-fields-customer-name-address-template-docu...](./insert-merge-fields-customer-name-address-template-documentbuilder.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert merge fields customer name address template documentbuilder |
| [insert-page-break-field-before-each-region-start-new-pa...](./insert-page-break-field-before-each-region-start-new-pages-each-repeat.cs) | `Aspose`, `Words`, `Document` | Insert page break field before each region start new pages each repeat |
| [insert-table-placeholder-define-mail-merge-region-table...](./insert-table-placeholder-define-mail-merge-region-table-rows-documentbuilder.cs) | `DataTable`, `Rows`, `Document` | Insert table placeholder define mail merge region table rows documentbuilder |
| [mail-merge-template-programmatically-documentbuilder-ad...](./mail-merge-template-programmatically-documentbuilder-add-static-header-text.cs) | `Font`, `Document`, `DocumentBuilder` | Mail merge template programmatically documentbuilder add static header text |
| [mailmerge-execute-collection-objects-batch-merged-docum...](./mailmerge-execute-collection-objects-batch-merged-documents.cs) | `Customer`, `Document`, `DocumentBuilder` | Mailmerge execute collection objects batch merged documents |
| [mailmerge-executewithregions-merge-data-multiple-nested...](./mailmerge-executewithregions-merge-data-multiple-nested-regions-within-template.cs) | `Order`, `Aspose`, `Orders` | Mailmerge executewithregions merge data multiple nested regions within template |
| [mailmergeregioninfo-obtain-region-start-end-positions-v...](./mailmergeregioninfo-obtain-region-start-end-positions-validation-purposes.cs) | `Aspose`, `Document`, `Words` | Mailmergeregioninfo obtain region start end positions validation purposes |
| [merged-document-as-docx-document-saveformat-docx-after-...](./merged-document-as-docx-document-saveformat-docx-after-mail-merge.cs) | `Document`, `DocumentBuilder`, `Aspose` | Merged document as docx document saveformat docx after mail merge |
| [merged-document-as-pdf-document-saveformat-pdf-after-ma...](./merged-document-as-pdf-document-saveformat-pdf-after-mail-merge.cs) | `Columns`, `DataTable`, `Document` | Merged document as pdf document saveformat pdf after mail merge |
| [multiple-merged-documents-cloning-template-after-each-m...](./multiple-merged-documents-cloning-template-after-each-mail-merge-operation.cs) | `Document`, `Aspose`, `Columns` | Multiple merged documents cloning template after each mail merge operation |
| [perform-mail-merge-xml-data-source-loaded-dataset-docx-...](./perform-mail-merge-xml-data-source-loaded-dataset-docx-output.cs) | `DataSet`, `Document`, `Aspose` | Perform mail merge xml data source loaded dataset docx output |
| [perform-mail-merge-xml-data-source-xml-schema-data-dataset](./perform-mail-merge-xml-data-source-xml-schema-data-dataset.cs) | `DataSet`, `Document`, `System` | Perform mail merge xml data source xml schema data dataset |
| [retrieve-mail-merge-region-metadata-mailmergeregioninfo...](./retrieve-mail-merge-region-metadata-mailmergeregioninfo-verify-start-end-positions.cs) | `Aspose`, `Document`, `Words` | Retrieve mail merge region metadata mailmergeregioninfo verify start end posi... |
| [set-text-property-insert-formatted-dates-specific-cultu...](./set-text-property-insert-formatted-dates-specific-culture-format-merge-fields.cs) | `Aspose`, `Document`, `DocumentBuilder` | Set text property insert formatted dates specific culture format merge fields |
| [set-text-property-merge-field-apply-bold-formatting-ins...](./set-text-property-merge-field-apply-bold-formatting-inserted-names.cs) | `DocumentBuilder`, `Document`, `DataTable` | Set text property merge field apply bold formatting inserted names |
| [xml-data-dataset-readxml-method-mail-merge-source](./xml-data-dataset-readxml-method-mail-merge-source.cs) | `DataSet`, `Document`, `System` | Xml data dataset readxml method mail merge source |

## Category Statistics
- Total examples: 30

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for mail-merge patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082503`
<!-- AUTOGENERATED:END -->
