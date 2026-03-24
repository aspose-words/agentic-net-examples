---
name: content-control
description: C# examples for content-control using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - content-control

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **content-control** category.
This folder contains standalone C# examples for content-control operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **content-control**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (26/26 files) ← category-specific
- `using Aspose.Words.Markup;` (24/26 files)
- `using System;` (23/26 files)
- `using System.IO;` (8/26 files)
- `using Aspose.Words.Saving;` (4/26 files)
- `using System.Text;` (3/26 files)
- `using Aspose.Words.Fields;` (3/26 files)
- `using System.Drawing;` (2/26 files)
- `using System.Xml;` (2/26 files)
- `using System.Collections.Generic;` (2/26 files)
- `using System.Linq;` (2/26 files)
- `using Aspose.Words.Drawing;` (2/26 files)

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
| [add-picture-content-control-that-references-external-im...](./add-picture-content-control-that-references-external-image-file-embed-it.cs) | `Document`, `DocumentBuilder`, `StructuredDocumentTag` | Add picture content control that references external image file embed it |
| [apply-custom-style-text-inside-rich-text-content-contro...](./apply-custom-style-text-inside-rich-text-content-control-programmatically.cs) | `Font`, `Document`, `DocumentBuilder` | Apply custom style text inside rich text content control programmatically |
| [apply-custom-xml-mapping-plain-text-content-control-syn...](./apply-custom-xml-mapping-plain-text-content-control-synchronize-external-data-fields.cs) | `Document`, `StructuredDocumentTag`, `Guid` | Apply custom xml mapping plain text content control synchronize external data... |
| [batch-process-folder-word-files-inserting-header-conten...](./batch-process-folder-word-files-inserting-header-content-control-document-metadata.cs) | `BuiltInDocumentProperties`, `Document`, `DocumentBuilder` | Batch process folder word files inserting header content control document met... |
| [bind-dropdown-list-content-control-xml-data-source-popu...](./bind-dropdown-list-content-control-xml-data-source-populate-options-dynamically.cs) | `ListItems`, `Document`, `StructuredDocumentTag` | Bind dropdown list content control xml data source populate options dynamically |
| [configure-content-control-allow-only-numeric-input-enfo...](./configure-content-control-allow-only-numeric-input-enforce-validation-during-editing.cs) | `Document`, `DocumentBuilder`, `TextFormFieldType` | Configure content control allow only numeric input enforce validation during... |
| [content-control-embed-hyperlink-verify-its-target-url-a...](./content-control-embed-hyperlink-verify-its-target-url-after-document-conversion.cs) | `Document`, `DocumentBuilder`, `StructuredDocumentTag` | Content control embed hyperlink verify its target url after document conversion |
| [content-control-store-custom-metadata-extract-it-indexi...](./content-control-store-custom-metadata-extract-it-indexing-search-engine.cs) | `Document`, `DocumentBuilder`, `StructuredDocumentTag` | Content control store custom metadata extract it indexing search engine |
| [convert-docx-document-containing-content-controls-pdf-w...](./convert-docx-document-containing-content-controls-pdf-while-preserving-control.cs) | `Document`, `DocumentBuilder`, `StructuredDocumentTag` | Convert docx document containing content controls pdf while preserving control |
| [detect-list-any-nested-content-controls-within-repeatin...](./detect-list-any-nested-content-controls-within-repeating-section-structural-inspection.cs) | `StructuredDocumentTag`, `SdtType`, `MarkupLevel` | Detect list any nested content controls within repeating section structural i... |
| [doc-file-add-date-picker-content-control-result-as-docx](./doc-file-add-date-picker-content-control-result-as-docx.cs) | `DocumentBuilder`, `Document`, `StructuredDocumentTag` | Doc file add date picker content control result as docx |
| [export-contents-all-checkbox-content-controls-csv-file-...](./export-contents-all-checkbox-content-controls-csv-file-data-analysis.cs) | `Document`, `StringBuilder`, `NodeType` | Export contents all checkbox content controls csv file data analysis |
| [export-document-containing-content-controls-xps-format-...](./export-document-containing-content-controls-xps-format-while-preserving-control.cs) | `Document`, `DocumentBuilder`, `StructuredDocumentTag` | Export document containing content controls xps format while preserving control |
| [implement-error-handling-missing-xml-nodes-when-binding...](./implement-error-handling-missing-xml-nodes-when-binding-data-content-control.cs) | `XmlMapping`, `Document`, `StructuredDocumentTag` | Implement error handling missing xml nodes when binding data content control |
| [iterate-through-all-content-controls-document-summary-r...](./iterate-through-all-content-controls-document-summary-report-their-types.cs) | `Document`, `StringBuilder`, `DocumentBuilder` | Iterate through all content controls document summary report their types |
| [lock-content-control-prevent-user-editing-enforce-read-...](./lock-content-control-prevent-user-editing-enforce-read-only-behavior-final-document.cs) | `Document`, `DocumentBuilder`, `StructuredDocumentTag` | Lock content control prevent user editing enforce read only behavior final do... |
| [pdf-compliant-document-word-file-while-keeping-content-...](./pdf-compliant-document-word-file-while-keeping-content-control-tags-intact.cs) | `Document`, `DocumentBuilder`, `PdfSaveOptions` | Pdf compliant document word file while keeping content control tags intact |
| [programmatically-clear-contents-content-control-without...](./programmatically-clear-contents-content-control-without-deleting-control-itself.cs) | `StructuredDocumentTag`, `Document`, `DocumentBuilder` | Programmatically clear contents content control without deleting control itself |
| [programmatically-set-title-tag-properties-content-contr...](./programmatically-set-title-tag-properties-content-control-later-identification.cs) | `Document`, `StructuredDocumentTag`, `DocumentBuilder` | Programmatically set title tag properties content control later identification |
| [remove-all-picture-content-controls-document-replace-th...](./remove-all-picture-content-controls-document-replace-them-inline-images.cs) | `Document`, `Shape`, `NodeType` | Remove all picture content controls document replace them inline images |
| [repeating-section-content-control-that-repeats-table-ro...](./repeating-section-content-control-that-repeats-table-row-each-item-collection.cs) | `StructuredDocumentTag`, `SdtType`, `MarkupLevel` | Repeating section content control that repeats table row each item collection |
| [replace-placeholder-text-content-control-values-diction...](./replace-placeholder-text-content-control-values-dictionary-user-inputs.cs) | `Document`, `DocumentBuilder`, `Collections` | Replace placeholder text content control values dictionary user inputs |
| [retrieve-inner-xml-content-control-transform-it-xslt-st...](./retrieve-inner-xml-content-control-transform-it-xslt-stylesheet.cs) | `XmlDocument`, `Document`, `XslCompiledTransform` | Retrieve inner xml content control transform it xslt stylesheet |
| [set-placeholder-text-color-inside-content-control-match...](./set-placeholder-text-color-inside-content-control-match-document-theme.cs) | `Document`, `DocumentBuilder`, `StructuredDocumentTag` | Set placeholder text color inside content control match document theme |
| [update-tag-all-content-controls-document-follow-standar...](./update-tag-all-content-controls-document-follow-standardized-naming-convention.cs) | `Document`, `Input`, `Range` | Update tag all content controls document follow standardized naming convention |
| [validate-that-required-content-controls-contain-non-emp...](./validate-that-required-content-controls-contain-non-empty-text-before-document.cs) | `Document`, `InvalidOperationException`, `Input` | Validate that required content controls contain non empty text before document |

## Category Statistics
- Total examples: 26

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for content-control patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
