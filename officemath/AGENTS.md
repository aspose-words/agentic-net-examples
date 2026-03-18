---
name: officemath
description: C# examples for officemath using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - officemath

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **officemath** category.
This folder contains standalone C# examples for officemath operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **officemath**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (30/30 files) ← category-specific
- `using System;` (29/30 files)
- `using Aspose.Words.Math;` (26/30 files)
- `using System.IO;` (6/30 files)
- `using Aspose.Words.Saving;` (5/30 files)
- `using Aspose.Words.Fields;` (4/30 files)
- `using System.Linq;` (3/30 files)
- `using System.Collections.Generic;` (3/30 files)
- `using Aspose.Words.Loading;` (1/30 files)
- `using Aspose.Words.Replacing;` (1/30 files)
- `using System.Text;` (1/30 files)
- `using Aspose.Words.Rendering;` (1/30 files)

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
| [apply-uniform-justification-setting-all-officemath-equa...](./apply-uniform-justification-setting-all-officemath-equations-across-multiple-sections.cs) | `Document`, `Aspose`, `OfficeMathDisplayType` | Apply uniform justification setting all officemath equations across multiple... |
| [batch-convert-collection-docx-files-containing-officema...](./batch-convert-collection-docx-files-containing-officemath-pdf-while-preserving.cs) | `Aspose`, `Document`, `Words` | Batch convert collection docx files containing officemath pdf while preserving |
| [batch-process-that-inserts-predefined-officemath-equati...](./batch-process-that-inserts-predefined-officemath-equation-each-paragraph-document.cs) | `Aspose`, `Document`, `DocumentBuilder` | Batch process that inserts predefined officemath equation each paragraph docu... |
| [clone-existing-officemath-object-insert-clone-at-new-lo...](./clone-existing-officemath-object-insert-clone-at-new-location.cs) | `Document`, `Aspose`, `Paragraph` | Clone existing officemath object insert clone at new location |
| [configure-officemath-object-display-separate-line-rathe...](./configure-officemath-object-display-separate-line-rather-than-inline-within-paragraph.cs) | `Document`, `Aspose`, `Words` | Configure officemath object display separate line rather than inline within p... |
| [delete-unwanted-officemath-node-document-adjust-surroun...](./delete-unwanted-officemath-node-document-adjust-surrounding-paragraph-spacing.cs) | `Document`, `Aspose`, `ParagraphFormat` | Delete unwanted officemath node document adjust surrounding paragraph spacing |
| [docx-document-containing-officemath-equations-enumerate...](./docx-document-containing-officemath-equations-enumerate-each-equation-node.cs) | `Document`, `Aspose`, `Words` | Docx document containing officemath equations enumerate each equation node |
| [docx-file-replace-specific-officemath-equations-based-i...](./docx-file-replace-specific-officemath-equations-based-identifier-as-docx.cs) | `Document`, `FindReplaceOptions`, `Aspose` | Docx file replace specific officemath equations based identifier as docx |
| [export-document-containing-officemath-equations-pdf-ver...](./export-document-containing-officemath-equations-pdf-verify-equation-layout-remains.cs) | `Document`, `InvalidOperationException`, `PdfSaveOptions` | Export document containing officemath equations pdf verify equation layout re... |
| [extract-all-officemath-equations-document-write-them-te...](./extract-all-officemath-equations-document-write-them-text-file.cs) | `Document`, `StringBuilder`, `System` | Extract all officemath equations document write them text file |
| [function-that-returns-true-if-officemath-node-matches-s...](./function-that-returns-true-if-officemath-node-matches-specified-mathobjecttype-criteria.cs) | `Document`, `Aspose`, `MathObjectType` | Function that returns true if officemath node matches specified mathobjecttyp... |
| [insert-officemath-equation-latex-string-representation-...](./insert-officemath-equation-latex-string-representation-specific-document-location.cs) | `Aspose`, `Document`, `DocumentBuilder` | Insert officemath equation latex string representation specific document loca... |
| [insert-officemath-equation-mathml-string-paragraph-docu...](./insert-officemath-equation-mathml-string-paragraph-documentbuilder.cs) | `Aspose`, `Document`, `DocumentBuilder` | Insert officemath equation mathml string paragraph documentbuilder |
| [iterate-over-all-officemath-nodes-document-count-total-...](./iterate-over-all-officemath-nodes-document-count-total-number-equations.cs) | `Document`, `Aspose`, `Words` | Iterate over all officemath nodes document count total number equations |
| [macro-that-toggles-display-mode-selected-officemath-equ...](./macro-that-toggles-display-mode-selected-officemath-equations-between-inline-separate.cs) | `OfficeMathDisplayType`, `Document`, `Aspose` | Macro that toggles display mode selected officemath equations between inline... |
| [modified-document-as-docx-while-preserving-all-officema...](./modified-document-as-docx-while-preserving-all-officemath-equations-their-formatting.cs) | `Document`, `DocumentBuilder`, `Aspose` | Modified document as docx while preserving all officemath equations their for... |
| [multiple-docx-files-iterate-officemath-equations-standa...](./multiple-docx-files-iterate-officemath-equations-standardize-justification-across-all.cs) | `Document`, `System`, `Aspose` | Multiple docx files iterate officemath equations standardize justification ac... |
| [new-officemath-object-documentbuilder-insert-it-as-inli...](./new-officemath-object-documentbuilder-insert-it-as-inline-equation.cs) | `Aspose`, `Document`, `DocumentBuilder` | New officemath object documentbuilder insert it as inline equation |
| [programmatically-change-all-officemath-equations-inline...](./programmatically-change-all-officemath-equations-inline-separate-line-display-large.cs) | `Document`, `Aspose`, `MathObjectType` | Programmatically change all officemath equations inline separate line display... |
| [programmatically-set-display-mode-officemath-equations-...](./programmatically-set-display-mode-officemath-equations-inline-compact-document-layout.cs) | `Document`, `Aspose`, `Input` | Programmatically set display mode officemath equations inline compact documen... |
| [read-mathobjecttype-officemath-nodes-log-any-unsupporte...](./read-mathobjecttype-officemath-nodes-log-any-unsupported-equation-types-review.cs) | `MathObjectType`, `Aspose`, `Document` | Read mathobjecttype officemath nodes log any unsupported equation types review |
| [replace-all-inline-officemath-equations-separate-line-d...](./replace-all-inline-officemath-equations-separate-line-display-enhance-visual-clarity.cs) | `Document`, `Aspose`, `OfficeMathDisplayType` | Replace all inline officemath equations separate line display enhance visual... |
| [replace-content-existing-officemath-object-new-equation...](./replace-content-existing-officemath-object-new-equation-defined-string.cs) | `Aspose`, `Document`, `DocumentBuilder` | Replace content existing officemath object new equation defined string |
| [report-listing-each-officemath-equation-s-mathobjecttyp...](./report-listing-each-officemath-equation-s-mathobjecttype-its-position-within-document.cs) | `Document`, `DocumentBuilder`, `Aspose` | Report listing each officemath equation s mathobjecttype its position within... |
| [retrieve-mathobjecttype-each-officemath-node-determine-...](./retrieve-mathobjecttype-each-officemath-node-determine-whether-it-is-fraction-radical.cs) | `MathObjectType`, `Document`, `Aspose` | Retrieve mathobjecttype each officemath node determine whether it is fraction... |
| [set-justification-officemath-equation-center-alignment-...](./set-justification-officemath-equation-center-alignment-justification-property.cs) | `Document`, `Aspose`, `Words` | Set justification officemath equation center alignment justification property |
| [update-justification-all-officemath-equations-right-ali...](./update-justification-all-officemath-equations-right-alignment-template-document.cs) | `Document`, `Aspose`, `Words` | Update justification all officemath equations right alignment template document |
| [validate-that-after-bulk-justification-changes-no-offic...](./validate-that-after-bulk-justification-changes-no-officemath-equation-exceeds-page.cs) | `Aspose`, `Document`, `OfficeMathRenderer` | Validate that after bulk justification changes no officemath equation exceeds... |
| [validate-that-each-officemath-object-conforms-expected-...](./validate-that-each-officemath-object-conforms-expected-mathobjecttype-after.cs) | `MathObjectType`, `Aspose`, `Document` | Validate that each officemath object conforms expected mathobjecttype after |
| [validate-that-exported-pdf-retains-exact-positioning-of...](./validate-that-exported-pdf-retains-exact-positioning-officemath-equations-as-source.cs) | `SizeInPoints`, `NodeType`, `Document` | Validate that exported pdf retains exact positioning officemath equations as... |

## Category Statistics
- Total examples: 30

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for officemath patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082512`
<!-- AUTOGENERATED:END -->
