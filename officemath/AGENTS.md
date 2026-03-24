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

- `using System;` (18/18 files) ← category-specific
- `using Aspose.Words;` (17/18 files)
- `using Aspose.Words.Math;` (14/18 files)
- `using System.IO;` (7/18 files)
- `using Aspose.Words.Saving;` (3/18 files)
- `using Aspose.Words.Fields;` (2/18 files)
- `using System.Linq;` (2/18 files)
- `using System.Collections.Generic;` (2/18 files)
- `using Aspose.Words.Loading;` (1/18 files)
- `using System.Text;` (1/18 files)
- `using Aspose.Words.Replacing;` (1/18 files)
- `using System.Drawing;` (1/18 files)

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
| [batch-convert-collection-docx-files-containing-officema...](./batch-convert-collection-docx-files-containing-officemath-pdf-while-preserving.cs) | `Document`, `PdfSaveOptions`, `LoadOptions` | Batch convert collection docx files containing officemath pdf while preserving |
| [batch-process-that-inserts-predefined-officemath-equati...](./batch-process-that-inserts-predefined-officemath-equation-each-paragraph-document.cs) | `Encoding`, `StringBuilder`, `Input` | Batch process that inserts predefined officemath equation each paragraph docu... |
| [delete-unwanted-officemath-node-document-adjust-surroun...](./delete-unwanted-officemath-node-document-adjust-surrounding-paragraph-spacing.cs) | `Document`, `ParagraphFormat`, `Input` | Delete unwanted officemath node document adjust surrounding paragraph spacing |
| [docx-file-replace-specific-officemath-equations-based-i...](./docx-file-replace-specific-officemath-equations-based-identifier-as-docx.cs) | `Document`, `FindReplaceOptions`, `Input` | Docx file replace specific officemath equations based identifier as docx |
| [function-that-returns-true-if-officemath-node-matches-s...](./function-that-returns-true-if-officemath-node-matches-specified-mathobjecttype-criteria.cs) | `Document`, `MathObjectType`, `OfficeMath` | Function that returns true if officemath node matches specified mathobjecttyp... |
| [insert-officemath-equation-mathml-string-paragraph-docu...](./insert-officemath-equation-mathml-string-paragraph-documentbuilder.cs) | `Document`, `DocumentBuilder`, `HTML` | Insert officemath equation mathml string paragraph documentbuilder |
| [iterate-over-all-officemath-nodes-document-count-total-...](./iterate-over-all-officemath-nodes-document-count-total-number-equations.cs) | `Document`, `Input`, `NodeType` | Iterate over all officemath nodes document count total number equations |
| [macro-that-toggles-display-mode-selected-officemath-equ...](./macro-that-toggles-display-mode-selected-officemath-equations-between-inline-separate.cs) | `OfficeMathDisplayType`, `Document`, `MathObjectType` | Macro that toggles display mode selected officemath equations between inline... |
| [modified-document-as-docx-while-preserving-all-officema...](./modified-document-as-docx-while-preserving-all-officemath-equations-their-formatting.cs) | `DocumentBuilder`, `Document`, `InputWithMath` | Modified document as docx while preserving all officemath equations their for... |
| [multiple-docx-files-iterate-officemath-equations-standa...](./multiple-docx-files-iterate-officemath-equations-standardize-justification-across-all.cs) | `Document`, `OfficeMathDisplayType`, `AppContext` | Multiple docx files iterate officemath equations standardize justification ac... |
| [programmatically-change-all-officemath-equations-inline...](./programmatically-change-all-officemath-equations-inline-separate-line-display-large.cs) | `Document`, `MathObjectType`, `OfficeMathDisplayType` | Programmatically change all officemath equations inline separate line display... |
| [programmatically-set-display-mode-officemath-equations-...](./programmatically-set-display-mode-officemath-equations-inline-compact-document-layout.cs) | `Document`, `Input`, `MathObjectType` | Programmatically set display mode officemath equations inline compact documen... |
| [read-mathobjecttype-officemath-nodes-log-any-unsupporte...](./read-mathobjecttype-officemath-nodes-log-any-unsupported-equation-types-review.cs) | `MathObjectType`, `Document`, `DocumentBuilder` | Read mathobjecttype officemath nodes log any unsupported equation types review |
| [replace-all-inline-officemath-equations-separate-line-d...](./replace-all-inline-officemath-equations-separate-line-display-enhance-visual-clarity.cs) | `Document`, `OfficeMathDisplayType`, `Input` | Replace all inline officemath equations separate line display enhance visual... |
| [replace-content-existing-officemath-object-new-equation...](./replace-content-existing-officemath-object-new-equation-defined-string.cs) | `Document`, `DocumentBuilder`, `Input` | Replace content existing officemath object new equation defined string |
| [report-listing-each-officemath-equation-s-mathobjecttyp...](./report-listing-each-officemath-equation-s-mathobjecttype-its-position-within-document.cs) | `Document`, `DocumentBuilder`, `OfficeMathSource` | Report listing each officemath equation s mathobjecttype its position within... |
| [validate-that-after-bulk-justification-changes-no-offic...](./validate-that-after-bulk-justification-changes-no-officemath-equation-exceeds-page.cs) | `Document`, `NodeType`, `Collections` | Validate that after bulk justification changes no officemath equation exceeds... |
| [validate-that-exported-pdf-retains-exact-positioning-of...](./validate-that-exported-pdf-retains-exact-positioning-officemath-equations-as-source.cs) | `SizeInPoints`, `Document`, `NodeType` | Validate that exported pdf retains exact positioning officemath equations as... |

## Category Statistics
- Total examples: 18

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for officemath patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
