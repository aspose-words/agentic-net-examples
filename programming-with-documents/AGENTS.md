---
name: programming-with-documents
description: C# examples for programming-with-documents using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - programming-with-documents

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **programming-with-documents** category.
This folder contains standalone C# examples for programming-with-documents operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **programming-with-documents**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (106/106 files) ← category-specific
- `using System;` (89/106 files)
- `using Aspose.Words.Tables;` (37/106 files)
- `using Aspose.Words.Drawing;` (31/106 files)
- `using System.Drawing;` (17/106 files)
- `using Aspose.Words.Saving;` (16/106 files)
- `using Aspose.Words.Lists;` (12/106 files)
- `using Aspose.Words.Fields;` (10/106 files)
- `using System.IO;` (9/106 files)
- `using Aspose.Words.Notes;` (7/106 files)
- `using System.Data;` (3/106 files)
- `using System.Globalization;` (3/106 files)

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
| [activate-odd-even-page-headers-setting-section-pagesetu...](./activate-odd-even-page-headers-setting-section-pagesetup-differentoddandevenpages.cs) | `Document`, `DocumentBuilder`, `HeaderFooterType` | Activate odd even page headers setting section pagesetup differentoddandevenp... |
| [add-caption-below-table-paragraph-styled-as-caption-ref...](./add-caption-below-table-paragraph-styled-as-caption-reference-it.cs) | `DocumentBuilder`, `Document`, `ParagraphFormat` | Add caption below table paragraph styled as caption reference it |
| [add-comment-paragraph-programmatically-retrieve-comment...](./add-comment-paragraph-programmatically-retrieve-comment-text-review.cs) | `Document`, `DocumentBuilder`, `Comment` | Add comment paragraph programmatically retrieve comment text review |
| [add-endnote-reference-text-insertfootnote-specify-endno...](./add-endnote-reference-text-insertfootnote-specify-endnote-type.cs) | `Document`, `DocumentBuilder`, `FootnoteType` | Add endnote reference text insertfootnote specify endnote type |
| [add-odd-page-header-containing-right-aligned-chapter-ti...](./add-odd-page-header-containing-right-aligned-chapter-title-custom-font.cs) | `Font`, `Document`, `DocumentBuilder` | Add odd page header containing right aligned chapter title custom font |
| [add-page-break-before-each-new-heading-paragraph-ensure...](./add-page-break-before-each-new-heading-paragraph-ensure-heading-style-remains-applied.cs) | `ParagraphFormat`, `StyleIdentifier`, `Document` | Add page break before each new heading paragraph ensure heading style remains... |
| [add-paragraph-background-shading-light-yellow-highlight...](./add-paragraph-background-shading-light-yellow-highlight-important-information.cs) | `ParagraphFormat`, `Shading`, `Document` | Add paragraph background shading light yellow highlight important information |
| [add-paragraph-custom-field-that-displays-document-s-aut...](./add-paragraph-custom-field-that-displays-document-s-author-name-dynamically.cs) | `Document`, `DocumentBuilder`, `FieldType` | Add paragraph custom field that displays document s author name dynamically |
| [add-paragraph-drop-cap-character-specify-number-lines-i...](./add-paragraph-drop-cap-character-specify-number-lines-it-spans.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Add paragraph drop cap character specify number lines it spans |
| [add-paragraph-hanging-indent-0-25-inches-citation-forma...](./add-paragraph-hanging-indent-0-25-inches-citation-formatting.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Add paragraph hanging indent 0 25 inches citation formatting |
| [add-text-box-document-header-ensure-it-appears-every-pa...](./add-text-box-document-header-ensure-it-appears-every-page-section.cs) | `Document`, `DocumentBuilder`, `HeaderFooterType` | Add text box document header ensure it appears every page section |
| [adjust-column-widths-proportionally-fit-content-after-d...](./adjust-column-widths-proportionally-fit-content-after-data-table.cs) | `Document`, `Input`, `NodeType` | Adjust column widths proportionally fit content after data table |
| [adjust-left-margin-first-page-header-align-main-text-ma...](./adjust-left-margin-first-page-header-align-main-text-margin.cs) | `PageSetup`, `Document`, `DocumentBuilder` | Adjust left margin first page header align main text margin |
| [apply-bold-underline-formatting-header-text-configuring...](./apply-bold-underline-formatting-header-text-configuring-documentbuilder-font-before.cs) | `Document`, `DocumentBuilder`, `Font` | Apply bold underline formatting header text configuring documentbuilder font... |
| [apply-built-grid-table-5-dark-style-entire-table-consis...](./apply-built-grid-table-5-dark-style-entire-table-consistent-formatting.cs) | `TableStyleOptions`, `Document`, `DocumentBuilder` | Apply built grid table 5 dark style entire table consistent formatting |
| [apply-built-heading-2-style-paragraph-adjust-its-spacin...](./apply-built-heading-2-style-paragraph-adjust-its-spacing-before-after.cs) | `ParagraphFormat`, `Document`, `DocumentBuilder` | Apply built heading 2 style paragraph adjust its spacing before after |
| [apply-built-quote-style-paragraph-increase-its-left-ind...](./apply-built-quote-style-paragraph-increase-its-left-indent-emphasis.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Apply built quote style paragraph increase its left indent emphasis |
| [apply-built-title-style-first-paragraph-ensure-it-appea...](./apply-built-title-style-first-paragraph-ensure-it-appears-document-outline.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Apply built title style first paragraph ensure it appears document outline |
| [apply-conditional-formatting-rule-table-cells-that-high...](./apply-conditional-formatting-rule-table-cells-that-highlights-values-above-threshold.cs) | `Document`, `DocumentBuilder`, `Rows` | Apply conditional formatting rule table cells that highlights values above th... |
| [apply-custom-list-template-all-document-lists-enforce-m...](./apply-custom-list-template-all-document-lists-enforce-maximum-nine-nesting-levels.cs) | `Font`, `ListFormat`, `Document` | Apply custom list template all document lists enforce maximum nine nesting le... |
| [apply-custom-paragraph-style-that-includes-both-left-ri...](./apply-custom-paragraph-style-that-includes-both-left-right-borders-light-gray.cs) | `ParagraphFormat`, `Color`, `Document` | Apply custom paragraph style that includes both left right borders light gray |
| [apply-custom-paragraph-style-that-includes-left-indent-...](./apply-custom-paragraph-style-that-includes-left-indent-0-5-inches-line-spacing-1-5.cs) | `ParagraphFormat`, `Document`, `DocumentBuilder` | Apply custom paragraph style that includes left indent 0 5 inches line spacin... |
| [apply-custom-paragraph-style-toc-entries-modifying-toc-...](./apply-custom-paragraph-style-toc-entries-modifying-toc-field-s-style-property.cs) | `ParagraphFormat`, `StyleIdentifier`, `Font` | Apply custom paragraph style toc entries modifying toc field s style property |
| [apply-custom-style-paragraph-that-includes-border-backg...](./apply-custom-style-paragraph-that-includes-border-background-color-indentation.cs) | `BorderType`, `Color`, `ParagraphFormat` | Apply custom style paragraph that includes border background color indentation |
| [apply-gradient-fill-text-box-ensure-gradient-renders-co...](./apply-gradient-fill-text-box-ensure-gradient-renders-correctly-pdf.cs) | `Document`, `DocumentBuilder`, `Fill` | Apply gradient fill text box ensure gradient renders correctly pdf |
| [apply-listformat-removenumbers-method-convert-numbered-...](./apply-listformat-removenumbers-method-convert-numbered-list-back-plain-paragraphs.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Apply listformat removenumbers method convert numbered list back plain paragr... |
| [apply-shadow-effect-text-box-shape-export-document-as-pdf](./apply-shadow-effect-text-box-shape-export-document-as-pdf.cs) | `ShadowFormat`, `Document`, `DocumentBuilder` | Apply shadow effect text box shape export document as pdf |
| [apply-table-style-that-includes-alternating-column-shad...](./apply-table-style-that-includes-alternating-column-shading-better-readability.cs) | `Document`, `DocumentBuilder`, `ConditionalStyleType` | Apply table style that includes alternating column shading better readability |
| [apply-table-style-that-includes-top-border-bottom-borde...](./apply-table-style-that-includes-top-border-bottom-border-no-side-borders.cs) | `Borders`, `LineStyle`, `Color` | Apply table style that includes top border bottom border no side borders |
| [batch-process-set-documents-replace-all-tables-predefin...](./batch-process-set-documents-replace-all-tables-predefined-style-margin-settings.cs) | `Document`, `Borders`, `Color` | Batch process set documents replace all tables predefined style margin settings |
| ... | | *and 76 more files* |

## Category Statistics
- Total examples: 106

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for programming-with-documents patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
