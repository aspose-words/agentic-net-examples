---
name: fonts
description: C# examples for fonts using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - fonts

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **fonts** category.
This folder contains standalone C# examples for fonts operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **fonts**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (35/35 files) ← category-specific
- `using System;` (31/35 files)
- `using System.Drawing;` (11/35 files)
- `using Aspose.Words.Fonts;` (8/35 files)
- `using Aspose.Words.Drawing;` (6/35 files)
- `using System.IO;` (2/35 files)
- `using Aspose.Words.Tables;` (2/35 files)
- `using Aspose.Words.Saving;` (2/35 files)
- `using System.Globalization;` (1/35 files)
- `using Aspose.Words.Lists;` (1/35 files)
- `using Aspose.Words.Replacing;` (1/35 files)
- `using System.Linq;` (1/35 files)

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
| [adjust-font-color-dynamically-based-paragraph-index-loo...](./adjust-font-color-dynamically-based-paragraph-index-loop-conditional-logic.cs) | `Document`, `DocumentBuilder`, `Color` | Adjust font color dynamically based paragraph index loop conditional logic |
| [adjust-paragraph-s-line-spacing-1-5-lines-assigning-fon...](./adjust-paragraph-s-line-spacing-1-5-lines-assigning-font-linespacing-1-5.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Adjust paragraph s line spacing 1 5 lines assigning font linespacing 1 5 |
| [apply-bold-italic-styles-run-setting-font-bold-font-ita...](./apply-bold-italic-styles-run-setting-font-bold-font-italic-true.cs) | `Document`, `Run`, `Font` | Apply bold italic styles run setting font bold font italic true |
| [apply-bold-italic-underline-simultaneously-run-respecti...](./apply-bold-italic-underline-simultaneously-run-respective-font-properties.cs) | `Font`, `Document`, `Run` | Apply bold italic underline simultaneously run respective font properties |
| [apply-east-asian-emphasis-mark-dot-run-setting-font-emp...](./apply-east-asian-emphasis-mark-dot-run-setting-font-emphasismark-emphasismark-dot.cs) | `Document`, `DocumentBuilder`, `Font` | Apply east asian emphasis mark dot run setting font emphasismark emphasismark... |
| [apply-east-asian-emphasis-mark-only-when-document-langu...](./apply-east-asian-emphasis-mark-only-when-document-language-is-japanese-checking.cs) | `Font`, `Document`, `DocumentBuilder` | Apply east asian emphasis mark only when document language is japanese checking |
| [apply-predefined-style-that-includes-specific-font-sett...](./apply-predefined-style-that-includes-specific-font-settings-range-paragraphs.cs) | `Font`, `Aspose`, `Words` | Apply predefined style that includes specific font settings range paragraphs |
| [apply-semi-transparent-fill-text-setting-font-fill-colo...](./apply-semi-transparent-fill-text-setting-font-fill-color-font-fill-transparency.cs) | `Font`, `Fill`, `Document` | Apply semi transparent fill text setting font fill color font fill transparency |
| [apply-underline-style-run-setting-font-underline-underl...](./apply-underline-style-run-setting-font-underline-underlinetype-single.cs) | `Document`, `Paragraph`, `Run` | Apply underline style run setting font underline underlinetype single |
| [assign-blue-color-run-s-font-setting-font-color-system-...](./assign-blue-color-run-s-font-setting-font-color-system-drawing-color-blue.cs) | `Document`, `Run`, `Aspose` | Assign blue color run s font setting font color system drawing color blue |
| [change-font-size-run-fourteen-points-font-size-property](./change-font-size-run-fourteen-points-font-size-property.cs) | `Document`, `Run`, `Aspose` | Change font size run fourteen points font size property |
| [configure-font-substitution-map-missing-garamond-locall...](./configure-font-substitution-map-missing-garamond-locally-installed-georgia-font.cs) | `Document`, `DocumentBuilder`, `FontSettings` | Configure font substitution map missing garamond locally installed georgia font |
| [copy-font-formatting-one-run-another-run-font-clone-method](./copy-font-formatting-one-run-another-run-font-clone-method.cs) | `Font`, `Run`, `Document` | Copy font formatting one run another run font clone method |
| [custom-font-substitution-table-it-xml-configuration-file](./custom-font-substitution-table-it-xml-configuration-file.cs) | `Aspose`, `Document`, `FontSettings` | Custom font substitution table it xml configuration file |
| [define-custom-font-substitution-rule-that-replaces-miss...](./define-custom-font-substitution-rule-that-replaces-missing-times-new-roman-calibri.cs) | `Document`, `FontSettings`, `DocumentBuilder` | Define custom font substitution rule that replaces missing times new roman ca... |
| [define-style-that-sets-font-name-size-color-then-apply-...](./define-style-that-sets-font-name-size-color-then-apply-it-selected-text.cs) | `Font`, `Document`, `DocumentBuilder` | Define style that sets font name size color then apply it selected text |
| [document-change-all-headings-bold-sixteen-point-font-then](./document-change-all-headings-bold-sixteen-point-font-then.cs) | `Document`, `Aspose`, `NodeType` | Document change all headings bold sixteen point font then |
| [documentbuilder-font-name-set-default-font-newly-insert...](./documentbuilder-font-name-set-default-font-newly-inserted-content-before-adding-text.cs) | `Document`, `DocumentBuilder`, `Font` | Documentbuilder font name set default font newly inserted content before addi... |
| [documentbuilder-set-default-font-size-all-subsequently-...](./documentbuilder-set-default-font-size-all-subsequently-inserted-paragraphs.cs) | `Document`, `DocumentBuilder`, `Aspose` | Documentbuilder set default font size all subsequently inserted paragraphs |
| [embed-all-used-fonts-document-when-enabling-embedfullfo...](./embed-all-used-fonts-document-when-enabling-embedfullfonts-option.cs) | `Document`, `Aspose`, `PdfSaveOptions` | Embed all used fonts document when enabling embedfullfonts option |
| [enable-automatic-font-substitution-missing-fonts-during...](./enable-automatic-font-substitution-missing-fonts-during-document-configuring.cs) | `LoadOptions`, `Aspose`, `FontSettings` | Enable automatic font substitution missing fonts during document configuring |
| [enable-font-embedding-when-documents-pdf-ensure-visual-...](./enable-font-embedding-when-documents-pdf-ensure-visual-fidelity-across-platforms.cs) | `Document`, `PdfSaveOptions`, `Aspose` | Enable font embedding when documents pdf ensure visual fidelity across platforms |
| [font-object-set-its-size-color-then-assign-it-multiple-...](./font-object-set-its-size-color-then-assign-it-multiple-runs-consistency.cs) | `Document`, `DocumentBuilder`, `Font` | Font object set its size color then assign it multiple runs consistency |
| [implement-batch-processing-set-helvetica-font-all-runs-...](./implement-batch-processing-set-helvetica-font-all-runs-collection-documents.cs) | `Document`, `Aspose`, `System` | Implement batch processing set helvetica font all runs collection documents |
| [programmatically-replace-all-occurrences-specific-font-...](./programmatically-replace-all-occurrences-specific-font-another-across-entire-document.cs) | `Document`, `Aspose`, `Font` | Programmatically replace all occurrences specific font another across entire... |
| [programmatically-reset-line-spacing-default-all-paragra...](./programmatically-reset-line-spacing-default-all-paragraphs-setting-font-linespacing-0.cs) | `ParagraphFormat`, `Document`, `DocumentBuilder` | Programmatically reset line spacing default all paragraphs setting font lines... |
| [reset-all-font-attributes-run-defaults-calling-font-cle...](./reset-all-font-attributes-run-defaults-calling-font-clearformatting.cs) | `Document`, `Aspose`, `Input` | Reset all font attributes run defaults calling font clearformatting |
| [retrieve-display-current-font-size-run-debugging-purposes](./retrieve-display-current-font-size-run-debugging-purposes.cs) | `Document`, `Run`, `Aspose` | Retrieve display current font size run debugging purposes |
| [retrieve-display-emphasismark-value-run-debugging-east-...](./retrieve-display-emphasismark-value-run-debugging-east-asian-formatting.cs) | `Document`, `DocumentBuilder`, `Aspose` | Retrieve display emphasismark value run debugging east asian formatting |
| [retrieve-font-object-paragraph-s-first-run-paragraph-ru...](./retrieve-font-object-paragraph-s-first-run-paragraph-runs-0-font.cs) | `Document`, `Section`, `Body` | Retrieve font object paragraph s first run paragraph runs 0 font |
| ... | | *and 5 more files* |

## Category Statistics
- Total examples: 35

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for fonts patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082357`
<!-- AUTOGENERATED:END -->
