---
name: fonts
description: Verified C# examples for fonts scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Fonts

## Purpose

Provide concise, reproducible code snippets that illustrate proper font manipulation in Aspose.Words, enabling AI agents to generate reliable examples that compile, run, and validate the intended visual changes.

## Non-negotiable conventions

- Do not import System.Drawing in this category; use Aspose.Drawing only when a drawing-specific color is required.
- Never reference bare Font or Color types without disambiguation when both Aspose.Words and Aspose.Drawing define them.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Document Formatting workflow**: 27 examples
- **Font Settings And Sources workflow**: 4 examples
- **Embedding And Rendering workflow**: 2 examples
- **Reporting And Inspection workflow**: 1 examples
- **Input Bootstrap workflow**: 1 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Always assert that the expected Font property (e.g., Underline, LineSpacingRule, EmphasisMark) matches the value set before saving the document.

## File-to-task reference

- `retrieve-the-font-object-from-a-paragraph-s-first-run-using-paragraph-runs-0-font.cs`
  - Task: Retrieve the Font object from a Paragraph's first Run using Paragraph.Runs[0].Font.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `use-documentbuilder-font-name-to-set-default-font-for-newly-inserted-content-before-adding.cs`
  - Task: Use DocumentBuilder.Font.Name to set default font for newly inserted content before adding text.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `use-documentbuilder-to-set-default-font-size-for-all-subsequently-inserted-paragraphs.cs`
  - Task: Use DocumentBuilder to set default font size for all subsequently inserted paragraphs.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `define-a-style-that-sets-font-name-size-and-color-then-apply-it-to-selected-text.cs`
  - Task: Define a style that sets font name, size, and color, then apply it to selected text.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-predefined-style-that-includes-specific-font-settings-to-a-range-of-paragraphs.cs`
  - Task: Apply a predefined style that includes specific font settings to a range of paragraphs.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `set-the-font-name-of-a-run-to-arial-via-the-font-name-property.cs`
  - Task: Set the font name of a Run to Arial via the Font.Name property.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `change-the-font-size-of-a-run-to-fourteen-points-using-the-font-size-property.cs`
  - Task: Change the font size of a Run to fourteen points using the Font.Size property.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `assign-a-blue-color-to-a-run-s-font-by-setting-font-color-to-system-drawing-color-blue.cs`
  - Task: Assign a blue color to a Run's font by setting Font.Color to System.Drawing.Color.Blue.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `apply-bold-and-italic-styles-to-a-run-by-setting-font-bold-and-font-italic-to-true.cs`
  - Task: Apply bold and italic styles to a Run by setting Font.Bold and Font.Italic to true.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `apply-underline-style-to-a-run-by-setting-font-underline-to-underlinetype-single.cs`
  - Task: Apply underline style to a Run by setting Font.Underline to UnderlineType.Single.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: existing_repo
- `apply-bold-italic-and-underline-simultaneously-to-a-run-using-respective-font-properties.cs`
  - Task: Apply bold, italic, and underline simultaneously to a Run using respective Font properties.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: existing_repo
- `set-text-fill-color-to-red-and-thirty-percent-transparency-using-font-fill-color-and-font.cs`
  - Task: Set text fill color to red and thirty percent transparency using Font.Fill.Color and Font.Fill.Transparency.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-semi-transparent-fill-to-text-by-setting-font-fill-color-and-font-fill-transparenc.cs`
  - Task: Apply a semi-transparent fill to text by setting Font.Fill.Color and Font.Fill.Transparency together.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `set-the-underline-type-to-double-wave-for-a-run-by-assigning-font-underline-to-underlinety.cs`
  - Task: Set the underline type to double wave for a Run by assigning Font.Underline to UnderlineType.DoubleWave.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: existing_repo
- `adjust-a-paragraph-s-line-spacing-to-1-5-lines-by-assigning-font-linespacing-to-1-5.cs`
  - Task: Adjust a paragraph's line spacing to 1.5 lines by assigning Font.LineSpacing to 1.5.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `set-line-spacing-to-double-for-a-paragraph-by-assigning-font-linespacing-to-2-0.cs`
  - Task: Set line spacing to double for a paragraph by assigning Font.LineSpacing to 2.0.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-reset-line-spacing-to-default-for-all-paragraphs-by-setting-font-linespac.cs`
  - Task: Programmatically reset line spacing to default for all paragraphs by setting Font.LineSpacing to 0.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `apply-east-asian-emphasis-mark-dot-to-a-run-by-setting-font-emphasismark-to-emphasismark-d.cs`
  - Task: Apply East Asian emphasis mark Dot to a Run by setting Font.EmphasisMark to EmphasisMark.Dot.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: existing_repo
- `apply-east-asian-emphasis-mark-only-when-document-language-is-japanese-by-checking-languag.cs`
  - Task: Apply East Asian emphasis mark only when document language is Japanese by checking language before setting Font.EmphasisMark.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: existing_repo
- `retrieve-and-display-the-emphasismark-value-of-a-run-for-debugging-east-asian-formatting.cs`
  - Task: Retrieve and display the EmphasisMark value of a Run for debugging East Asian formatting.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: llm
- `retrieve-and-display-the-current-font-size-of-a-run-for-debugging-purposes.cs`
  - Task: Retrieve and display the current font size of a Run for debugging purposes.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `define-a-custom-font-substitution-rule-that-replaces-missing-times-new-roman-with-calibri.cs`
  - Task: Define a custom font substitution rule that replaces missing Times New Roman with Calibri.
  - Workflow: font-settings-and-sources
  - Outputs: docx
  - Selected engine: llm
- `configure-font-substitution-to-map-missing-garamond-to-a-locally-installed-georgia-font.cs`
  - Task: Configure font substitution to map missing Garamond to a locally installed Georgia font.
  - Workflow: font-settings-and-sources
  - Outputs: docx
  - Selected engine: existing_repo
- `enable-automatic-font-substitution-for-missing-fonts-during-document-loading-by-configurin.cs`
  - Task: Enable automatic font substitution for missing fonts during document loading by configuring FontSettings.
  - Workflow: font-settings-and-sources
  - Outputs: docx
  - Selected engine: existing_repo
- `create-a-custom-font-substitution-table-and-load-it-from-an-xml-configuration-file.cs`
  - Task: Create a custom font substitution table and load it from an XML configuration file.
  - Workflow: font-settings-and-sources
  - Outputs: xml
  - Selected engine: existing_repo
- `embed-all-used-fonts-into-the-document-when-saving-by-enabling-the-embedfullfonts-option.cs`
  - Task: Embed all used fonts into the document when saving by enabling the EmbedFullFonts option.
  - Workflow: embedding-and-rendering
  - Outputs: docx
  - Selected engine: mcp
- `enable-font-embedding-when-saving-documents-to-pdf-to-ensure-visual-fidelity-across-platfo.cs`
  - Task: Enable font embedding when saving documents to PDF to ensure visual fidelity across platforms.
  - Workflow: embedding-and-rendering
  - Outputs: pdf
  - Selected engine: mcp
- `reset-all-font-attributes-of-a-run-to-defaults-by-calling-font-clearformatting.cs`
  - Task: Reset all font attributes of a Run to defaults by calling Font.ClearFormatting.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `copy-font-formatting-from-one-run-to-another-using-run-font-clone-method.cs`
  - Task: Copy font formatting from one Run to another using Run.Font.Clone method.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `create-a-reusable-method-that-applies-a-specific-font-name-and-size-to-any-given-run.cs`
  - Task: Create a reusable method that applies a specific font name and size to any given Run.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `create-a-font-object-set-its-size-and-color-then-assign-it-to-multiple-runs-for-consistenc.cs`
  - Task: Create a Font object, set its size and color, then assign it to multiple Runs for consistency.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-replace-all-occurrences-of-a-specific-font-with-another-across-the-entire.cs`
  - Task: Programmatically replace all occurrences of a specific font with another across the entire document.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `adjust-font-color-dynamically-based-on-paragraph-index-using-a-loop-and-conditional-logic.cs`
  - Task: Adjust font color dynamically based on paragraph index using a loop and conditional logic.
  - Workflow: reporting-and-inspection
  - Outputs: docx
  - Selected engine: mcp
- `implement-batch-processing-to-set-helvetica-font-for-all-runs-in-a-collection-of-documents.cs`
  - Task: Implement batch processing to set Helvetica font for all Runs in a collection of documents.
  - Workflow: input-bootstrap
  - Outputs: docx
  - Selected engine: mcp
- `load-a-document-change-all-headings-to-use-a-bold-sixteen-point-font-then-save.cs`
  - Task: Load a document, change all headings to use a bold sixteen-point font, then save.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Underline enum misuse**
  - Symptom: Compilation fails because UnderlineType is not defined in Aspose.Words.
  - Root cause: The code references a non-existent UnderlineType enum instead of the correct Aspose.Words.Underline enum.
  - Preferred fix: Replace UnderlineType with Aspose.Words.Underline and assign values such as Underline.Single or Underline.WavyDouble.
  - Prevention: Consult the Aspose.Words API reference for the exact enum name and use IntelliSense to select the correct member.

- **Line spacing property misuse**
  - Symptom: The compiler reports that Font.LineSpacing is read-only and cannot be assigned.
  - Root cause: Line spacing is a paragraph-level attribute, not a font attribute.
  - Preferred fix: Set ParagraphFormat.LineSpacingRule to Multiple and ParagraphFormat.LineSpacing to the desired multiplier (e.g., 1.5 or 2.0).
  - Prevention: Apply line-spacing changes through ParagraphFormat rather than Font, and reset to default by assigning 0 to ParagraphFormat.LineSpacing.

- **EmphasisMark enum mismatch**
  - Symptom: Compilation fails because EmphasisMark does not contain a member named Dot.
  - Root cause: The Aspose.Words.Fonts.EmphasisMark enum does not define a Dot value; the closest supported value is OverSolidCircle.
  - Preferred fix: Use EmphasisMark.OverSolidCircle (or another valid member) from Aspose.Words.Fonts.
  - Prevention: Reference the Aspose.Words.Fonts namespace and verify available enum members before coding.

- **Font substitution API misuse**
  - Symptom: Compilation fails due to missing members such as FontSubstitutionSettings.FontSubstitutes or DefaultFontName.
  - Root cause: The code uses outdated or non-existent properties of FontSettings; the current API exposes FontSettings.SubstitutionSettings.
  - Preferred fix: Configure substitution via FontSettings.SubstitutionSettings.DefaultFontName and SubstitutionSettings.AddSubstitutes.
  - Prevention: Review the latest Aspose.Words documentation for FontSettings and use the SubstitutionSettings property for custom mappings.

- **Color type ambiguity**
  - Symptom: Compilation errors arise from implicit conversion between Aspose.Drawing.Color and System.Drawing.Color.
  - Root cause: Font.Color and Font.Fill.Color expect System.Drawing.Color, while sample code creates Aspose.Drawing.Color objects.
  - Preferred fix: Convert Aspose.Drawing.Color to System.Drawing.Color using System.Drawing.Color.FromArgb(asposeColor.ToArgb()).
  - Prevention: Always create the final color value as System.Drawing.Color before assigning it to Font properties.

- **Namespace directive misuse**
  - Symptom: CS0138 errors indicate that a using directive targets a type rather than a namespace.
  - Root cause: The code attempts to write "using Font;" or "using Style;" instead of proper namespace imports.
  - Preferred fix: Remove type-level using statements; import only namespaces such as Aspose.Words, Aspose.Words.Fonts, or use fully qualified type names.
  - Prevention: Limit using directives to namespaces and rely on fully qualified names when type name collisions exist.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.10.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
dotnet add package Aspose.Drawing.Common --version 26.10.0
dotnet add package Newtonsoft.Json
```

```powershell
Copy-Item ..\fonts\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update examples when new Aspose.Words versions introduce renamed enums or additional font-related properties.
