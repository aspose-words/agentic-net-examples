---
name: fonts
description: Verified C# examples for fonts scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Fonts

## Purpose

Provide an AI coding assistant with clear guidance on applying font-related properties in Aspose.Words while avoiding common pitfalls such as type ambiguity, incorrect enum usage, and missing resource validation.

## Non-negotiable conventions

- Do not import System.Drawing in this category; use Aspose.Drawing only for true drawing scenarios.
- Never reference a bare Font or Color type when both Aspose.Drawing and Aspose.Words provide definitions.
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

1. Validate that each formatting change is reflected in the document model and that the output file exists before completing the example.

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
  - Selected engine: mcp
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
  - Selected engine: llm
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
  - Selected engine: mcp
- `retrieve-and-display-the-current-font-size-of-a-run-for-debugging-purposes.cs`
  - Task: Retrieve and display the current font size of a Run for debugging purposes.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: mcp
- `define-a-custom-font-substitution-rule-that-replaces-missing-times-new-roman-with-calibri.cs`
  - Task: Define a custom font substitution rule that replaces missing Times New Roman with Calibri.
  - Workflow: font-settings-and-sources
  - Outputs: docx
  - Selected engine: mcp
- `configure-font-substitution-to-map-missing-garamond-to-a-locally-installed-georgia-font.cs`
  - Task: Configure font substitution to map missing Garamond to a locally installed Georgia font.
  - Workflow: font-settings-and-sources
  - Outputs: docx
  - Selected engine: llm
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
  - Selected engine: llm
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

- **Font type ambiguity**
  - Symptom: CS0104 reports an ambiguous reference between Aspose.Drawing.Font and Aspose.Words.Font.
  - Root cause: Both namespaces are imported without disambiguation, causing the compiler to be unable to choose the correct Font type.
  - Preferred fix: Remove the conflicting using directive or qualify the type explicitly (e.g., Aspose.Words.Font).
  - Prevention: Always import only Aspose.Words for font operations and use fully qualified names when a conflict is possible.

- **Color type ambiguity**
  - Symptom: Compile errors arise because the Color identifier resolves to multiple types.
  - Root cause: Simultaneous inclusion of System.Drawing and Aspose.Drawing leads to ambiguous Color references.
  - Preferred fix: Use Aspose.Drawing.Color for color creation and convert to System.Drawing.Color only when required by an API.
  - Prevention: Declare the intended Color namespace explicitly and avoid generic using statements for System.Drawing.

- **Incorrect enum or property usage**
  - Symptom: CS0103, CS0234, CS0117, or CS0200 errors indicate missing members such as UnderlineType, EmphasisMark, or read-only Font.LineSpacing.
  - Root cause: The code references enums or properties that do not exist in the Aspose.Words version or are read-only.
  - Preferred fix: Use Aspose.Words.Underline for underline styles, Aspose.Words.EmphasisMark for East-Asian emphasis, and adjust line spacing via ParagraphFormat.LineSpacing and LineSpacingRule.
  - Prevention: Consult the Aspose.Words API reference for the correct enum names and mutable properties before coding.

- **Missing FontSettings API**
  - Symptom: CS0246 and CS1061 errors show that FontSettings, FontSubstitutionSettings, or LoadOptions cannot be found.
  - Root cause: The required types reside in the Aspose.Words.Fonts namespace and may require a newer package version.
  - Preferred fix: Add a reference to Aspose.Words.Fonts, include `using Aspose.Words.Fonts;`, and ensure the project references a version that contains FontSettings.
  - Prevention: Verify the availability of font-substitution classes in the target Aspose.Words version before implementing substitution logic.

- **NullReference at runtime**
  - Symptom: A NullReferenceException occurs when accessing a previous sibling or builder property.
  - Root cause: The code assumes that a paragraph or run exists without checking for null after document modifications.
  - Preferred fix: Retrieve the Run directly after insertion or test for null before dereferencing.
  - Prevention: Always validate that the target node is non-null before accessing its members.

- **Missing local bootstrap inputs and weak validation**
  - Symptom: Examples fail because they assume external files or folders exist and do not confirm that formatting changes took effect.
  - Root cause: The sample code does not create required resources or verify the document state before saving.
  - Preferred fix: Create temporary documents, apply formatting, and assert property values and file existence within the example.
  - Prevention: Include deterministic setup and explicit validation steps in every sample.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
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

- Update the examples whenever a new Aspose.Words version changes enum names or introduces new font-related APIs, and revise the strict API bans accordingly.
