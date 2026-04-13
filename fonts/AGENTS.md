---
name: fonts
description: Verified C# examples for font scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Fonts

## Purpose

This folder is a **live, curated example set** for font scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free document font formatting, font settings, substitution, embedding, fallback, rendering, and font-report workflows using direct Aspose.Words APIs.

## Non-negotiable conventions

- Use native Aspose.Words font APIs directly.
- Use Aspose.Words.Font for document text formatting.
- Use Aspose.Drawing.Font only for true drawing scenarios and with a fully qualified type name.
- Do not use using System.Drawing.
- Do not use bare Font or bare Color when ambiguity is possible.
- Bootstrap local source documents, files, streams, or folders whenever the task implies an existing source.
- Guard maybe-null values to avoid nullable-reference warnings such as CS8600, CS8602, and CS8604.

## Recommended workflow selection

- **Document formatting workflow**: 27 examples
- **Font settings / sources workflow**: 4 examples
- **Embedding / rendering workflow**: 2 examples
- **Reporting / inspection workflow**: 1 examples
- **Input-bootstrap workflow**: 1 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. Source documents, files, streams, or folders must be bootstrapped locally whenever the task implies an existing input.
3. Font changes, reports, embedding, substitution, fallback, or rendering results must be validated when the task expects them.
4. Requested report or export files must actually be written.
5. Font and color type usage must be explicit when ambiguity is possible.

## File-to-task reference

- `retrieve-the-font-object-from-a-paragraph-s-first-run-using-paragraph-runs-0-font.cs`
  - Task: Retrieve the Font object from a Paragraph's first Run using Paragraph.Runs[0].Font.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `use-documentbuilder-font-name-to-set-default-font-for-newly-inserted-content-before-adding.cs`
  - Task: Use DocumentBuilder.Font.Name to set default font for newly inserted content before adding text.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `use-documentbuilder-to-set-default-font-size-for-all-subsequently-inserted-paragraphs.cs`
  - Task: Use DocumentBuilder to set default font size for all subsequently inserted paragraphs.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `define-a-style-that-sets-font-name-size-and-color-then-apply-it-to-selected-text.cs`
  - Task: Define a style that sets font name, size, and color, then apply it to selected text.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `apply-a-predefined-style-that-includes-specific-font-settings-to-a-range-of-paragraphs.cs`
  - Task: Apply a predefined style that includes specific font settings to a range of paragraphs.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `set-the-font-name-of-a-run-to-arial-via-the-font-name-property.cs`
  - Task: Set the font name of a Run to Arial via the Font.Name property.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `change-the-font-size-of-a-run-to-fourteen-points-using-the-font-size-property.cs`
  - Task: Change the font size of a Run to fourteen points using the Font.Size property.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `assign-a-blue-color-to-a-run-s-font-by-setting-font-color-to-system-drawing-color-blue.cs`
  - Task: Assign a blue color to a Run's font by setting Font.Color to System.Drawing.Color.Blue.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `apply-bold-and-italic-styles-to-a-run-by-setting-font-bold-and-font-italic-to-true.cs`
  - Task: Apply bold and italic styles to a Run by setting Font.Bold and Font.Italic to true.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `apply-underline-style-to-a-run-by-setting-font-underline-to-underlinetype-single.cs`
  - Task: Apply underline style to a Run by setting Font.Underline to UnderlineType.Single.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `apply-bold-italic-and-underline-simultaneously-to-a-run-using-respective-font-properties.cs`
  - Task: Apply bold, italic, and underline simultaneously to a Run using respective Font properties.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `set-text-fill-color-to-red-and-thirty-percent-transparency-using-font-fill-color-and-font.cs`
  - Task: Set text fill color to red and thirty percent transparency using Font.Fill.Color and Font.Fill.Transparency.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `apply-a-semi-transparent-fill-to-text-by-setting-font-fill-color-and-font-fill-transparenc.cs`
  - Task: Apply a semi‑transparent fill to text by setting Font.Fill.Color and Font.Fill.Transparency together.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `set-the-underline-type-to-double-wave-for-a-run-by-assigning-font-underline-to-underlinety.cs`
  - Task: Set the underline type to double wave for a Run by assigning Font.Underline to UnderlineType.DoubleWave.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `adjust-a-paragraph-s-line-spacing-to-1-5-lines-by-assigning-font-linespacing-to-1-5.cs`
  - Task: Adjust a paragraph's line spacing to 1.5 lines by assigning Font.LineSpacing to 1.5.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `set-line-spacing-to-double-for-a-paragraph-by-assigning-font-linespacing-to-2-0.cs`
  - Task: Set line spacing to double for a paragraph by assigning Font.LineSpacing to 2.0.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `programmatically-reset-line-spacing-to-default-for-all-paragraphs-by-setting-font-linespac.cs`
  - Task: Programmatically reset line spacing to default for all paragraphs by setting Font.LineSpacing to 0.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `apply-east-asian-emphasis-mark-dot-to-a-run-by-setting-font-emphasismark-to-emphasismark-d.cs`
  - Task: Apply East Asian emphasis mark Dot to a Run by setting Font.EmphasisMark to EmphasisMark.Dot.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `apply-east-asian-emphasis-mark-only-when-document-language-is-japanese-by-checking-languag.cs`
  - Task: Apply East Asian emphasis mark only when document language is Japanese by checking language before setting Font.EmphasisMark.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `retrieve-and-display-the-emphasismark-value-of-a-run-for-debugging-east-asian-formatting.cs`
  - Task: Retrieve and display the EmphasisMark value of a Run for debugging East Asian formatting.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `retrieve-and-display-the-current-font-size-of-a-run-for-debugging-purposes.cs`
  - Task: Retrieve and display the current font size of a Run for debugging purposes.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `define-a-custom-font-substitution-rule-that-replaces-missing-times-new-roman-with-calibri.cs`
  - Task: Define a custom font substitution rule that replaces missing Times New Roman with Calibri.
  - Workflow: font-settings-and-sources
  - Outputs: docx
  - Selected engine: verified
- `configure-font-substitution-to-map-missing-garamond-to-a-locally-installed-georgia-font.cs`
  - Task: Configure font substitution to map missing Garamond to a locally installed Georgia font.
  - Workflow: font-settings-and-sources
  - Outputs: docx
  - Selected engine: verified
- `enable-automatic-font-substitution-for-missing-fonts-during-document-loading-by-configurin.cs`
  - Task: Enable automatic font substitution for missing fonts during document loading by configuring FontSettings.
  - Workflow: font-settings-and-sources
  - Outputs: docx
  - Selected engine: verified
- `create-a-custom-font-substitution-table-and-load-it-from-an-xml-configuration-file.cs`
  - Task: Create a custom font substitution table and load it from an XML configuration file.
  - Workflow: font-settings-and-sources
  - Outputs: xml
  - Selected engine: verified
- `embed-all-used-fonts-into-the-document-when-saving-by-enabling-the-embedfullfonts-option.cs`
  - Task: Embed all used fonts into the document when saving by enabling the EmbedFullFonts option.
  - Workflow: embedding-and-rendering
  - Outputs: docx
  - Selected engine: verified
- `enable-font-embedding-when-saving-documents-to-pdf-to-ensure-visual-fidelity-across-platfo.cs`
  - Task: Enable font embedding when saving documents to PDF to ensure visual fidelity across platforms.
  - Workflow: embedding-and-rendering
  - Outputs: pdf
  - Selected engine: verified
- `reset-all-font-attributes-of-a-run-to-defaults-by-calling-font-clearformatting.cs`
  - Task: Reset all font attributes of a Run to defaults by calling Font.ClearFormatting.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `copy-font-formatting-from-one-run-to-another-using-run-font-clone-method.cs`
  - Task: Copy font formatting from one Run to another using Run.Font.Clone method.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `create-a-reusable-method-that-applies-a-specific-font-name-and-size-to-any-given-run.cs`
  - Task: Create a reusable method that applies a specific font name and size to any given Run.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `create-a-font-object-set-its-size-and-color-then-assign-it-to-multiple-runs-for-consistenc.cs`
  - Task: Create a Font object, set its size and color, then assign it to multiple Runs for consistency.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `programmatically-replace-all-occurrences-of-a-specific-font-with-another-across-the-entire.cs`
  - Task: Programmatically replace all occurrences of a specific font with another across the entire document.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified
- `adjust-font-color-dynamically-based-on-paragraph-index-using-a-loop-and-conditional-logic.cs`
  - Task: Adjust font color dynamically based on paragraph index using a loop and conditional logic.
  - Workflow: reporting-and-inspection
  - Outputs: docx
  - Selected engine: verified
- `implement-batch-processing-to-set-helvetica-font-for-all-runs-in-a-collection-of-documents.cs`
  - Task: Implement batch processing to set Helvetica font for all Runs in a collection of documents.
  - Workflow: input-bootstrap
  - Outputs: docx
  - Selected engine: verified
- `load-a-document-change-all-headings-to-use-a-bold-sixteen-point-font-then-save.cs`
  - Task: Load a document, change all headings to use a bold sixteen‑point font, then save.
  - Workflow: document-formatting
  - Outputs: docx
  - Selected engine: verified

## Common failure patterns and preferred agent fixes

- **Font type ambiguity**
  - Symptom: Compile errors because Font is ambiguous between Aspose.Drawing.Font and Aspose.Words.Font.
  - Preferred fix: Use fully qualified names and use Aspose.Words.Font for document formatting versus Aspose.Drawing.Font only for real drawing scenarios.

- **Color type ambiguity**
  - Symptom: Compile or assignment issues because Color is ambiguous or the wrong color type is used for the target property.
  - Preferred fix: Use fully qualified Aspose.Drawing.Color or explicit System.Drawing.Color conversion only where the Aspose.Words API expects it.

- **Wrong formatting object**
  - Symptom: A drawing font object is treated as if it were a document text-formatting object.
  - Preferred fix: Apply document font values through Run.Font, Style.Font, ParagraphBreakFont, or DocumentBuilder.Font.

- **Missing local bootstrap inputs**
  - Symptom: Examples assume existing documents, folders, or font resources already exist.
  - Preferred fix: Create local sample documents and deterministic local scenarios inside the example.

- **Weak validation**
  - Symptom: Examples save output but do not verify that the intended font change, report, or render actually occurred.
  - Preferred fix: Validate the changed font property, configured font setting, or output file existence before finishing.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`
- Additional package: `Aspose.Drawing.Common`
- Additional package: `Newtonsoft.Json`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.Drawing.Common
dotnet add package Newtonsoft.Json
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\fonts\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer direct Aspose.Words font APIs over speculative shortcuts.
