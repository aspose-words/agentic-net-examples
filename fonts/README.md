# Fonts Examples for Aspose.Words for .NET

The Fonts category demonstrates how to manipulate text appearance in Aspose.Words documents using the library's Font, Style, and DocumentBuilder APIs. Verified examples show correct handling of font properties, underline styles, line spacing, emphasis marks, and color conversion while avoiding common ambiguities.

## Snapshot

- Category: Fonts
- Slug: fonts
- Total examples: 35
- Publish-ready successful examples: 35 / 35
- Source run: 20261102_031946_08fb5b
- Document Formatting examples: 27
- Font Settings And Sources examples: 4
- Embedding And Rendering examples: 2
- Reporting And Inspection examples: 1
- Input Bootstrap examples: 1

## Category rules that shaped these examples

- Do not import System.Drawing in this category; use Aspose.Drawing only when a drawing-specific color is required.
- Never reference bare Font or Color types without disambiguation when both Aspose.Words and Aspose.Drawing define them.
- Apply text formatting through Run.Font, Style.Font, ParagraphFormat, or DocumentBuilder.Font, never through Aspose.Drawing.Font.
- Convert Aspose.Drawing.Color to System.Drawing.Color before assigning it to Font.Color or Font.Fill.Color.
- Create local sample documents, apply formatting, and validate the properties before persisting the file.
- Validate that objects such as Document, Paragraph, and Run are instantiated before accessing their Font members to avoid null-reference warnings.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.10.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
dotnet add package Aspose.Drawing.Common --version 26.10.0
dotnet add package Newtonsoft.Json

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\fonts\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `fonts/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0

# PowerShell example
Copy-Item ..\fonts\retrieve-the-font-object-from-a-paragraph-s-first-run-using-paragraph-runs-0-font.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `retrieve-the-font-object-from-a-paragraph-s-first-run-using-paragraph-runs-0-font.cs` | Retrieve the Font object from a Paragraph's first Run using Paragraph.Runs[0].Font. | document-formatting | docx | mcp |
| 2 | `use-documentbuilder-font-name-to-set-default-font-for-newly-inserted-content-before-adding.cs` | Use DocumentBuilder.Font.Name to set default font for newly inserted content before adding text. | document-formatting | docx | mcp |
| 3 | `use-documentbuilder-to-set-default-font-size-for-all-subsequently-inserted-paragraphs.cs` | Use DocumentBuilder to set default font size for all subsequently inserted paragraphs. | document-formatting | docx | mcp |
| 4 | `define-a-style-that-sets-font-name-size-and-color-then-apply-it-to-selected-text.cs` | Define a style that sets font name, size, and color, then apply it to selected text. | document-formatting | docx | mcp |
| 5 | `apply-a-predefined-style-that-includes-specific-font-settings-to-a-range-of-paragraphs.cs` | Apply a predefined style that includes specific font settings to a range of paragraphs. | document-formatting | docx | mcp |
| 6 | `set-the-font-name-of-a-run-to-arial-via-the-font-name-property.cs` | Set the font name of a Run to Arial via the Font.Name property. | document-formatting | docx | mcp |
| 7 | `change-the-font-size-of-a-run-to-fourteen-points-using-the-font-size-property.cs` | Change the font size of a Run to fourteen points using the Font.Size property. | document-formatting | docx | mcp |
| 8 | `assign-a-blue-color-to-a-run-s-font-by-setting-font-color-to-system-drawing-color-blue.cs` | Assign a blue color to a Run's font by setting Font.Color to System.Drawing.Color.Blue. | document-formatting | docx | mcp |
| 9 | `apply-bold-and-italic-styles-to-a-run-by-setting-font-bold-and-font-italic-to-true.cs` | Apply bold and italic styles to a Run by setting Font.Bold and Font.Italic to true. | document-formatting | docx | mcp |
| 10 | `apply-underline-style-to-a-run-by-setting-font-underline-to-underlinetype-single.cs` | Apply underline style to a Run by setting Font.Underline to UnderlineType.Single. | document-formatting | docx | existing_repo |
| 11 | `apply-bold-italic-and-underline-simultaneously-to-a-run-using-respective-font-properties.cs` | Apply bold, italic, and underline simultaneously to a Run using respective Font properties. | document-formatting | docx | existing_repo |
| 12 | `set-text-fill-color-to-red-and-thirty-percent-transparency-using-font-fill-color-and-font.cs` | Set text fill color to red and thirty percent transparency using Font.Fill.Color and Font.Fill.Transparency. | document-formatting | docx | mcp |
| 13 | `apply-a-semi-transparent-fill-to-text-by-setting-font-fill-color-and-font-fill-transparenc.cs` | Apply a semi-transparent fill to text by setting Font.Fill.Color and Font.Fill.Transparency together. | document-formatting | docx | mcp |
| 14 | `set-the-underline-type-to-double-wave-for-a-run-by-assigning-font-underline-to-underlinety.cs` | Set the underline type to double wave for a Run by assigning Font.Underline to UnderlineType.DoubleWave. | document-formatting | docx | existing_repo |
| 15 | `adjust-a-paragraph-s-line-spacing-to-1-5-lines-by-assigning-font-linespacing-to-1-5.cs` | Adjust a paragraph's line spacing to 1.5 lines by assigning Font.LineSpacing to 1.5. | document-formatting | docx | mcp |
| 16 | `set-line-spacing-to-double-for-a-paragraph-by-assigning-font-linespacing-to-2-0.cs` | Set line spacing to double for a paragraph by assigning Font.LineSpacing to 2.0. | document-formatting | docx | mcp |
| 17 | `programmatically-reset-line-spacing-to-default-for-all-paragraphs-by-setting-font-linespac.cs` | Programmatically reset line spacing to default for all paragraphs by setting Font.LineSpacing to 0. | document-formatting | docx | mcp |
| 18 | `apply-east-asian-emphasis-mark-dot-to-a-run-by-setting-font-emphasismark-to-emphasismark-d.cs` | Apply East Asian emphasis mark Dot to a Run by setting Font.EmphasisMark to EmphasisMark.Dot. | document-formatting | docx | existing_repo |
| 19 | `apply-east-asian-emphasis-mark-only-when-document-language-is-japanese-by-checking-languag.cs` | Apply East Asian emphasis mark only when document language is Japanese by checking language before setting Font.EmphasisMark. | document-formatting | docx | existing_repo |
| 20 | `retrieve-and-display-the-emphasismark-value-of-a-run-for-debugging-east-asian-formatting.cs` | Retrieve and display the EmphasisMark value of a Run for debugging East Asian formatting. | document-formatting | docx | llm |
| 21 | `retrieve-and-display-the-current-font-size-of-a-run-for-debugging-purposes.cs` | Retrieve and display the current font size of a Run for debugging purposes. | document-formatting | docx | mcp |
| 22 | `define-a-custom-font-substitution-rule-that-replaces-missing-times-new-roman-with-calibri.cs` | Define a custom font substitution rule that replaces missing Times New Roman with Calibri. | font-settings-and-sources | docx | llm |
| 23 | `configure-font-substitution-to-map-missing-garamond-to-a-locally-installed-georgia-font.cs` | Configure font substitution to map missing Garamond to a locally installed Georgia font. | font-settings-and-sources | docx | existing_repo |
| 24 | `enable-automatic-font-substitution-for-missing-fonts-during-document-loading-by-configurin.cs` | Enable automatic font substitution for missing fonts during document loading by configuring FontSettings. | font-settings-and-sources | docx | existing_repo |
| 25 | `create-a-custom-font-substitution-table-and-load-it-from-an-xml-configuration-file.cs` | Create a custom font substitution table and load it from an XML configuration file. | font-settings-and-sources | xml | existing_repo |
| 26 | `embed-all-used-fonts-into-the-document-when-saving-by-enabling-the-embedfullfonts-option.cs` | Embed all used fonts into the document when saving by enabling the EmbedFullFonts option. | embedding-and-rendering | docx | mcp |
| 27 | `enable-font-embedding-when-saving-documents-to-pdf-to-ensure-visual-fidelity-across-platfo.cs` | Enable font embedding when saving documents to PDF to ensure visual fidelity across platforms. | embedding-and-rendering | pdf | mcp |
| 28 | `reset-all-font-attributes-of-a-run-to-defaults-by-calling-font-clearformatting.cs` | Reset all font attributes of a Run to defaults by calling Font.ClearFormatting. | document-formatting | docx | mcp |
| 29 | `copy-font-formatting-from-one-run-to-another-using-run-font-clone-method.cs` | Copy font formatting from one Run to another using Run.Font.Clone method. | document-formatting | docx | mcp |
| 30 | `create-a-reusable-method-that-applies-a-specific-font-name-and-size-to-any-given-run.cs` | Create a reusable method that applies a specific font name and size to any given Run. | document-formatting | docx | mcp |
| 31 | `create-a-font-object-set-its-size-and-color-then-assign-it-to-multiple-runs-for-consistenc.cs` | Create a Font object, set its size and color, then assign it to multiple Runs for consistency. | document-formatting | docx | mcp |
| 32 | `programmatically-replace-all-occurrences-of-a-specific-font-with-another-across-the-entire.cs` | Programmatically replace all occurrences of a specific font with another across the entire document. | document-formatting | docx | mcp |
| 33 | `adjust-font-color-dynamically-based-on-paragraph-index-using-a-loop-and-conditional-logic.cs` | Adjust font color dynamically based on paragraph index using a loop and conditional logic. | reporting-and-inspection | docx | mcp |
| 34 | `implement-batch-processing-to-set-helvetica-font-for-all-runs-in-a-collection-of-documents.cs` | Implement batch processing to set Helvetica font for all Runs in a collection of documents. | input-bootstrap | docx | mcp |
| 35 | `load-a-document-change-all-headings-to-use-a-bold-sixteen-point-font-then-save.cs` | Load a document, change all headings to use a bold sixteen-point font, then save. | document-formatting | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Underline enum misuse

- Symptom: Compilation fails because UnderlineType is not defined in Aspose.Words.
- Root cause: The code references a non-existent UnderlineType enum instead of the correct Aspose.Words.Underline enum.
- Preferred fix: Replace UnderlineType with Aspose.Words.Underline and assign values such as Underline.Single or Underline.WavyDouble.
- Prevention: Consult the Aspose.Words API reference for the exact enum name and use IntelliSense to select the correct member.

### Line spacing property misuse

- Symptom: The compiler reports that Font.LineSpacing is read-only and cannot be assigned.
- Root cause: Line spacing is a paragraph-level attribute, not a font attribute.
- Preferred fix: Set ParagraphFormat.LineSpacingRule to Multiple and ParagraphFormat.LineSpacing to the desired multiplier (e.g., 1.5 or 2.0).
- Prevention: Apply line-spacing changes through ParagraphFormat rather than Font, and reset to default by assigning 0 to ParagraphFormat.LineSpacing.

### EmphasisMark enum mismatch

- Symptom: Compilation fails because EmphasisMark does not contain a member named Dot.
- Root cause: The Aspose.Words.Fonts.EmphasisMark enum does not define a Dot value; the closest supported value is OverSolidCircle.
- Preferred fix: Use EmphasisMark.OverSolidCircle (or another valid member) from Aspose.Words.Fonts.
- Prevention: Reference the Aspose.Words.Fonts namespace and verify available enum members before coding.

### Font substitution API misuse

- Symptom: Compilation fails due to missing members such as FontSubstitutionSettings.FontSubstitutes or DefaultFontName.
- Root cause: The code uses outdated or non-existent properties of FontSettings; the current API exposes FontSettings.SubstitutionSettings.
- Preferred fix: Configure substitution via FontSettings.SubstitutionSettings.DefaultFontName and SubstitutionSettings.AddSubstitutes.
- Prevention: Review the latest Aspose.Words documentation for FontSettings and use the SubstitutionSettings property for custom mappings.

### Color type ambiguity

- Symptom: Compilation errors arise from implicit conversion between Aspose.Drawing.Color and System.Drawing.Color.
- Root cause: Font.Color and Font.Fill.Color expect System.Drawing.Color, while sample code creates Aspose.Drawing.Color objects.
- Preferred fix: Convert Aspose.Drawing.Color to System.Drawing.Color using System.Drawing.Color.FromArgb(asposeColor.ToArgb()).
- Prevention: Always create the final color value as System.Drawing.Color before assigning it to Font properties.

### Namespace directive misuse

- Symptom: CS0138 errors indicate that a using directive targets a type rather than a namespace.
- Root cause: The code attempts to write "using Font;" or "using Style;" instead of proper namespace imports.
- Preferred fix: Remove type-level using statements; import only namespaces such as Aspose.Words, Aspose.Words.Fonts, or use fully qualified type names.
- Prevention: Limit using directives to namespaces and rely on fully qualified names when type name collisions exist.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Update examples when new Aspose.Words versions introduce renamed enums or additional font-related properties.
