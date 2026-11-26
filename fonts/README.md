# Fonts Examples for Aspose.Words for .NET

The Fonts category demonstrates how to manipulate text formatting in Aspose.Words documents using the Aspose.Words.Font API. Verified examples show correct style creation, underline application, combined bold/italic/underline formatting, line-spacing adjustments, and East-Asian emphasis handling.

## Snapshot

- Category: Fonts
- Slug: fonts
- Total examples: 35
- Publish-ready successful examples: 35 / 35
- Source run: 20261124_190740_1c00f5
- Document Formatting examples: 27
- Font Settings And Sources examples: 4
- Embedding And Rendering examples: 2
- Reporting And Inspection examples: 1
- Input Bootstrap examples: 1

## Category rules that shaped these examples

- Do not import System.Drawing in this category; use Aspose.Drawing only for true drawing scenarios.
- Never reference a bare Font or Color type when both Aspose.Drawing and Aspose.Words provide definitions.
- Apply font attributes through Aspose.Words.Font on Run, Style, ParagraphFormat, or DocumentBuilder objects.
- When a drawing-font type is required, qualify it as Aspose.Drawing.Font and convert colors via Aspose.Drawing.Color.ToArgb().
- Bootstrap sample documents and verify output files within each example.
- Check that objects such as Paragraph or Run are non-null before accessing their members.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.11.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
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
dotnet add package Aspose.Words --version 26.11.0

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
| 11 | `apply-bold-italic-and-underline-simultaneously-to-a-run-using-respective-font-properties.cs` | Apply bold, italic, and underline simultaneously to a Run using respective Font properties. | document-formatting | docx | mcp |
| 12 | `set-text-fill-color-to-red-and-thirty-percent-transparency-using-font-fill-color-and-font.cs` | Set text fill color to red and thirty percent transparency using Font.Fill.Color and Font.Fill.Transparency. | document-formatting | docx | mcp |
| 13 | `apply-a-semi-transparent-fill-to-text-by-setting-font-fill-color-and-font-fill-transparenc.cs` | Apply a semi-transparent fill to text by setting Font.Fill.Color and Font.Fill.Transparency together. | document-formatting | docx | mcp |
| 14 | `set-the-underline-type-to-double-wave-for-a-run-by-assigning-font-underline-to-underlinety.cs` | Set the underline type to double wave for a Run by assigning Font.Underline to UnderlineType.DoubleWave. | document-formatting | docx | llm |
| 15 | `adjust-a-paragraph-s-line-spacing-to-1-5-lines-by-assigning-font-linespacing-to-1-5.cs` | Adjust a paragraph's line spacing to 1.5 lines by assigning Font.LineSpacing to 1.5. | document-formatting | docx | mcp |
| 16 | `set-line-spacing-to-double-for-a-paragraph-by-assigning-font-linespacing-to-2-0.cs` | Set line spacing to double for a paragraph by assigning Font.LineSpacing to 2.0. | document-formatting | docx | mcp |
| 17 | `programmatically-reset-line-spacing-to-default-for-all-paragraphs-by-setting-font-linespac.cs` | Programmatically reset line spacing to default for all paragraphs by setting Font.LineSpacing to 0. | document-formatting | docx | mcp |
| 18 | `apply-east-asian-emphasis-mark-dot-to-a-run-by-setting-font-emphasismark-to-emphasismark-d.cs` | Apply East Asian emphasis mark Dot to a Run by setting Font.EmphasisMark to EmphasisMark.Dot. | document-formatting | docx | existing_repo |
| 19 | `apply-east-asian-emphasis-mark-only-when-document-language-is-japanese-by-checking-languag.cs` | Apply East Asian emphasis mark only when document language is Japanese by checking language before setting Font.EmphasisMark. | document-formatting | docx | existing_repo |
| 20 | `retrieve-and-display-the-emphasismark-value-of-a-run-for-debugging-east-asian-formatting.cs` | Retrieve and display the EmphasisMark value of a Run for debugging East Asian formatting. | document-formatting | docx | mcp |
| 21 | `retrieve-and-display-the-current-font-size-of-a-run-for-debugging-purposes.cs` | Retrieve and display the current font size of a Run for debugging purposes. | document-formatting | docx | mcp |
| 22 | `define-a-custom-font-substitution-rule-that-replaces-missing-times-new-roman-with-calibri.cs` | Define a custom font substitution rule that replaces missing Times New Roman with Calibri. | font-settings-and-sources | docx | mcp |
| 23 | `configure-font-substitution-to-map-missing-garamond-to-a-locally-installed-georgia-font.cs` | Configure font substitution to map missing Garamond to a locally installed Georgia font. | font-settings-and-sources | docx | llm |
| 24 | `enable-automatic-font-substitution-for-missing-fonts-during-document-loading-by-configurin.cs` | Enable automatic font substitution for missing fonts during document loading by configuring FontSettings. | font-settings-and-sources | docx | existing_repo |
| 25 | `create-a-custom-font-substitution-table-and-load-it-from-an-xml-configuration-file.cs` | Create a custom font substitution table and load it from an XML configuration file. | font-settings-and-sources | xml | existing_repo |
| 26 | `embed-all-used-fonts-into-the-document-when-saving-by-enabling-the-embedfullfonts-option.cs` | Embed all used fonts into the document when saving by enabling the EmbedFullFonts option. | embedding-and-rendering | docx | llm |
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

### Font type ambiguity

- Symptom: CS0104 reports an ambiguous reference between Aspose.Drawing.Font and Aspose.Words.Font.
- Root cause: Both namespaces are imported without disambiguation, causing the compiler to be unable to choose the correct Font type.
- Preferred fix: Remove the conflicting using directive or qualify the type explicitly (e.g., Aspose.Words.Font).
- Prevention: Always import only Aspose.Words for font operations and use fully qualified names when a conflict is possible.

### Color type ambiguity

- Symptom: Compile errors arise because the Color identifier resolves to multiple types.
- Root cause: Simultaneous inclusion of System.Drawing and Aspose.Drawing leads to ambiguous Color references.
- Preferred fix: Use Aspose.Drawing.Color for color creation and convert to System.Drawing.Color only when required by an API.
- Prevention: Declare the intended Color namespace explicitly and avoid generic using statements for System.Drawing.

### Incorrect enum or property usage

- Symptom: CS0103, CS0234, CS0117, or CS0200 errors indicate missing members such as UnderlineType, EmphasisMark, or read-only Font.LineSpacing.
- Root cause: The code references enums or properties that do not exist in the Aspose.Words version or are read-only.
- Preferred fix: Use Aspose.Words.Underline for underline styles, Aspose.Words.EmphasisMark for East-Asian emphasis, and adjust line spacing via ParagraphFormat.LineSpacing and LineSpacingRule.
- Prevention: Consult the Aspose.Words API reference for the correct enum names and mutable properties before coding.

### Missing FontSettings API

- Symptom: CS0246 and CS1061 errors show that FontSettings, FontSubstitutionSettings, or LoadOptions cannot be found.
- Root cause: The required types reside in the Aspose.Words.Fonts namespace and may require a newer package version.
- Preferred fix: Add a reference to Aspose.Words.Fonts, include `using Aspose.Words.Fonts;`, and ensure the project references a version that contains FontSettings.
- Prevention: Verify the availability of font-substitution classes in the target Aspose.Words version before implementing substitution logic.

### NullReference at runtime

- Symptom: A NullReferenceException occurs when accessing a previous sibling or builder property.
- Root cause: The code assumes that a paragraph or run exists without checking for null after document modifications.
- Preferred fix: Retrieve the Run directly after insertion or test for null before dereferencing.
- Prevention: Always validate that the target node is non-null before accessing its members.

### Missing local bootstrap inputs and weak validation

- Symptom: Examples fail because they assume external files or folders exist and do not confirm that formatting changes took effect.
- Root cause: The sample code does not create required resources or verify the document state before saving.
- Preferred fix: Create temporary documents, apply formatting, and assert property values and file existence within the example.
- Prevention: Include deterministic setup and explicit validation steps in every sample.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Update the examples whenever a new Aspose.Words version changes enum names or introduces new font-related APIs, and revise the strict API bans accordingly.
