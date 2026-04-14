# Fonts Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **Fonts** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Fonts**
- Slug: **fonts**
- Total examples: **35**
- Publish-ready successful examples: **35 / 35**
- Document formatting examples: **27**
- Font settings / sources examples: **4**
- Embedding / rendering examples: **2**
- Reporting / inspection examples: **1**
- Input-bootstrap examples: **1**

## Category rules that shaped these examples

- Use native Aspose.Words font APIs directly.
- Use Aspose.Words.Font for document text formatting and Aspose.Drawing.Font only for true drawing scenarios.
- Do not use using System.Drawing, and do not use bare Font or bare Color when ambiguity is possible.
- Create realistic local sample inputs whenever the task mentions an existing DOC, DOCX, HTML, TXT, stream, folder, or font-related source.
- Validate that the intended font change, report, embedding, substitution, fallback, or rendered output actually occurred.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`
- Aspose.Drawing.Common
- Newtonsoft.Json

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.Drawing.Common
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
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Aspose.Drawing.Common
dotnet add package Newtonsoft.Json

# PowerShell example
Copy-Item ..\fonts\retrieve-the-font-object-from-a-paragraph-s-first-run-using-paragraph-runs-0-font.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `retrieve-the-font-object-from-a-paragraph-s-first-run-using-paragraph-runs-0-font.cs` | Retrieve the Font object from a Paragraph's first Run using Paragraph.Runs[0].Font. | document-formatting | docx | verified |
| 2 | `use-documentbuilder-font-name-to-set-default-font-for-newly-inserted-content-before-adding.cs` | Use DocumentBuilder.Font.Name to set default font for newly inserted content before adding text. | document-formatting | docx | verified |
| 3 | `use-documentbuilder-to-set-default-font-size-for-all-subsequently-inserted-paragraphs.cs` | Use DocumentBuilder to set default font size for all subsequently inserted paragraphs. | document-formatting | docx | verified |
| 4 | `define-a-style-that-sets-font-name-size-and-color-then-apply-it-to-selected-text.cs` | Define a style that sets font name, size, and color, then apply it to selected text. | document-formatting | docx | verified |
| 5 | `apply-a-predefined-style-that-includes-specific-font-settings-to-a-range-of-paragraphs.cs` | Apply a predefined style that includes specific font settings to a range of paragraphs. | document-formatting | docx | verified |
| 6 | `set-the-font-name-of-a-run-to-arial-via-the-font-name-property.cs` | Set the font name of a Run to Arial via the Font.Name property. | document-formatting | docx | verified |
| 7 | `change-the-font-size-of-a-run-to-fourteen-points-using-the-font-size-property.cs` | Change the font size of a Run to fourteen points using the Font.Size property. | document-formatting | docx | verified |
| 8 | `assign-a-blue-color-to-a-run-s-font-by-setting-font-color-to-system-drawing-color-blue.cs` | Assign a blue color to a Run's font by setting Font.Color to System.Drawing.Color.Blue. | document-formatting | docx | verified |
| 9 | `apply-bold-and-italic-styles-to-a-run-by-setting-font-bold-and-font-italic-to-true.cs` | Apply bold and italic styles to a Run by setting Font.Bold and Font.Italic to true. | document-formatting | docx | verified |
| 10 | `apply-underline-style-to-a-run-by-setting-font-underline-to-underlinetype-single.cs` | Apply underline style to a Run by setting Font.Underline to UnderlineType.Single. | document-formatting | docx | verified |
| 11 | `apply-bold-italic-and-underline-simultaneously-to-a-run-using-respective-font-properties.cs` | Apply bold, italic, and underline simultaneously to a Run using respective Font properties. | document-formatting | docx | verified |
| 12 | `set-text-fill-color-to-red-and-thirty-percent-transparency-using-font-fill-color-and-font.cs` | Set text fill color to red and thirty percent transparency using Font.Fill.Color and Font.Fill.Transparency. | document-formatting | docx | verified |
| 13 | `apply-a-semi-transparent-fill-to-text-by-setting-font-fill-color-and-font-fill-transparenc.cs` | Apply a semi‑transparent fill to text by setting Font.Fill.Color and Font.Fill.Transparency together. | document-formatting | docx | verified |
| 14 | `set-the-underline-type-to-double-wave-for-a-run-by-assigning-font-underline-to-underlinety.cs` | Set the underline type to double wave for a Run by assigning Font.Underline to UnderlineType.DoubleWave. | document-formatting | docx | verified |
| 15 | `adjust-a-paragraph-s-line-spacing-to-1-5-lines-by-assigning-font-linespacing-to-1-5.cs` | Adjust a paragraph's line spacing to 1.5 lines by assigning Font.LineSpacing to 1.5. | document-formatting | docx | verified |
| 16 | `set-line-spacing-to-double-for-a-paragraph-by-assigning-font-linespacing-to-2-0.cs` | Set line spacing to double for a paragraph by assigning Font.LineSpacing to 2.0. | document-formatting | docx | verified |
| 17 | `programmatically-reset-line-spacing-to-default-for-all-paragraphs-by-setting-font-linespac.cs` | Programmatically reset line spacing to default for all paragraphs by setting Font.LineSpacing to 0. | document-formatting | docx | verified |
| 18 | `apply-east-asian-emphasis-mark-dot-to-a-run-by-setting-font-emphasismark-to-emphasismark-d.cs` | Apply East Asian emphasis mark Dot to a Run by setting Font.EmphasisMark to EmphasisMark.Dot. | document-formatting | docx | verified |
| 19 | `apply-east-asian-emphasis-mark-only-when-document-language-is-japanese-by-checking-languag.cs` | Apply East Asian emphasis mark only when document language is Japanese by checking language before setting Font.EmphasisMark. | document-formatting | docx | verified |
| 20 | `retrieve-and-display-the-emphasismark-value-of-a-run-for-debugging-east-asian-formatting.cs` | Retrieve and display the EmphasisMark value of a Run for debugging East Asian formatting. | document-formatting | docx | verified |
| 21 | `retrieve-and-display-the-current-font-size-of-a-run-for-debugging-purposes.cs` | Retrieve and display the current font size of a Run for debugging purposes. | document-formatting | docx | verified |
| 22 | `define-a-custom-font-substitution-rule-that-replaces-missing-times-new-roman-with-calibri.cs` | Define a custom font substitution rule that replaces missing Times New Roman with Calibri. | font-settings-and-sources | docx | verified |
| 23 | `configure-font-substitution-to-map-missing-garamond-to-a-locally-installed-georgia-font.cs` | Configure font substitution to map missing Garamond to a locally installed Georgia font. | font-settings-and-sources | docx | verified |
| 24 | `enable-automatic-font-substitution-for-missing-fonts-during-document-loading-by-configurin.cs` | Enable automatic font substitution for missing fonts during document loading by configuring FontSettings. | font-settings-and-sources | docx | verified |
| 25 | `create-a-custom-font-substitution-table-and-load-it-from-an-xml-configuration-file.cs` | Create a custom font substitution table and load it from an XML configuration file. | font-settings-and-sources | xml | verified |
| 26 | `embed-all-used-fonts-into-the-document-when-saving-by-enabling-the-embedfullfonts-option.cs` | Embed all used fonts into the document when saving by enabling the EmbedFullFonts option. | embedding-and-rendering | docx | verified |
| 27 | `enable-font-embedding-when-saving-documents-to-pdf-to-ensure-visual-fidelity-across-platfo.cs` | Enable font embedding when saving documents to PDF to ensure visual fidelity across platforms. | embedding-and-rendering | pdf | verified |
| 28 | `reset-all-font-attributes-of-a-run-to-defaults-by-calling-font-clearformatting.cs` | Reset all font attributes of a Run to defaults by calling Font.ClearFormatting. | document-formatting | docx | verified |
| 29 | `copy-font-formatting-from-one-run-to-another-using-run-font-clone-method.cs` | Copy font formatting from one Run to another using Run.Font.Clone method. | document-formatting | docx | verified |
| 30 | `create-a-reusable-method-that-applies-a-specific-font-name-and-size-to-any-given-run.cs` | Create a reusable method that applies a specific font name and size to any given Run. | document-formatting | docx | verified |
| 31 | `create-a-font-object-set-its-size-and-color-then-assign-it-to-multiple-runs-for-consistenc.cs` | Create a Font object, set its size and color, then assign it to multiple Runs for consistency. | document-formatting | docx | verified |
| 32 | `programmatically-replace-all-occurrences-of-a-specific-font-with-another-across-the-entire.cs` | Programmatically replace all occurrences of a specific font with another across the entire document. | document-formatting | docx | verified |
| 33 | `adjust-font-color-dynamically-based-on-paragraph-index-using-a-loop-and-conditional-logic.cs` | Adjust font color dynamically based on paragraph index using a loop and conditional logic. | reporting-and-inspection | docx | verified |
| 34 | `implement-batch-processing-to-set-helvetica-font-for-all-runs-in-a-collection-of-documents.cs` | Implement batch processing to set Helvetica font for all Runs in a collection of documents. | input-bootstrap | docx | verified |
| 35 | `load-a-document-change-all-headings-to-use-a-bold-sixteen-point-font-then-save.cs` | Load a document, change all headings to use a bold sixteen‑point font, then save. | document-formatting | docx | verified |

## Common failure patterns seen during generation and how they were corrected

### Font type ambiguity

- Symptom: Compile errors because Font is ambiguous between Aspose.Drawing.Font and Aspose.Words.Font.
- Fix: Use fully qualified names and use Aspose.Words.Font for document formatting versus Aspose.Drawing.Font only for real drawing scenarios.

### Color type ambiguity

- Symptom: Compile or assignment issues because Color is ambiguous or the wrong color type is used for the target property.
- Fix: Use fully qualified Aspose.Drawing.Color or explicit System.Drawing.Color conversion only where the Aspose.Words API expects it.

### Wrong formatting object

- Symptom: A drawing font object is treated as if it were a document text-formatting object.
- Fix: Apply document font values through Run.Font, Style.Font, ParagraphBreakFont, or DocumentBuilder.Font.

### Missing local bootstrap inputs

- Symptom: Examples assume existing documents, folders, or font resources already exist.
- Fix: Create local sample documents and deterministic local scenarios inside the example.

### Weak validation

- Symptom: Examples save output but do not verify that the intended font change, report, or render actually occurred.
- Fix: Validate the changed font property, configured font setting, or output file existence before finishing.

## Notes for maintainers

- This category is now **100% publish-ready** for the current run.
- Preserve file-to-task traceability when updating this folder.
- For future updates, keep the examples standalone and continue bootstrapping local inputs inside the example whenever external sources are mentioned.
