# Hyphenation Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Hyphenation category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Hyphenation**
- Slug: **hyphenation**
- Total examples: **40**
- Publish-ready successful examples: **40 / 40**
- Dictionary registration examples: **14**
- Rendering / export examples: **8**
- Document setup / validation examples: **18**

## Category rules that shaped these examples

- Use native Aspose.Words hyphenation APIs directly.
- Create realistic local sample inputs whenever the task mentions an existing DOC, DOCX, TXT, stream, folder, or dictionary source.
- Create deterministic local dictionary files for live examples instead of depending on remote downloads.
- Use deterministic layout only when the task requires visibly hyphenated output.
- Validate that the requested rendered or report output is actually written.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\hyphenation\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `hyphenation/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# PowerShell example
Copy-Item ..\hyphenation\load-a-docx-file-register-a-custom-hunspell-dictionary-and-enable-automatic-hyphenation.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-register-a-custom-hunspell-dictionary-and-enable-automatic-hyphenation.cs` | Load a DOCX file, register a custom Hunspell dictionary, and enable automatic hyphenation. | dictionary-registration | docx | verified |
| 2 | `register-an-external-libreoffice-dictionary-from-github-and-apply-it-to-documents-written.cs` | Register an external LibreOffice dictionary from GitHub and apply it to documents written in Spanish. | dictionary-registration | doc | verified |
| 3 | `create-a-script-that-downloads-the-latest-libreoffice-hyphenation-dictionaries-and-registe.cs` | Create a script that downloads the latest LibreOffice hyphenation dictionaries and registers them automatically. | document-setup-and-validation | docx | verified |
| 4 | `update-an-existing-hunspell-dictionary-by-adding-custom-hyphenation-patterns-for-technical.cs` | Update an existing Hunspell dictionary by adding custom hyphenation patterns for technical terminology. | dictionary-registration | docx | verified |
| 5 | `set-the-hyphenation-language-to-en-us-and-verify-word-breaking-matches-microsoft-word-beha.cs` | Set the hyphenation language to en‑US and verify word breaking matches Microsoft Word behavior. | dictionary-registration | docx | verified |
| 6 | `load-a-docx-containing-mixed-languages-set-appropriate-hyphenation-language-for-each-secti.cs` | Load a DOCX containing mixed languages, set appropriate hyphenation language for each section. | dictionary-registration | docx | verified |
| 7 | `set-hyphenation-to-respect-word-boundaries-in-arabic-script-by-loading-appropriate-right-t.cs` | Set hyphenation to respect word boundaries in Arabic script by loading appropriate right‑to‑left dictionary. | dictionary-registration | docx | verified |
| 8 | `configure-hyphenation-to-respect-compound-word-rules-of-microsoft-word-for-german-language.cs` | Configure hyphenation to respect compound‑word rules of Microsoft Word for German language documents. | dictionary-registration | docx | verified |
| 9 | `set-minimum-word-length-for-hyphenation-to-five-characters-and-observe-its-effect-on-short.cs` | Set minimum word length for hyphenation to five characters and observe its effect on short words. | document-setup-and-validation | docx | verified |
| 10 | `validate-that-hyphenation-does-not-occur-for-words-shorter-than-the-configured-minimum-len.cs` | Validate that hyphenation does not occur for words shorter than the configured minimum length. | document-setup-and-validation | docx | verified |
| 11 | `enable-hyphenation-globally-for-a-document-then-override-it-for-a-single-table-cell.cs` | Enable hyphenation globally for a document, then override it for a single table cell. | document-setup-and-validation | docx | verified |
| 12 | `disable-hyphenation-for-headings-while-keeping-it-enabled-for-body-paragraphs-in-a-report.cs` | Disable hyphenation for headings while keeping it enabled for body paragraphs in a report. | document-setup-and-validation | docx | verified |
| 13 | `apply-hyphenation-only-to-a-selected-range-using-documentbuilder-and-verify-layout-changes.cs` | Apply hyphenation only to a selected range using DocumentBuilder and verify layout changes. | document-setup-and-validation | docx | verified |
| 14 | `apply-hyphenation-to-a-specific-paragraph-by-disabling-it-for-the-surrounding-sections.cs` | Apply hyphenation to a specific paragraph by disabling it for the surrounding sections. | document-setup-and-validation | docx | verified |
| 15 | `load-a-pdf-file-enable-hyphenation-and-render-the-result-to-an-image-for-visual-inspection.cs` | Load a PDF file, enable hyphenation, and render the result to an image for visual inspection. | rendering-and-export | pdf | verified |
| 16 | `use-the-api-to-query-hyphenation-status-of-each-word-in-a-paragraph-and-log-the-results.cs` | Use the API to query hyphenation status of each word in a paragraph and log the results. | document-setup-and-validation | docx | verified |
| 17 | `check-whether-the-word-hyphenation-will-be-hyphenated-using-the-api-before-document-genera.cs` | Check whether the word 'hyphenation' will be hyphenated using the API before document generation. | document-setup-and-validation | docx | verified |
| 18 | `develop-a-function-that-returns-true-if-a-given-word-will-be-hyphenated-under-current-sett.cs` | Develop a function that returns true if a given word will be hyphenated under current settings. | document-setup-and-validation | docx | verified |
| 19 | `retrieve-the-hyphenation-patterns-for-french-language-and-log-them-for-debugging-purposes.cs` | Retrieve the hyphenation patterns for French language and log them for debugging purposes. | dictionary-registration | docx | verified |
| 20 | `list-all-available-hyphenation-dictionaries-on-the-system-and-display-their-language-codes.cs` | List all available hyphenation dictionaries on the system and display their language codes. | dictionary-registration | docx | verified |
| 21 | `programmatically-adjust-paragraph-justification-after-hyphenation-to-prevent-excessive-whi.cs` | Programmatically adjust paragraph justification after hyphenation to prevent excessive white space gaps. | document-setup-and-validation | docx | verified |
| 22 | `apply-hyphenation-to-a-document-then-programmatically-adjust-line-spacing-to-improve-reada.cs` | Apply hyphenation to a document, then programmatically adjust line spacing to improve readability. | document-setup-and-validation | docx | verified |
| 23 | `measure-pagination-differences-after-enabling-hyphenation-on-a-multi-section-report-docume.cs` | Measure pagination differences after enabling hyphenation on a multi‑section report document for analysis. | document-setup-and-validation | docx | verified |
| 24 | `compare-layout-differences-between-a-document-saved-with-hyphenation-disabled-and-the-same.cs` | Compare layout differences between a document saved with hyphenation disabled and the same document with hyphenation enabled. | document-setup-and-validation | docx | verified |
| 25 | `load-multiple-documents-from-a-folder-enable-hyphenation-and-save-each-as-a-hyphenated-pdf.cs` | Load multiple documents from a folder, enable hyphenation, and save each as a hyphenated PDF. | rendering-and-export | pdf | verified |
| 26 | `batch-process-a-collection-of-docx-files-applying-language-specific-hyphenation-and-export.cs` | Batch process a collection of DOCX files, applying language‑specific hyphenation and exporting each to PDF. | dictionary-registration | docx, pdf | verified |
| 27 | `batch-convert-docx-files-to-pdf-while-preserving-hyphenation-and-log-any-documents-that-fa.cs` | Batch convert DOCX files to PDF while preserving hyphenation, and log any documents that fail. | rendering-and-export | docx, pdf | verified |
| 28 | `create-a-new-document-enable-automatic-hyphenation-from-the-start-and-save-it-as-a-docx-fi.cs` | Create a new document, enable automatic hyphenation from the start, and save it as a DOCX file. | document-setup-and-validation | docx | verified |
| 29 | `load-a-docx-enable-hyphenation-and-export-the-document-to-docx-preserving-hyphenation-mark.cs` | Load a DOCX, enable hyphenation, and export the document to DOCX preserving hyphenation marks. | rendering-and-export | docx, doc | verified |
| 30 | `enable-hyphenation-for-a-document-then-export-to-docx-and-verify-hyphenation-marks-are-ret.cs` | Enable hyphenation for a document, then export to DOCX and verify hyphenation marks are retained. | rendering-and-export | docx, doc | verified |
| 31 | `export-a-hyphenated-document-to-pdf-and-compare-file-size-with-a-non-hyphenated-version.cs` | Export a hyphenated document to PDF and compare file size with a non‑hyphenated version. | rendering-and-export | pdf | verified |
| 32 | `generate-a-pdf-from-a-hyphenated-document-and-ensure-that-hyphenation-marks-are-not-visibl.cs` | Generate a PDF from a hyphenated document and ensure that hyphenation marks are not visible in the output. | rendering-and-export | pdf | verified |
| 33 | `load-a-docx-set-hyphenation-language-to-ru-ru-and-save-the-result-as-pdf.cs` | Load a DOCX, set hyphenation language to ru‑RU, and save the result as PDF. | dictionary-registration | docx, pdf | verified |
| 34 | `measure-rendering-time-differences-between-pdf-generation-with-and-without-hyphenation-ena.cs` | Measure rendering time differences between PDF generation with and without hyphenation enabled. | rendering-and-export | pdf | verified |
| 35 | `create-a-console-application-that-accepts-a-document-path-hyphenation-language-code-and-ou.cs` | Create a console application that accepts a document path, hyphenation language code, and outputs a hyphenated PDF. | dictionary-registration | pdf | verified |
| 36 | `implement-error-handling-for-loading-incompatible-hyphenation-dictionaries-and-provide-des.cs` | Implement error handling for loading incompatible hyphenation dictionaries and provide descriptive messages. | document-setup-and-validation | docx | verified |
| 37 | `integrate-hyphenation-dictionary-updates-into-a-ci-pipeline-to-keep-language-patterns-curr.cs` | Integrate hyphenation dictionary updates into a CI pipeline to keep language patterns current. | dictionary-registration | docx | verified |
| 38 | `test-hyphenation-impact-on-page-count-by-generating-a-long-report-with-and-without-hyphena.cs` | Test hyphenation impact on page count by generating a long report with and without hyphenation. | document-setup-and-validation | docx | verified |
| 39 | `load-a-docx-disable-hyphenation-for-footnotes-only-and-compare-footnote-layout-before-and.cs` | Load a DOCX, disable hyphenation for footnotes only, and compare footnote layout before and after. | document-setup-and-validation | docx | verified |
| 40 | `create-a-unit-test-that-verifies-hyphenation-patterns-are-applied-correctly-for-italian-la.cs` | Create a unit test that verifies hyphenation patterns are applied correctly for Italian language documents. | dictionary-registration | docx | verified |

## Common failure patterns seen during generation and how they were corrected

### Remote dictionary dependency

- Symptom: Example depends on a GitHub or other external URL for a hyphenation dictionary file.
- Fix: Create a deterministic local dictionary file inside the example and register that local file instead.

### Wrong or invented hyphenation API usage

- Symptom: Compile failures or ineffective behavior because unsupported hyphenation helpers or callback hooks are used.
- Fix: Use only documented Aspose.Words hyphenation registration and callback APIs.

### Visible hyphenation not triggered

- Symptom: Example configures hyphenation but the layout never becomes narrow enough for hyphenation to matter.
- Fix: Use deterministic page or layout settings and sufficiently long sample text so hyphenation can realistically occur.

### Missing local bootstrap input

- Symptom: Example assumes an existing document, stream, folder, or dictionary already exists.
- Fix: Create local sample documents, source files, folders, and dictionary files inside the example.

## Notes for maintainers

- This category is publish-ready for the current run.
- One known quality correction was to prefer local deterministic dictionary files over a remote GitHub dictionary link.
- Preserve file-to-task traceability when updating this folder.
