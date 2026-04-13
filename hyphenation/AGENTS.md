---
name: hyphenation
description: Verified C# examples for hyphenation scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Hyphenation

## Purpose

This folder is a live, curated example set for hyphenation scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free hyphenation dictionary registration, callback usage, document setup, and rendering/export workflows using direct Aspose.Words APIs.

## Non-negotiable conventions

- Use native Aspose.Words hyphenation APIs directly.
- Bootstrap local source documents, files, streams, or folders whenever the task implies an existing source.
- Create deterministic local dictionary files when a task requires a hyphenation dictionary.
- Do not depend on remote dictionary downloads for live examples.
- Validate that the configured hyphenation result, report, or export actually exists.
- Guard maybe-null values to avoid nullable-reference warnings such as CS8600, CS8602, and CS8604.

## Recommended workflow selection

- **Dictionary registration workflow**: 14 examples
- **Rendering / export workflow**: 8 examples
- **Document setup / validation workflow**: 18 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. Source documents, streams, folders, and dictionary files must be bootstrapped locally whenever the task implies an existing input.
3. Hyphenation configuration must use documented Aspose.Words APIs only.
4. Requested report or rendered output files must actually be written.
5. Live examples should prefer deterministic local dictionary files rather than remote downloads.

## File-to-task reference

- `load-a-docx-file-register-a-custom-hunspell-dictionary-and-enable-automatic-hyphenation.cs`
  - Task: Load a DOCX file, register a custom Hunspell dictionary, and enable automatic hyphenation.
  - Workflow: dictionary-registration
  - Outputs: docx
  - Selected engine: verified
- `register-an-external-libreoffice-dictionary-from-github-and-apply-it-to-documents-written.cs`
  - Task: Register an external LibreOffice dictionary from GitHub and apply it to documents written in Spanish.
  - Workflow: dictionary-registration
  - Outputs: doc
  - Selected engine: verified
- `create-a-script-that-downloads-the-latest-libreoffice-hyphenation-dictionaries-and-registe.cs`
  - Task: Create a script that downloads the latest LibreOffice hyphenation dictionaries and registers them automatically.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `update-an-existing-hunspell-dictionary-by-adding-custom-hyphenation-patterns-for-technical.cs`
  - Task: Update an existing Hunspell dictionary by adding custom hyphenation patterns for technical terminology.
  - Workflow: dictionary-registration
  - Outputs: docx
  - Selected engine: verified
- `set-the-hyphenation-language-to-en-us-and-verify-word-breaking-matches-microsoft-word-beha.cs`
  - Task: Set the hyphenation language to en‑US and verify word breaking matches Microsoft Word behavior.
  - Workflow: dictionary-registration
  - Outputs: docx
  - Selected engine: verified
- `load-a-docx-containing-mixed-languages-set-appropriate-hyphenation-language-for-each-secti.cs`
  - Task: Load a DOCX containing mixed languages, set appropriate hyphenation language for each section.
  - Workflow: dictionary-registration
  - Outputs: docx
  - Selected engine: verified
- `set-hyphenation-to-respect-word-boundaries-in-arabic-script-by-loading-appropriate-right-t.cs`
  - Task: Set hyphenation to respect word boundaries in Arabic script by loading appropriate right‑to‑left dictionary.
  - Workflow: dictionary-registration
  - Outputs: docx
  - Selected engine: verified
- `configure-hyphenation-to-respect-compound-word-rules-of-microsoft-word-for-german-language.cs`
  - Task: Configure hyphenation to respect compound‑word rules of Microsoft Word for German language documents.
  - Workflow: dictionary-registration
  - Outputs: docx
  - Selected engine: verified
- `set-minimum-word-length-for-hyphenation-to-five-characters-and-observe-its-effect-on-short.cs`
  - Task: Set minimum word length for hyphenation to five characters and observe its effect on short words.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `validate-that-hyphenation-does-not-occur-for-words-shorter-than-the-configured-minimum-len.cs`
  - Task: Validate that hyphenation does not occur for words shorter than the configured minimum length.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `enable-hyphenation-globally-for-a-document-then-override-it-for-a-single-table-cell.cs`
  - Task: Enable hyphenation globally for a document, then override it for a single table cell.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `disable-hyphenation-for-headings-while-keeping-it-enabled-for-body-paragraphs-in-a-report.cs`
  - Task: Disable hyphenation for headings while keeping it enabled for body paragraphs in a report.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `apply-hyphenation-only-to-a-selected-range-using-documentbuilder-and-verify-layout-changes.cs`
  - Task: Apply hyphenation only to a selected range using DocumentBuilder and verify layout changes.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `apply-hyphenation-to-a-specific-paragraph-by-disabling-it-for-the-surrounding-sections.cs`
  - Task: Apply hyphenation to a specific paragraph by disabling it for the surrounding sections.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `load-a-pdf-file-enable-hyphenation-and-render-the-result-to-an-image-for-visual-inspection.cs`
  - Task: Load a PDF file, enable hyphenation, and render the result to an image for visual inspection.
  - Workflow: rendering-and-export
  - Outputs: pdf
  - Selected engine: verified
- `use-the-api-to-query-hyphenation-status-of-each-word-in-a-paragraph-and-log-the-results.cs`
  - Task: Use the API to query hyphenation status of each word in a paragraph and log the results.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `check-whether-the-word-hyphenation-will-be-hyphenated-using-the-api-before-document-genera.cs`
  - Task: Check whether the word 'hyphenation' will be hyphenated using the API before document generation.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `develop-a-function-that-returns-true-if-a-given-word-will-be-hyphenated-under-current-sett.cs`
  - Task: Develop a function that returns true if a given word will be hyphenated under current settings.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `retrieve-the-hyphenation-patterns-for-french-language-and-log-them-for-debugging-purposes.cs`
  - Task: Retrieve the hyphenation patterns for French language and log them for debugging purposes.
  - Workflow: dictionary-registration
  - Outputs: docx
  - Selected engine: verified
- `list-all-available-hyphenation-dictionaries-on-the-system-and-display-their-language-codes.cs`
  - Task: List all available hyphenation dictionaries on the system and display their language codes.
  - Workflow: dictionary-registration
  - Outputs: docx
  - Selected engine: verified
- `programmatically-adjust-paragraph-justification-after-hyphenation-to-prevent-excessive-whi.cs`
  - Task: Programmatically adjust paragraph justification after hyphenation to prevent excessive white space gaps.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `apply-hyphenation-to-a-document-then-programmatically-adjust-line-spacing-to-improve-reada.cs`
  - Task: Apply hyphenation to a document, then programmatically adjust line spacing to improve readability.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `measure-pagination-differences-after-enabling-hyphenation-on-a-multi-section-report-docume.cs`
  - Task: Measure pagination differences after enabling hyphenation on a multi‑section report document for analysis.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `compare-layout-differences-between-a-document-saved-with-hyphenation-disabled-and-the-same.cs`
  - Task: Compare layout differences between a document saved with hyphenation disabled and the same document with hyphenation enabled.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `load-multiple-documents-from-a-folder-enable-hyphenation-and-save-each-as-a-hyphenated-pdf.cs`
  - Task: Load multiple documents from a folder, enable hyphenation, and save each as a hyphenated PDF.
  - Workflow: rendering-and-export
  - Outputs: pdf
  - Selected engine: verified
- `batch-process-a-collection-of-docx-files-applying-language-specific-hyphenation-and-export.cs`
  - Task: Batch process a collection of DOCX files, applying language‑specific hyphenation and exporting each to PDF.
  - Workflow: dictionary-registration
  - Outputs: docx, pdf
  - Selected engine: verified
- `batch-convert-docx-files-to-pdf-while-preserving-hyphenation-and-log-any-documents-that-fa.cs`
  - Task: Batch convert DOCX files to PDF while preserving hyphenation, and log any documents that fail.
  - Workflow: rendering-and-export
  - Outputs: docx, pdf
  - Selected engine: verified
- `create-a-new-document-enable-automatic-hyphenation-from-the-start-and-save-it-as-a-docx-fi.cs`
  - Task: Create a new document, enable automatic hyphenation from the start, and save it as a DOCX file.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `load-a-docx-enable-hyphenation-and-export-the-document-to-docx-preserving-hyphenation-mark.cs`
  - Task: Load a DOCX, enable hyphenation, and export the document to DOCX preserving hyphenation marks.
  - Workflow: rendering-and-export
  - Outputs: docx, doc
  - Selected engine: verified
- `enable-hyphenation-for-a-document-then-export-to-docx-and-verify-hyphenation-marks-are-ret.cs`
  - Task: Enable hyphenation for a document, then export to DOCX and verify hyphenation marks are retained.
  - Workflow: rendering-and-export
  - Outputs: docx, doc
  - Selected engine: verified
- `export-a-hyphenated-document-to-pdf-and-compare-file-size-with-a-non-hyphenated-version.cs`
  - Task: Export a hyphenated document to PDF and compare file size with a non‑hyphenated version.
  - Workflow: rendering-and-export
  - Outputs: pdf
  - Selected engine: verified
- `generate-a-pdf-from-a-hyphenated-document-and-ensure-that-hyphenation-marks-are-not-visibl.cs`
  - Task: Generate a PDF from a hyphenated document and ensure that hyphenation marks are not visible in the output.
  - Workflow: rendering-and-export
  - Outputs: pdf
  - Selected engine: verified
- `load-a-docx-set-hyphenation-language-to-ru-ru-and-save-the-result-as-pdf.cs`
  - Task: Load a DOCX, set hyphenation language to ru‑RU, and save the result as PDF.
  - Workflow: dictionary-registration
  - Outputs: docx, pdf
  - Selected engine: verified
- `measure-rendering-time-differences-between-pdf-generation-with-and-without-hyphenation-ena.cs`
  - Task: Measure rendering time differences between PDF generation with and without hyphenation enabled.
  - Workflow: rendering-and-export
  - Outputs: pdf
  - Selected engine: verified
- `create-a-console-application-that-accepts-a-document-path-hyphenation-language-code-and-ou.cs`
  - Task: Create a console application that accepts a document path, hyphenation language code, and outputs a hyphenated PDF.
  - Workflow: dictionary-registration
  - Outputs: pdf
  - Selected engine: verified
- `implement-error-handling-for-loading-incompatible-hyphenation-dictionaries-and-provide-des.cs`
  - Task: Implement error handling for loading incompatible hyphenation dictionaries and provide descriptive messages.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `integrate-hyphenation-dictionary-updates-into-a-ci-pipeline-to-keep-language-patterns-curr.cs`
  - Task: Integrate hyphenation dictionary updates into a CI pipeline to keep language patterns current.
  - Workflow: dictionary-registration
  - Outputs: docx
  - Selected engine: verified
- `test-hyphenation-impact-on-page-count-by-generating-a-long-report-with-and-without-hyphena.cs`
  - Task: Test hyphenation impact on page count by generating a long report with and without hyphenation.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `load-a-docx-disable-hyphenation-for-footnotes-only-and-compare-footnote-layout-before-and.cs`
  - Task: Load a DOCX, disable hyphenation for footnotes only, and compare footnote layout before and after.
  - Workflow: document-setup-and-validation
  - Outputs: docx
  - Selected engine: verified
- `create-a-unit-test-that-verifies-hyphenation-patterns-are-applied-correctly-for-italian-la.cs`
  - Task: Create a unit test that verifies hyphenation patterns are applied correctly for Italian language documents.
  - Workflow: dictionary-registration
  - Outputs: docx
  - Selected engine: verified

## Common failure patterns and preferred agent fixes

- **Remote dictionary dependency**
  - Symptom: Example depends on a GitHub or other external URL for a hyphenation dictionary file.
  - Preferred fix: Create a deterministic local dictionary file inside the example and register that local file instead.

- **Wrong or invented hyphenation API usage**
  - Symptom: Compile failures or ineffective behavior because unsupported hyphenation helpers or callback hooks are used.
  - Preferred fix: Use only documented Aspose.Words hyphenation registration and callback APIs.

- **Visible hyphenation not triggered**
  - Symptom: Example configures hyphenation but the layout never becomes narrow enough for hyphenation to matter.
  - Preferred fix: Use deterministic page or layout settings and sufficiently long sample text so hyphenation can realistically occur.

- **Missing local bootstrap input**
  - Symptom: Example assumes an existing document, stream, folder, or dictionary already exists.
  - Preferred fix: Create local sample documents, source files, folders, and dictionary files inside the example.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\hyphenation\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer local deterministic dictionary files over remote downloads.
