---
name: hyphenation
description: Verified C# examples for Hyphenation scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Hyphenation

## Purpose

This folder is a live, curated example set for Hyphenation scenarios. Treat every `.cs` file as a standalone console application. The goal is correct, warning-free examples that use documented Aspose APIs and match the original task intent.

## Non-negotiable conventions

- Use native Aspose.Words hyphenation APIs directly.
- Bootstrap local source documents, files, streams, or folders whenever the task implies an existing source.
- Create deterministic local dictionary files when a task requires a hyphenation dictionary.
- Do not depend on remote dictionary downloads for live examples.
- Validate that the configured hyphenation result, report, or export actually exists.
- Guard maybe-null values to avoid nullable-reference warnings such as CS8600, CS8602, and CS8604.

## Recommended workflow selection

- Dictionary Registration workflow: 14 examples
- Document Setup And Validation workflow: 18 examples
- Rendering And Export workflow: 8 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. API usage must be supported by the configured package versions.
3. Exported outputs must actually be written by the example.
4. Validation scenarios must inspect only the behavior requested by the task.
5. Examples that depend on files, folders, streams, images, or data should bootstrap those inputs locally during the example run.

## File-to-task reference

- `load-a-docx-file-register-a-custom-hunspell-dictionary-and-enable-automatic-hyphenation.cs`
  - Task: Load a DOCX file, register a custom Hunspell dictionary, and enable automatic hyphenation.
  - Workflow: Dictionary Registration
  - Outputs: docx
  - Selected engine: mcp
- `register-an-external-libreoffice-dictionary-from-github-and-apply-it-to-documents-written.cs`
  - Task: Register an external LibreOffice dictionary from GitHub and apply it to documents written in Spanish.
  - Workflow: Dictionary Registration
  - Outputs: doc
  - Selected engine: mcp
- `create-a-script-that-downloads-the-latest-libreoffice-hyphenation-dictionaries-and-registe.cs`
  - Task: Create a script that downloads the latest LibreOffice hyphenation dictionaries and registers them automatically.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `update-an-existing-hunspell-dictionary-by-adding-custom-hyphenation-patterns-for-technical.cs`
  - Task: Update an existing Hunspell dictionary by adding custom hyphenation patterns for technical terminology.
  - Workflow: Dictionary Registration
  - Outputs: docx
  - Selected engine: mcp
- `set-the-hyphenation-language-to-en-us-and-verify-word-breaking-matches-microsoft-word-beha.cs`
  - Task: Set the hyphenation language to en-US and verify word breaking matches Microsoft Word behavior.
  - Workflow: Dictionary Registration
  - Outputs: docx
  - Selected engine: mcp
- `load-a-docx-containing-mixed-languages-set-appropriate-hyphenation-language-for-each-secti.cs`
  - Task: Load a DOCX containing mixed languages, set appropriate hyphenation language for each section.
  - Workflow: Dictionary Registration
  - Outputs: docx
  - Selected engine: mcp
- `set-hyphenation-to-respect-word-boundaries-in-arabic-script-by-loading-appropriate-right-t.cs`
  - Task: Set hyphenation to respect word boundaries in Arabic script by loading appropriate right-to-left dictionary.
  - Workflow: Dictionary Registration
  - Outputs: docx
  - Selected engine: mcp
- `configure-hyphenation-to-respect-compound-word-rules-of-microsoft-word-for-german-language.cs`
  - Task: Configure hyphenation to respect compound-word rules of Microsoft Word for German language documents.
  - Workflow: Dictionary Registration
  - Outputs: docx
  - Selected engine: mcp
- `set-minimum-word-length-for-hyphenation-to-five-characters-and-observe-its-effect-on-short.cs`
  - Task: Set minimum word length for hyphenation to five characters and observe its effect on short words.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `validate-that-hyphenation-does-not-occur-for-words-shorter-than-the-configured-minimum-len.cs`
  - Task: Validate that hyphenation does not occur for words shorter than the configured minimum length.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `enable-hyphenation-globally-for-a-document-then-override-it-for-a-single-table-cell.cs`
  - Task: Enable hyphenation globally for a document, then override it for a single table cell.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `disable-hyphenation-for-headings-while-keeping-it-enabled-for-body-paragraphs-in-a-report.cs`
  - Task: Disable hyphenation for headings while keeping it enabled for body paragraphs in a report.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `apply-hyphenation-only-to-a-selected-range-using-documentbuilder-and-verify-layout-changes.cs`
  - Task: Apply hyphenation only to a selected range using DocumentBuilder and verify layout changes.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `apply-hyphenation-to-a-specific-paragraph-by-disabling-it-for-the-surrounding-sections.cs`
  - Task: Apply hyphenation to a specific paragraph by disabling it for the surrounding sections.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `load-a-pdf-file-enable-hyphenation-and-render-the-result-to-an-image-for-visual-inspection.cs`
  - Task: Load a PDF file, enable hyphenation, and render the result to an image for visual inspection.
  - Workflow: Rendering And Export
  - Outputs: pdf
  - Selected engine: mcp
- `use-the-api-to-query-hyphenation-status-of-each-word-in-a-paragraph-and-log-the-results.cs`
  - Task: Use the API to query hyphenation status of each word in a paragraph and log the results.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `check-whether-the-word-hyphenation-will-be-hyphenated-using-the-api-before-document-genera.cs`
  - Task: Check whether the word 'hyphenation' will be hyphenated using the API before document generation.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `develop-a-function-that-returns-true-if-a-given-word-will-be-hyphenated-under-current-sett.cs`
  - Task: Develop a function that returns true if a given word will be hyphenated under current settings.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-the-hyphenation-patterns-for-french-language-and-log-them-for-debugging-purposes.cs`
  - Task: Retrieve the hyphenation patterns for French language and log them for debugging purposes.
  - Workflow: Dictionary Registration
  - Outputs: docx
  - Selected engine: mcp
- `list-all-available-hyphenation-dictionaries-on-the-system-and-display-their-language-codes.cs`
  - Task: List all available hyphenation dictionaries on the system and display their language codes.
  - Workflow: Dictionary Registration
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-adjust-paragraph-justification-after-hyphenation-to-prevent-excessive-whi.cs`
  - Task: Programmatically adjust paragraph justification after hyphenation to prevent excessive white space gaps.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `apply-hyphenation-to-a-document-then-programmatically-adjust-line-spacing-to-improve-reada.cs`
  - Task: Apply hyphenation to a document, then programmatically adjust line spacing to improve readability.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `measure-pagination-differences-after-enabling-hyphenation-on-a-multi-section-report-docume.cs`
  - Task: Measure pagination differences after enabling hyphenation on a multi-section report document for analysis.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `compare-layout-differences-between-a-document-saved-with-hyphenation-disabled-and-the-same.cs`
  - Task: Compare layout differences between a document saved with hyphenation disabled and the same document with hyphenation enabled.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `load-multiple-documents-from-a-folder-enable-hyphenation-and-save-each-as-a-hyphenated-pdf.cs`
  - Task: Load multiple documents from a folder, enable hyphenation, and save each as a hyphenated PDF.
  - Workflow: Rendering And Export
  - Outputs: pdf
  - Selected engine: mcp
- `batch-process-a-collection-of-docx-files-applying-language-specific-hyphenation-and-export.cs`
  - Task: Batch process a collection of DOCX files, applying language-specific hyphenation and exporting each to PDF.
  - Workflow: Dictionary Registration
  - Outputs: docx, pdf
  - Selected engine: mcp
- `batch-convert-docx-files-to-pdf-while-preserving-hyphenation-and-log-any-documents-that-fa.cs`
  - Task: Batch convert DOCX files to PDF while preserving hyphenation, and log any documents that fail.
  - Workflow: Rendering And Export
  - Outputs: docx, pdf
  - Selected engine: mcp
- `create-a-new-document-enable-automatic-hyphenation-from-the-start-and-save-it-as-a-docx-fi.cs`
  - Task: Create a new document, enable automatic hyphenation from the start, and save it as a DOCX file.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `load-a-docx-enable-hyphenation-and-export-the-document-to-docx-preserving-hyphenation-mark.cs`
  - Task: Load a DOCX, enable hyphenation, and export the document to DOCX preserving hyphenation marks.
  - Workflow: Rendering And Export
  - Outputs: docx, doc
  - Selected engine: mcp
- `enable-hyphenation-for-a-document-then-export-to-docx-and-verify-hyphenation-marks-are-ret.cs`
  - Task: Enable hyphenation for a document, then export to DOCX and verify hyphenation marks are retained.
  - Workflow: Rendering And Export
  - Outputs: docx, doc
  - Selected engine: mcp
- `export-a-hyphenated-document-to-pdf-and-compare-file-size-with-a-non-hyphenated-version.cs`
  - Task: Export a hyphenated document to PDF and compare file size with a non-hyphenated version.
  - Workflow: Rendering And Export
  - Outputs: pdf
  - Selected engine: mcp
- `generate-a-pdf-from-a-hyphenated-document-and-ensure-that-hyphenation-marks-are-not-visibl.cs`
  - Task: Generate a PDF from a hyphenated document and ensure that hyphenation marks are not visible in the output.
  - Workflow: Rendering And Export
  - Outputs: pdf
  - Selected engine: mcp
- `load-a-docx-set-hyphenation-language-to-ru-ru-and-save-the-result-as-pdf.cs`
  - Task: Load a DOCX, set hyphenation language to ru-RU, and save the result as PDF.
  - Workflow: Dictionary Registration
  - Outputs: docx, pdf
  - Selected engine: mcp
- `measure-rendering-time-differences-between-pdf-generation-with-and-without-hyphenation-ena.cs`
  - Task: Measure rendering time differences between PDF generation with and without hyphenation enabled.
  - Workflow: Rendering And Export
  - Outputs: pdf
  - Selected engine: mcp
- `create-a-console-application-that-accepts-a-document-path-hyphenation-language-code-and-ou.cs`
  - Task: Create a console application that accepts a document path, hyphenation language code, and outputs a hyphenated PDF.
  - Workflow: Dictionary Registration
  - Outputs: pdf
  - Selected engine: llm
- `implement-error-handling-for-loading-incompatible-hyphenation-dictionaries-and-provide-des.cs`
  - Task: Implement error handling for loading incompatible hyphenation dictionaries and provide descriptive messages.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `integrate-hyphenation-dictionary-updates-into-a-ci-pipeline-to-keep-language-patterns-curr.cs`
  - Task: Integrate hyphenation dictionary updates into a CI pipeline to keep language patterns current.
  - Workflow: Dictionary Registration
  - Outputs: docx
  - Selected engine: mcp
- `test-hyphenation-impact-on-page-count-by-generating-a-long-report-with-and-without-hyphena.cs`
  - Task: Test hyphenation impact on page count by generating a long report with and without hyphenation.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: mcp
- `load-a-docx-disable-hyphenation-for-footnotes-only-and-compare-footnote-layout-before-and.cs`
  - Task: Load a DOCX, disable hyphenation for footnotes only, and compare footnote layout before and after.
  - Workflow: Document Setup And Validation
  - Outputs: docx
  - Selected engine: existing_repo
- `create-a-unit-test-that-verifies-hyphenation-patterns-are-applied-correctly-for-italian-la.cs`
  - Task: Create a unit test that verifies hyphenation patterns are applied correctly for Italian language documents.
  - Workflow: Dictionary Registration
  - Outputs: docx
  - Selected engine: existing_repo

## Common failure patterns and preferred agent fixes

- Remote dictionary dependency
  - Symptom: Example depends on a GitHub or other external URL for a hyphenation dictionary file.
  - Preferred fix: Create a deterministic local dictionary file inside the example and register that local file instead.

- Wrong or invented hyphenation API usage
  - Symptom: Compile failures or ineffective behavior because unsupported hyphenation helpers or callback hooks are used.
  - Preferred fix: Use only documented Aspose.Words hyphenation registration and callback APIs.

- Visible hyphenation not triggered
  - Symptom: Example configures hyphenation but the layout never becomes narrow enough for hyphenation to matter.
  - Preferred fix: Use deterministic page or layout settings and sufficiently long sample text so hyphenation can realistically occur.

- Missing local bootstrap input
  - Symptom: Example assumes an existing document, stream, folder, or dictionary already exists.
  - Preferred fix: Create local sample documents, source files, folders, and dictionary files inside the example.

## Build and run contract

- Target framework: `net8.0`
- Package: `Aspose.Words` `26.5.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.5.0
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
- Keep examples standalone, bootstrap missing inputs locally, and prefer documented Aspose APIs over speculative shortcuts.
