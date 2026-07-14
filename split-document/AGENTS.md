---
name: split-document
description: Verified C# examples for Split-Document scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Split-Document

## Purpose

This folder is a live, curated example set for Split-Document scenarios. Treat every `.cs` file as a standalone console application. The goal is correct, warning-free examples that use documented Aspose APIs and match the original task intent.

## Non-negotiable conventions

- Always use documented Aspose.Words APIs directly.
- Always create local sample source documents when a task refers to an existing file, folder, stream, template, or input asset.
- Prefer documented split, extraction, import, and clone workflows.
- Keep validation narrow and task-specific.
- Do not invent split, page-extraction, or import helper APIs.

## Recommended workflow selection

- Split Workflow workflow: 30 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. API usage must be supported by the configured package versions.
3. Exported outputs must actually be written by the example.
4. Validation scenarios must inspect only the behavior requested by the task.
5. Examples that depend on files, folders, streams, images, or data should bootstrap those inputs locally during the example run.

## File-to-task reference

- `load-a-docx-source-document-using-the-document-class-before-splitting.cs`
  - Task: Load a DOCX source document using the Document class before splitting.
  - Workflow: Split Workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `create-a-documentsplitcriteria-object-and-set-split-mode-to-headings.cs`
  - Task: Create a DocumentSplitCriteria object and set split mode to headings.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-documentsplitcriteria-object-and-set-split-mode-to-sections.cs`
  - Task: Create a DocumentSplitCriteria object and set split mode to sections.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-documentsplitcriteria-object-and-set-split-mode-to-individual-pages.cs`
  - Task: Create a DocumentSplitCriteria object and set split mode to individual pages.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-documentsplitcriteria-object-and-set-split-mode-to-custom-page-ranges.cs`
  - Task: Create a DocumentSplitCriteria object and set split mode to custom page ranges.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: llm
- `combine-heading-and-section-flags-in-documentsplitcriteria-to-split-by-both-structures.cs`
  - Task: Combine heading and section flags in DocumentSplitCriteria to split by both structures.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `combine-page-and-heading-flags-in-documentsplitcriteria-to-start-each-part-on-a-new-page.cs`
  - Task: Combine page and heading flags in DocumentSplitCriteria to start each part on a new page.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `call-document-split-criteria-to-obtain-a-collection-of-split-document-objects.cs`
  - Task: Call Document.Split(criteria) to obtain a collection of split Document objects.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `iterate-over-the-split-document-collection-and-save-each-part-using-documentpartsavingcall.cs`
  - Task: Iterate over the split Document collection and save each part using DocumentPartSavingCallback.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `implement-documentpartsavingcallback-to-assign-filenames-based-on-original-heading-text.cs`
  - Task: Implement DocumentPartSavingCallback to assign filenames based on original heading text.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `implement-documentpartsavingcallback-to-select-docx-for-even-parts-and-pdf-for-odd-parts.cs`
  - Task: Implement DocumentPartSavingCallback to select DOCX for even parts and PDF for odd parts.
  - Workflow: Split Workflow
  - Outputs: docx, doc, pdf
  - Selected engine: existing_repo
- `save-split-parts-as-pdf-files-while-preserving-original-document-styles-and-layout.cs`
  - Task: Save split parts as PDF files while preserving original document styles and layout.
  - Workflow: Split Workflow
  - Outputs: doc, pdf
  - Selected engine: mcp
- `save-split-parts-as-docx-files-preserving-original-formatting-and-page-orientation.cs`
  - Task: Save split parts as DOCX files preserving original formatting and page orientation.
  - Workflow: Split Workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `handle-exceptions-when-attempting-to-split-to-unsupported-mhtml-format.cs`
  - Task: Handle exceptions when attempting to split to unsupported MHTML format.
  - Workflow: Split Workflow
  - Outputs: html, mhtml
  - Selected engine: mcp
- `after-splitting-open-each-output-document-programmatically-to-verify-headers-and-footers.cs`
  - Task: After splitting, open each output document programmatically to verify headers and footers.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `merge-selected-split-documents-by-loading-them-and-using-appenddocument-to-create-combined.cs`
  - Task: Merge selected split documents by loading them and using AppendDocument to create combined file.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `process-a-batch-of-docx-files-splitting-each-by-pages-and-saving-pdfs-to-a-folder.cs`
  - Task: Process a batch of DOCX files, splitting each by pages and saving PDFs to a folder.
  - Workflow: Split Workflow
  - Outputs: docx, doc, pdf
  - Selected engine: mcp
- `split-an-epub-source-into-chapters-and-save-each-chapter-as-html-preserving-layout.cs`
  - Task: Split an EPUB source into chapters and save each chapter as HTML preserving layout.
  - Workflow: Split Workflow
  - Outputs: html, epub
  - Selected engine: mcp
- `split-an-html-source-into-chapters-and-save-each-as-docx-while-preserving-inline-styles.cs`
  - Task: Split an HTML source into chapters and save each as DOCX while preserving inline styles.
  - Workflow: Split Workflow
  - Outputs: docx, doc, html
  - Selected engine: mcp
- `split-a-document-by-custom-page-ranges-like-1-3-5-7-and-export-each-range-as-pdf.cs`
  - Task: Split a document by custom page ranges like 1-3,5-7 and export each range as PDF.
  - Workflow: Split Workflow
  - Outputs: doc, pdf
  - Selected engine: mcp
- `split-a-large-word-file-into-50-page-chunks-and-save-each-chunk-as-pdf.cs`
  - Task: Split a large Word file into 50-page chunks and save each chunk as PDF.
  - Workflow: Split Workflow
  - Outputs: pdf
  - Selected engine: mcp
- `ensure-split-parts-retain-complete-table-rows-when-original-document-contains-spanning-tab.cs`
  - Task: Ensure split parts retain complete table rows when original document contains spanning tables.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `retain-original-page-orientation-for-each-split-part-preserving-landscape-pages.cs`
  - Task: Retain original page orientation for each split part, preserving landscape pages.
  - Workflow: Split Workflow
  - Outputs: docx
  - Selected engine: mcp
- `load-a-source-document-define-split-criteria-and-execute-split-operation-in-a-single-workf.cs`
  - Task: Load a source document, define split criteria, and execute split operation in a single workflow.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-documentsplitcriteria-enumeration-to-split-by-sections-and-then-merge-selected-parts.cs`
  - Task: Use DocumentSplitCriteria enumeration to split by sections and then merge selected parts.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `apply-documentpartsavingcallback-to-customize-file-naming-convention-for-each-split-output.cs`
  - Task: Apply DocumentPartSavingCallback to customize file naming convention for each split output.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `validate-that-each-split-docx-file-maintains-original-header-and-footer-content-after-savi.cs`
  - Task: Validate that each split DOCX file maintains original header and footer content after saving.
  - Workflow: Split Workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `execute-split-operation-on-multiple-documents-sequentially-storing-results-in-designated-o.cs`
  - Task: Execute split operation on multiple documents sequentially, storing results in designated output directories.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp
- `combine-page-range-and-heading-criteria-to-produce-parts-that-start-at-each-heading-on-new.cs`
  - Task: Combine page range and heading criteria to produce parts that start at each heading on new page.
  - Workflow: Split Workflow
  - Outputs: docx
  - Selected engine: mcp
- `use-documentsplitcriteria-to-split-by-sections-then-save-each-part-to-a-network-share-loca.cs`
  - Task: Use DocumentSplitCriteria to split by sections, then save each part to a network share location.
  - Workflow: Split Workflow
  - Outputs: doc
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- Unsupported API invention
  - Symptom: Generated code references members that do not exist in the selected package version.
  - Preferred fix: Replace invented members with documented Aspose.Words APIs already proven in this category.

- Missing local bootstrap inputs
  - Symptom: The example assumes source files, folders, images, or data already exist.
  - Preferred fix: Create deterministic local inputs before loading, processing, or validating them.

- Over-broad validation
  - Symptom: The example fails at runtime while checking unrelated document internals.
  - Preferred fix: Validate only the requested behavior and the existence of expected outputs.

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
Copy-Item ..\split-document\<example-file>.cs .\Program.cs
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
