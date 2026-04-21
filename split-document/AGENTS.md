---
name: split-document
description: Verified C# examples for document splitting workflows in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Split Document

## Purpose

This folder is a **live, curated example set** for split-document scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free use of documented Aspose.Words APIs for splitting by sections, pages, headings, bookmarks, ranges, and related extraction boundaries.

## Non-negotiable conventions

- Always use documented Aspose.Words APIs directly.
- Always create local sample source documents when a task refers to an existing file, folder, stream, template, or input asset.
- Prefer documented split, extraction, import, and clone workflows.
- Keep validation narrow and task-specific.
- Do not invent split, page-extraction, or import helper APIs.

## Recommended workflow selection

- **Split workflow**: 30 examples

This category performed best with light primary rules plus a narrow hardening patch for page-range and preservation-sensitive tasks.

## Validation priorities

1. The code must compile and run without manual input.
2. Required sample inputs must be bootstrapped locally inside the example.
3. Requested split output files must be created successfully.
4. Validation should focus only on the exact split boundary, page range, preserved content, or output count requested by the task.

## File-to-task reference

- `load-a-docx-source-document-using-the-document-class-before-splitting.cs`
  - Task: Load a DOCX source document using the Document class before splitting.
  - Workflow: split-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `create-a-documentsplitcriteria-object-and-set-split-mode-to-headings.cs`
  - Task: Create a DocumentSplitCriteria object and set split mode to headings.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-documentsplitcriteria-object-and-set-split-mode-to-sections.cs`
  - Task: Create a DocumentSplitCriteria object and set split mode to sections.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-documentsplitcriteria-object-and-set-split-mode-to-individual-pages.cs`
  - Task: Create a DocumentSplitCriteria object and set split mode to individual pages.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `create-a-documentsplitcriteria-object-and-set-split-mode-to-custom-page-ranges.cs`
  - Task: Create a DocumentSplitCriteria object and set split mode to custom page ranges.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `combine-heading-and-section-flags-in-documentsplitcriteria-to-split-by-both-structures.cs`
  - Task: Combine heading and section flags in DocumentSplitCriteria to split by both structures.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `combine-page-and-heading-flags-in-documentsplitcriteria-to-start-each-part-on-a-new-page.cs`
  - Task: Combine page and heading flags in DocumentSplitCriteria to start each part on a new page.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `call-document-split-criteria-to-obtain-a-collection-of-split-document-objects.cs`
  - Task: Call Document.Split(criteria) to obtain a collection of split Document objects.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `iterate-over-the-split-document-collection-and-save-each-part-using-documentpartsavingcall.cs`
  - Task: Iterate over the split Document collection and save each part using DocumentPartSavingCallback.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `implement-documentpartsavingcallback-to-assign-filenames-based-on-original-heading-text.cs`
  - Task: Implement DocumentPartSavingCallback to assign filenames based on original heading text.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `implement-documentpartsavingcallback-to-select-docx-for-even-parts-and-pdf-for-odd-parts.cs`
  - Task: Implement DocumentPartSavingCallback to select DOCX for even parts and PDF for odd parts.
  - Workflow: split-workflow
  - Outputs: docx, doc, pdf
  - Selected engine: mcp
- `save-split-parts-as-pdf-files-while-preserving-original-document-styles-and-layout.cs`
  - Task: Save split parts as PDF files while preserving original document styles and layout.
  - Workflow: split-workflow
  - Outputs: doc, pdf
  - Selected engine: mcp
- `save-split-parts-as-docx-files-preserving-original-formatting-and-page-orientation.cs`
  - Task: Save split parts as DOCX files preserving original formatting and page orientation.
  - Workflow: split-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `handle-exceptions-when-attempting-to-split-to-unsupported-mhtml-format.cs`
  - Task: Handle exceptions when attempting to split to unsupported MHTML format.
  - Workflow: split-workflow
  - Outputs: html, mhtml
  - Selected engine: mcp
- `after-splitting-open-each-output-document-programmatically-to-verify-headers-and-footers.cs`
  - Task: After splitting, open each output document programmatically to verify headers and footers.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `merge-selected-split-documents-by-loading-them-and-using-appenddocument-to-create-combined.cs`
  - Task: Merge selected split documents by loading them and using AppendDocument to create combined file.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `process-a-batch-of-docx-files-splitting-each-by-pages-and-saving-pdfs-to-a-folder.cs`
  - Task: Process a batch of DOCX files, splitting each by pages and saving PDFs to a folder.
  - Workflow: split-workflow
  - Outputs: docx, doc, pdf
  - Selected engine: mcp
- `split-an-epub-source-into-chapters-and-save-each-chapter-as-html-preserving-layout.cs`
  - Task: Split an EPUB source into chapters and save each chapter as HTML preserving layout.
  - Workflow: split-workflow
  - Outputs: html, epub
  - Selected engine: mcp
- `split-an-html-source-into-chapters-and-save-each-as-docx-while-preserving-inline-styles.cs`
  - Task: Split an HTML source into chapters and save each as DOCX while preserving inline styles.
  - Workflow: split-workflow
  - Outputs: docx, doc, html
  - Selected engine: mcp
- `split-a-document-by-custom-page-ranges-like-1-3-5-7-and-export-each-range-as-pdf.cs`
  - Task: Split a document by custom page ranges like 1-3,5-7 and export each range as PDF.
  - Workflow: split-workflow
  - Outputs: doc, pdf
  - Selected engine: mcp
- `split-a-large-word-file-into-50-page-chunks-and-save-each-chunk-as-pdf.cs`
  - Task: Split a large Word file into 50‑page chunks and save each chunk as PDF.
  - Workflow: split-workflow
  - Outputs: pdf
  - Selected engine: mcp
- `ensure-split-parts-retain-complete-table-rows-when-original-document-contains-spanning-tab.cs`
  - Task: Ensure split parts retain complete table rows when original document contains spanning tables.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `retain-original-page-orientation-for-each-split-part-preserving-landscape-pages.cs`
  - Task: Retain original page orientation for each split part, preserving landscape pages.
  - Workflow: split-workflow
  - Outputs: docx
  - Selected engine: mcp
- `load-a-source-document-define-split-criteria-and-execute-split-operation-in-a-single-workf.cs`
  - Task: Load a source document, define split criteria, and execute split operation in a single workflow.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-documentsplitcriteria-enumeration-to-split-by-sections-and-then-merge-selected-parts.cs`
  - Task: Use DocumentSplitCriteria enumeration to split by sections and then merge selected parts.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `apply-documentpartsavingcallback-to-customize-file-naming-convention-for-each-split-output.cs`
  - Task: Apply DocumentPartSavingCallback to customize file naming convention for each split output.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `validate-that-each-split-docx-file-maintains-original-header-and-footer-content-after-savi.cs`
  - Task: Validate that each split DOCX file maintains original header and footer content after saving.
  - Workflow: split-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `execute-split-operation-on-multiple-documents-sequentially-storing-results-in-designated-o.cs`
  - Task: Execute split operation on multiple documents sequentially, storing results in designated output directories.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp
- `combine-page-range-and-heading-criteria-to-produce-parts-that-start-at-each-heading-on-new.cs`
  - Task: Combine page range and heading criteria to produce parts that start at each heading on new page.
  - Workflow: split-workflow
  - Outputs: docx
  - Selected engine: mcp
- `use-documentsplitcriteria-to-split-by-sections-then-save-each-part-to-a-network-share-loca.cs`
  - Task: Use DocumentSplitCriteria to split by sections, then save each part to a network share location.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Invented split APIs**
  - Symptom: Build failures caused by unsupported members such as `Document.Split(...)`, `DocumentPageSplitter`, or wrong save-option properties.
  - Preferred fix: Use only documented page extraction, section extraction, import, clone, and callback workflows proven in the current package version.
- **Cross-document node insertion mistakes**
  - Symptom: Runtime failures when content extracted from one document is appended directly into another.
  - Preferred fix: Import, clone, or use `NodeImporter` before appending nodes to a destination document.
- **Header and footer preservation failures**
  - Symptom: Split files save correctly but reopen without the expected header or footer content.
  - Preferred fix: Preserve section structure during extraction and validate reopened DOCX outputs rather than brittle text-only checks.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required package

```bash
dotnet add package Aspose.Words --version 26.3.0
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

- Preserve exact file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the verified winner from the latest batch report rather than a merely compiling draft.
- Bootstrap file-based inputs locally instead of depending on machine-specific paths.
