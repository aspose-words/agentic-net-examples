---
name: split-document
description: Verified C# examples for split-document scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Split Document

## Purpose

Provide an AI coding assistant with concise, version-accurate guidance for implementing reliable document-splitting workflows that rely on documented Aspose.Words classes, avoid invented APIs, and include deterministic input creation and output validation.

## Non-negotiable conventions

- Never reference Aspose.Words members that are absent from version 26.10.0.
- Do not import or use namespaces that are not part of the official package.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Split Workflow workflow**: 30 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Confirm that each expected split file exists on disk after the operation.

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
  - Selected engine: llm
- `create-a-documentsplitcriteria-object-and-set-split-mode-to-custom-page-ranges.cs`
  - Task: Create a DocumentSplitCriteria object and set split mode to custom page ranges.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: llm
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
  - Selected engine: llm
- `implement-documentpartsavingcallback-to-assign-filenames-based-on-original-heading-text.cs`
  - Task: Implement DocumentPartSavingCallback to assign filenames based on original heading text.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: llm
- `implement-documentpartsavingcallback-to-select-docx-for-even-parts-and-pdf-for-odd-parts.cs`
  - Task: Implement DocumentPartSavingCallback to select DOCX for even parts and PDF for odd parts.
  - Workflow: split-workflow
  - Outputs: docx, doc, pdf
  - Selected engine: llm
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
  - Task: Split a large Word file into 50-page chunks and save each chunk as PDF.
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
  - Selected engine: llm
- `apply-documentpartsavingcallback-to-customize-file-naming-convention-for-each-split-output.cs`
  - Task: Apply DocumentPartSavingCallback to customize file naming convention for each split output.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: llm
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

- **Unsupported API usage**
  - Symptom: Code references members such as DocumentSplitCriteria.SplitMode, DocumentSplitMode, or Aspose.Words.Import that do not exist in the installed library.
  - Root cause: The example invents APIs that were never released in Aspose.Words 26.10.0.
  - Preferred fix: Replace invented members with documented alternatives, for example HtmlSaveOptions.DocumentSplitCriteria and HtmlSaveOptions.DocumentSplitHeadingLevel for heading-based splitting.
  - Prevention: Consult the official Aspose.Words API reference for the exact class and member names before writing code.

- **Invalid page range arguments**
  - Symptom: ExtractPages throws ArgumentOutOfRangeException because the start or end page index exceeds the document's page count.
  - Root cause: Page numbers are calculated incorrectly or zero-based indices are mixed with one-based values.
  - Preferred fix: Call Document.UpdatePageLayout() first, verify that 1 ≤ startPage ≤ endPage ≤ doc.PageCount, then invoke ExtractPages(startPage-1, pageCount).
  - Prevention: Always validate page numbers against the actual PageCount and use consistent zero-based indexing.

- **Cross-document node import errors**
  - Symptom: Runtime exception reports that a new child node was created from a different document than the target.
  - Root cause: Nodes are appended directly without using NodeImporter or cloning, violating Aspose.Words ownership rules.
  - Preferred fix: Import nodes with NodeImporter (ImportFormatMode.KeepSourceFormatting) or clone them before insertion.
  - Prevention: Never mix nodes from different Document instances without proper import or cloning.

- **Missing required namespaces or types**
  - Symptom: CS0234 errors indicate that namespaces such as Aspose.Words.Import, Aspose.Words.Building, or Aspose.Words.Splitting cannot be found.
  - Root cause: The code assumes the existence of non-existent sub-namespaces.
  - Preferred fix: Remove invalid using directives and reference only the namespaces listed in the package documentation (e.g., Aspose.Words, Aspose.Words.Saving).
  - Prevention: Rely on IDE intellisense or the official documentation to determine available namespaces.

- **Incorrect DocumentPartSavingCallback usage**
  - Symptom: CS1061 errors show that SaveOptions does not contain a DocumentPartSavingCallback property in the current version.
  - Root cause: The example attempts to assign a callback to a property that was removed or never existed.
  - Preferred fix: Omit the callback or implement IDocumentPartSavingCallback only where the property is supported, using the correct SaveOptions subclass.
  - Prevention: Verify the presence of callback properties in the specific Aspose.Words version before using them.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.10.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
```

```powershell
Copy-Item ..\split-document\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Review and update examples whenever a new Aspose.Words version adds or deprecates splitting-related members.
