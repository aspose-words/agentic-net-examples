---
name: split-document
description: Verified C# examples for split-document scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Split Document

## Purpose

Equip an AI coding assistant with concise, reproducible patterns for splitting documents by sections, headings, pages, or custom ranges, and for saving each part with optional callbacks, ensuring full compatibility with the documented Aspose.Words API.

## Non-negotiable conventions

- Do not reference members, types, or namespaces that are not present in Aspose.Words 26.11.0.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Split Workflow workflow**: 30 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Prioritize confirming that each expected output file exists and that its content matches the intended split scope.

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
  - Selected engine: mcp
- `combine-heading-and-section-flags-in-documentsplitcriteria-to-split-by-both-structures.cs`
  - Task: Combine heading and section flags in DocumentSplitCriteria to split by both structures.
  - Workflow: split-workflow
  - Outputs: doc
  - Selected engine: existing_repo
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
  - Selected engine: llm
- `save-split-parts-as-pdf-files-while-preserving-original-document-styles-and-layout.cs`
  - Task: Save split parts as PDF files while preserving original document styles and layout.
  - Workflow: split-workflow
  - Outputs: doc, pdf
  - Selected engine: llm
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
  - Selected engine: llm
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

- **Unsupported API usage**
  - Symptom: The code references types or members such as DocumentSplitCriteria.Heading, Importing namespace, or SplitDocumentOptions that do not exist in the selected package version.
  - Root cause: Invented or outdated API names were used instead of the documented members available in Aspose.Words 26.11.0.
  - Preferred fix: Replace the nonexistent members with the correct documented equivalents, e.g., use HtmlSaveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph and remove any Importing namespace references.
  - Prevention: Consult the official Aspose.Words API reference for the exact member names before writing split logic.

- **Invalid page range**
  - Symptom: Runtime throws ArgumentOutOfRangeException indicating that a page index or end index is out of range.
  - Root cause: Page-based splitting used zero-based indices without adjusting for the document's actual page count, or supplied ranges exceeded the document's total pages.
  - Preferred fix: Validate the document's PageCount and ensure start indices are zero-based while end indices do not exceed PageCount; adjust custom ranges accordingly.
  - Prevention: Always compute and check page boundaries before calling Document.ExtractPages.

- **Cross-document node insertion without import**
  - Symptom: Runtime throws ArgumentException or InvalidOperationException when inserting a node created from a different Document instance.
  - Root cause: Nodes were appended directly to a target document without using NodeImporter, violating Aspose.Words' ownership rules.
  - Preferred fix: Import each node with NodeImporter (ImportFormatMode.KeepSourceFormatting) before appending it to the destination document.
  - Prevention: Never mix nodes from different Document objects without explicit import.

- **Uninitialized object**
  - Symptom: NullReferenceException occurs during split processing, typically when accessing a DocumentBuilder, Section, or callback that was never instantiated.
  - Root cause: The workflow omitted required object creation steps or relied on external files that were not generated locally.
  - Preferred fix: Instantiate all required objects explicitly and create any sample source files before loading or processing them.
  - Prevention: Follow the rule to create local bootstrap inputs for every task that references a file, folder, or stream.

- **Incorrect callback implementation**
  - Symptom: Compilation fails because IDocumentPartSavingCallback or related members are missing or mismatched, or runtime errors occur when the callback is invoked.
  - Root cause: The example attempted to assign a lambda or set a property that does not exist on SaveOptions, or the callback class did not implement the required interface members.
  - Preferred fix: Define a class that implements IDocumentPartSavingCallback with the DocumentPartSaving method, and pass an instance to the appropriate SaveOptions property (e.g., HtmlSaveOptions.DocumentPartSavingCallback).
  - Prevention: Use the documented callback pattern and verify the interface signature against the Aspose.Words version in use.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
```

```powershell
Copy-Item ..\split-document\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update the examples whenever a new Aspose.Words version adds, renames, or deprecates split-related members to keep the documentation current.
