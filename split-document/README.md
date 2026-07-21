# Split-Document Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Split-Document category. Each file is a standalone console example selected from the verified 26.6.0 run.

## Snapshot

- Category: Split-Document
- Slug: split-document
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20260711_192617_b9179d
- Split Workflow examples: 30

## Category rules that shaped these examples

- Use native Aspose.Words APIs directly.
- Create local sample source documents when a task refers to an existing file, folder, stream, or template.
- Do not assume external files or folders already exist.
- Prefer documented section, node, heading, bookmark, and page extraction workflows only when relevant to the task.
- Keep validation narrow and task-specific.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.6.0

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.6.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\split-document\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `split-document/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.6.0

# PowerShell example
Copy-Item ..\split-document\load-a-docx-source-document-using-the-document-class-before-splitting.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-source-document-using-the-document-class-before-splitting.cs` | Load a DOCX source document using the Document class before splitting. | split-workflow | docx, doc | mcp |
| 2 | `create-a-documentsplitcriteria-object-and-set-split-mode-to-headings.cs` | Create a DocumentSplitCriteria object and set split mode to headings. | split-workflow | doc | mcp |
| 3 | `create-a-documentsplitcriteria-object-and-set-split-mode-to-sections.cs` | Create a DocumentSplitCriteria object and set split mode to sections. | split-workflow | doc | mcp |
| 4 | `create-a-documentsplitcriteria-object-and-set-split-mode-to-individual-pages.cs` | Create a DocumentSplitCriteria object and set split mode to individual pages. | split-workflow | doc | mcp |
| 5 | `create-a-documentsplitcriteria-object-and-set-split-mode-to-custom-page-ranges.cs` | Create a DocumentSplitCriteria object and set split mode to custom page ranges. | split-workflow | doc | mcp |
| 6 | `combine-heading-and-section-flags-in-documentsplitcriteria-to-split-by-both-structures.cs` | Combine heading and section flags in DocumentSplitCriteria to split by both structures. | split-workflow | doc | mcp |
| 7 | `combine-page-and-heading-flags-in-documentsplitcriteria-to-start-each-part-on-a-new-page.cs` | Combine page and heading flags in DocumentSplitCriteria to start each part on a new page. | split-workflow | doc | mcp |
| 8 | `call-document-split-criteria-to-obtain-a-collection-of-split-document-objects.cs` | Call Document.Split(criteria) to obtain a collection of split Document objects. | split-workflow | doc | mcp |
| 9 | `iterate-over-the-split-document-collection-and-save-each-part-using-documentpartsavingcall.cs` | Iterate over the split Document collection and save each part using DocumentPartSavingCallback. | split-workflow | doc | mcp |
| 10 | `implement-documentpartsavingcallback-to-assign-filenames-based-on-original-heading-text.cs` | Implement DocumentPartSavingCallback to assign filenames based on original heading text. | split-workflow | doc | mcp |
| 11 | `implement-documentpartsavingcallback-to-select-docx-for-even-parts-and-pdf-for-odd-parts.cs` | Implement DocumentPartSavingCallback to select DOCX for even parts and PDF for odd parts. | split-workflow | docx, doc, pdf | mcp |
| 12 | `save-split-parts-as-pdf-files-while-preserving-original-document-styles-and-layout.cs` | Save split parts as PDF files while preserving original document styles and layout. | split-workflow | doc, pdf | mcp |
| 13 | `save-split-parts-as-docx-files-preserving-original-formatting-and-page-orientation.cs` | Save split parts as DOCX files preserving original formatting and page orientation. | split-workflow | docx, doc | mcp |
| 14 | `handle-exceptions-when-attempting-to-split-to-unsupported-mhtml-format.cs` | Handle exceptions when attempting to split to unsupported MHTML format. | split-workflow | html, mhtml | mcp |
| 15 | `after-splitting-open-each-output-document-programmatically-to-verify-headers-and-footers.cs` | After splitting, open each output document programmatically to verify headers and footers. | split-workflow | doc | mcp |
| 16 | `merge-selected-split-documents-by-loading-them-and-using-appenddocument-to-create-combined.cs` | Merge selected split documents by loading them and using AppendDocument to create combined file. | split-workflow | doc | mcp |
| 17 | `process-a-batch-of-docx-files-splitting-each-by-pages-and-saving-pdfs-to-a-folder.cs` | Process a batch of DOCX files, splitting each by pages and saving PDFs to a folder. | split-workflow | docx, doc, pdf | mcp |
| 18 | `split-an-epub-source-into-chapters-and-save-each-chapter-as-html-preserving-layout.cs` | Split an EPUB source into chapters and save each chapter as HTML preserving layout. | split-workflow | html, epub | mcp |
| 19 | `split-an-html-source-into-chapters-and-save-each-as-docx-while-preserving-inline-styles.cs` | Split an HTML source into chapters and save each as DOCX while preserving inline styles. | split-workflow | docx, doc, html | llm |
| 20 | `split-a-document-by-custom-page-ranges-like-1-3-5-7-and-export-each-range-as-pdf.cs` | Split a document by custom page ranges like 1-3,5-7 and export each range as PDF. | split-workflow | doc, pdf | mcp |
| 21 | `split-a-large-word-file-into-50-page-chunks-and-save-each-chunk-as-pdf.cs` | Split a large Word file into 50-page chunks and save each chunk as PDF. | split-workflow | pdf | mcp |
| 22 | `ensure-split-parts-retain-complete-table-rows-when-original-document-contains-spanning-tab.cs` | Ensure split parts retain complete table rows when original document contains spanning tables. | split-workflow | doc | mcp |
| 23 | `retain-original-page-orientation-for-each-split-part-preserving-landscape-pages.cs` | Retain original page orientation for each split part, preserving landscape pages. | split-workflow | docx | mcp |
| 24 | `load-a-source-document-define-split-criteria-and-execute-split-operation-in-a-single-workf.cs` | Load a source document, define split criteria, and execute split operation in a single workflow. | split-workflow | doc | mcp |
| 25 | `use-documentsplitcriteria-enumeration-to-split-by-sections-and-then-merge-selected-parts.cs` | Use DocumentSplitCriteria enumeration to split by sections and then merge selected parts. | split-workflow | doc | llm |
| 26 | `apply-documentpartsavingcallback-to-customize-file-naming-convention-for-each-split-output.cs` | Apply DocumentPartSavingCallback to customize file naming convention for each split output. | split-workflow | doc | mcp |
| 27 | `validate-that-each-split-docx-file-maintains-original-header-and-footer-content-after-savi.cs` | Validate that each split DOCX file maintains original header and footer content after saving. | split-workflow | docx, doc | existing_repo |
| 28 | `execute-split-operation-on-multiple-documents-sequentially-storing-results-in-designated-o.cs` | Execute split operation on multiple documents sequentially, storing results in designated output directories. | split-workflow | doc | llm |
| 29 | `combine-page-range-and-heading-criteria-to-produce-parts-that-start-at-each-heading-on-new.cs` | Combine page range and heading criteria to produce parts that start at each heading on new page. | split-workflow | docx | mcp |
| 30 | `use-documentsplitcriteria-to-split-by-sections-then-save-each-part-to-a-network-share-loca.cs` | Use DocumentSplitCriteria to split by sections, then save each part to a network share location. | split-workflow | doc | mcp |

## Common failure patterns seen during generation and how they were corrected

### Unsupported API invention

- Symptom: Generated code references members that do not exist in the selected package version.
- Fix: Replace invented members with documented Aspose.Words APIs already proven in this category.

### Missing local bootstrap inputs

- Symptom: The example assumes source files, folders, images, or data already exist.
- Fix: Create deterministic local inputs before loading, processing, or validating them.

### Over-broad validation

- Symptom: The example fails at runtime while checking unrelated document internals.
- Fix: Validate only the requested behavior and the existence of expected outputs.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.6.0 run.
- Preserve file-to-task traceability when updating this folder.
- Keep examples standalone and bootstrap local inputs inside the example whenever external sources are mentioned.
