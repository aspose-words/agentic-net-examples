# Content Control Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **Content Control** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Content Control**
- Slug: **content-control**
- Total examples: **35**
- Publish-ready successful examples: **35 / 35**
- Native SDT API examples: **22**
- XML / JSON / export examples: **10**
- Input-bootstrap examples: **3**

## Category rules that shaped these examples

- Use native `StructuredDocumentTag` APIs directly with valid `SdtType` and `MarkupLevel` combinations.
- Create realistic local sample inputs inside the example whenever the task mentions streams, folders, templates, DOCX files, images, XML, or JSON.
- Use supported SDT properties, XML mapping, placeholder, and list-item APIs only.
- Use `Aspose.Drawing` instead of `System.Drawing` when drawing-related types are needed and use `Newtonsoft.Json` for JSON tasks.
- Avoid nullable-reference warnings by null-checking maybe-null results before dereference or assignment.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`
- Newtonsoft.Json

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Newtonsoft.Json

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\content-control\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `content-control/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Newtonsoft.Json

# PowerShell example
Copy-Item ..\content-control\insert-a-plain-text-content-control-at-a-specific-bookmark-in-a-docx-document.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `insert-a-plain-text-content-control-at-a-specific-bookmark-in-a-docx-document.cs` | Insert a plain text content control at a specific bookmark in a DOCX document. | input-bootstrap | docx | verified |
| 2 | `add-a-picture-content-control-that-references-an-external-image-file-and-embed-it-on-save.cs` | Add a picture content control that references an external image file and embed it on save. | native-sdt-api | docx | verified |
| 3 | `batch-process-a-folder-of-word-files-inserting-a-header-content-control-with-document-meta.cs` | Batch process a folder of Word files, inserting a header content control with document metadata. | input-bootstrap | docx | verified |
| 4 | `load-a-doc-file-add-a-date-picker-content-control-and-save-the-result-as-docx.cs` | Load a DOC file, add a date picker content control, and save the result as DOCX. | native-sdt-api | docx, doc | verified |
| 5 | `use-a-content-control-to-embed-an-ole-object-and-ensure-it-renders-correctly-after-convers.cs` | Use a content control to embed an OLE object and ensure it renders correctly after conversion to PDF. | native-sdt-api | pdf | verified |
| 6 | `use-a-content-control-to-embed-a-hyperlink-and-verify-its-target-url-after-document-conver.cs` | Use a content control to embed a hyperlink and verify its target URL after document conversion. | native-sdt-api | docx | verified |
| 7 | `create-a-repeating-section-content-control-that-repeats-a-table-row-for-each-item-in-a-col.cs` | Create a repeating section content control that repeats a table row for each item in a collection. | native-sdt-api | docx | verified |
| 8 | `create-a-content-control-that-repeats-a-paragraph-for-each-entry-in-a-json-array-during-im.cs` | Create a content control that repeats a paragraph for each entry in a JSON array during import. | xml-json-export | json | verified |
| 9 | `bind-a-dropdown-list-content-control-to-an-xml-data-source-and-populate-options-dynamicall.cs` | Bind a dropdown list content control to an XML data source and populate options dynamically. | xml-json-export | xml | verified |
| 10 | `apply-custom-xml-mapping-to-a-plain-text-content-control-to-synchronize-with-external-data.cs` | Apply custom XML mapping to a plain text content control to synchronize with external data fields. | xml-json-export | xml | verified |
| 11 | `apply-a-custom-style-to-the-text-inside-a-rich-text-content-control-programmatically.cs` | Apply a custom style to the text inside a rich text content control programmatically. | native-sdt-api | docx | verified |
| 12 | `programmatically-set-the-title-and-tag-properties-of-a-content-control-for-later-identific.cs` | Programmatically set the title and tag properties of a content control for later identification. | native-sdt-api | docx | verified |
| 13 | `update-the-tag-of-all-content-controls-in-a-document-to-follow-a-standardized-naming-conve.cs` | Update the tag of all content controls in a document to follow a standardized naming convention. | native-sdt-api | docx | verified |
| 14 | `configure-a-content-control-to-allow-only-numeric-input-and-enforce-validation-during-edit.cs` | Configure a content control to allow only numeric input and enforce validation during editing. | native-sdt-api | docx | verified |
| 15 | `validate-that-required-content-controls-contain-non-empty-text-before-saving-the-document.cs` | Validate that required content controls contain non‑empty text before saving the document. | native-sdt-api | docx | verified |
| 16 | `lock-a-content-control-to-prevent-user-editing-and-enforce-read-only-behavior-in-the-final.cs` | Lock a content control to prevent user editing and enforce read‑only behavior in the final document. | native-sdt-api | docx | verified |
| 17 | `set-the-placeholder-text-color-inside-a-content-control-to-match-the-document-theme.cs` | Set the placeholder text color inside a content control to match the document theme. | native-sdt-api | docx | verified |
| 18 | `replace-placeholder-text-in-a-content-control-with-values-from-a-dictionary-of-user-inputs.cs` | Replace placeholder text in a content control with values from a dictionary of user inputs. | native-sdt-api | docx | verified |
| 19 | `replace-the-contents-of-a-rich-text-content-control-with-formatted-html-retrieved-from-a-w.cs` | Replace the contents of a rich text content control with formatted HTML retrieved from a web service. | native-sdt-api | html | verified |
| 20 | `programmatically-clear-the-contents-of-a-content-control-without-deleting-the-control-itse.cs` | Programmatically clear the contents of a content control without deleting the control itself. | native-sdt-api | docx | verified |
| 21 | `programmatically-duplicate-a-content-control-and-insert-the-copy-at-a-different-location-i.cs` | Programmatically duplicate a content control and insert the copy at a different location in the document. | native-sdt-api | docx | verified |
| 22 | `remove-all-picture-content-controls-from-a-document-and-replace-them-with-inline-images.cs` | Remove all picture content controls from a document and replace them with inline images. | native-sdt-api | docx | verified |
| 23 | `detect-and-list-any-nested-content-controls-within-a-repeating-section-for-structural-insp.cs` | Detect and list any nested content controls within a repeating section for structural inspection. | native-sdt-api | docx | verified |
| 24 | `iterate-through-all-content-controls-in-a-document-and-generate-a-summary-report-of-their.cs` | Iterate through all content controls in a document and generate a summary report of their types. | xml-json-export | docx | verified |
| 25 | `retrieve-the-inner-xml-of-a-content-control-and-transform-it-using-an-xslt-stylesheet.cs` | Retrieve the inner XML of a content control and transform it using an XSLT stylesheet. | xml-json-export | xml | verified |
| 26 | `serialize-the-xml-mapping-of-all-content-controls-to-an-external-xsd-schema-file.cs` | Serialize the XML mapping of all content controls to an external XSD schema file. | xml-json-export | xml | verified |
| 27 | `export-the-contents-of-all-checkbox-content-controls-to-a-csv-file-for-data-analysis.cs` | Export the contents of all checkbox content controls to a CSV file for data analysis. | xml-json-export | csv | verified |
| 28 | `convert-a-docx-document-containing-content-controls-to-pdf-while-preserving-control-placeh.cs` | Convert a DOCX document containing content controls to PDF while preserving control placeholders. | input-bootstrap | pdf, docx | verified |
| 29 | `generate-a-pdf-a-compliant-document-from-a-word-file-while-keeping-content-control-tags-in.cs` | Generate a PDF/A compliant document from a Word file while keeping content control tags intact. | native-sdt-api | pdf | verified |
| 30 | `export-a-document-containing-content-controls-to-xps-format-while-preserving-control-bound.cs` | Export a document containing content controls to XPS format while preserving control boundaries. | xml-json-export | xps | verified |
| 31 | `convert-a-document-with-content-controls-to-html-while-maintaining-control-attributes-as-d.cs` | Convert a document with content controls to HTML while maintaining control attributes as data‑attributes. | native-sdt-api | html | verified |
| 32 | `merge-multiple-word-documents-preserving-existing-content-controls-and-updating-their-ids.cs` | Merge multiple Word documents, preserving existing content controls and updating their IDs. | native-sdt-api | docx | verified |
| 33 | `extract-all-repeating-section-content-controls-from-a-word-file-and-serialize-each-instanc.cs` | Extract all repeating section content controls from a Word file and serialize each instance to JSON. | xml-json-export | json | verified |
| 34 | `implement-error-handling-for-missing-xml-nodes-when-binding-data-to-a-content-control.cs` | Implement error handling for missing XML nodes when binding data to a content control. | xml-json-export | xml | verified |
| 35 | `use-a-content-control-to-store-custom-metadata-and-extract-it-for-indexing-in-a-search-eng.cs` | Use a content control to store custom metadata and extract it for indexing in a search engine. | native-sdt-api | docx | verified |

## Common failure patterns seen during generation and how they were corrected

### Invented SDT builder helpers

- Symptom: Code uses unsupported helpers such as StartStructuredDocumentTag or unsupported InsertStructuredDocumentTag overloads.
- Fix: Create StructuredDocumentTag nodes directly and insert them into valid parent nodes.

### Invalid SDT insertion location

- Symptom: Runtime error such as 'Cannot insert a node of this type at this location'.
- Fix: Use valid SdtType and MarkupLevel combinations and insert SDTs into supported block or inline containers.

### Wrong drawing or JSON library

- Symptom: Examples use System.Drawing or unsupported JSON serialization assumptions.
- Fix: Use Aspose.Drawing when needed and Newtonsoft.Json for JSON tasks.

### Invented repeating-section members

- Symptom: Code assumes convenience members such as SdtContent or unsupported repeating-section helpers.
- Fix: Enumerate actual StructuredDocumentTag nodes and inspect their child nodes through the normal document tree.

### Nullable warnings

- Symptom: CS8600, CS8602, or CS8604 around maybe-null nodes, paragraphs, or lookups.
- Fix: Use nullable locals and guard maybe-null values before dereference or assignment.

## Notes for maintainers

- This category is now **100% publish-ready** for the current run.
- Preserve file-to-task traceability when updating this folder.
- For future updates, keep the examples standalone and continue bootstrapping local inputs inside the example whenever external sources are mentioned.
