---
name: content-control
description: Verified C# examples for content control scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Content Control

## Purpose

This folder is a **live, curated example set** for content control scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free structured document tag workflows using direct Aspose.Words APIs.

## Non-negotiable conventions

- Use `StructuredDocumentTag` APIs directly with valid `SdtType` and `MarkupLevel` combinations.
- Do not use invented SDT builder helpers or unsupported insertion overloads.
- Bootstrap local DOCX, XML, image, stream, and folder inputs whenever the task mentions them.
- Use `Newtonsoft.Json` for JSON tasks and `Aspose.Drawing` instead of `System.Drawing` when drawing-related types are needed.
- Guard maybe-null values to avoid nullable-reference warnings such as `CS8600`, `CS8602`, and `CS8604`.

## Recommended workflow selection

- **Native SDT API workflow**: 22 examples
- **XML / JSON / export workflow**: 10 examples
- **Input-bootstrap workflow**: 3 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. StructuredDocumentTag operations must use real SDT nodes with valid insertion locations and supported properties.
3. Exported outputs (DOCX/JSON/XML/etc.) must actually be written by the example when required.
4. XML mapping, placeholder, and nested content scenarios must build the necessary local XML or document structure first.
5. Examples that depend on files, folders, streams, XML parts, or images should bootstrap those inputs locally during the example run.

## File-to-task reference

- `insert-a-plain-text-content-control-at-a-specific-bookmark-in-a-docx-document.cs`
  - Task: Insert a plain text content control at a specific bookmark in a DOCX document.
  - Workflow: input-bootstrap
  - Outputs: docx
  - Selected engine: verified
- `add-a-picture-content-control-that-references-an-external-image-file-and-embed-it-on-save.cs`
  - Task: Add a picture content control that references an external image file and embed it on save.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `batch-process-a-folder-of-word-files-inserting-a-header-content-control-with-document-meta.cs`
  - Task: Batch process a folder of Word files, inserting a header content control with document metadata.
  - Workflow: input-bootstrap
  - Outputs: docx
  - Selected engine: verified
- `load-a-doc-file-add-a-date-picker-content-control-and-save-the-result-as-docx.cs`
  - Task: Load a DOC file, add a date picker content control, and save the result as DOCX.
  - Workflow: native-sdt-api
  - Outputs: docx, doc
  - Selected engine: verified
- `use-a-content-control-to-embed-an-ole-object-and-ensure-it-renders-correctly-after-convers.cs`
  - Task: Use a content control to embed an OLE object and ensure it renders correctly after conversion to PDF.
  - Workflow: native-sdt-api
  - Outputs: pdf
  - Selected engine: verified
- `use-a-content-control-to-embed-a-hyperlink-and-verify-its-target-url-after-document-conver.cs`
  - Task: Use a content control to embed a hyperlink and verify its target URL after document conversion.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `create-a-repeating-section-content-control-that-repeats-a-table-row-for-each-item-in-a-col.cs`
  - Task: Create a repeating section content control that repeats a table row for each item in a collection.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `create-a-content-control-that-repeats-a-paragraph-for-each-entry-in-a-json-array-during-im.cs`
  - Task: Create a content control that repeats a paragraph for each entry in a JSON array during import.
  - Workflow: xml-json-export
  - Outputs: json
  - Selected engine: verified
- `bind-a-dropdown-list-content-control-to-an-xml-data-source-and-populate-options-dynamicall.cs`
  - Task: Bind a dropdown list content control to an XML data source and populate options dynamically.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: verified
- `apply-custom-xml-mapping-to-a-plain-text-content-control-to-synchronize-with-external-data.cs`
  - Task: Apply custom XML mapping to a plain text content control to synchronize with external data fields.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: verified
- `apply-a-custom-style-to-the-text-inside-a-rich-text-content-control-programmatically.cs`
  - Task: Apply a custom style to the text inside a rich text content control programmatically.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `programmatically-set-the-title-and-tag-properties-of-a-content-control-for-later-identific.cs`
  - Task: Programmatically set the title and tag properties of a content control for later identification.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `update-the-tag-of-all-content-controls-in-a-document-to-follow-a-standardized-naming-conve.cs`
  - Task: Update the tag of all content controls in a document to follow a standardized naming convention.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `configure-a-content-control-to-allow-only-numeric-input-and-enforce-validation-during-edit.cs`
  - Task: Configure a content control to allow only numeric input and enforce validation during editing.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `validate-that-required-content-controls-contain-non-empty-text-before-saving-the-document.cs`
  - Task: Validate that required content controls contain non‑empty text before saving the document.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `lock-a-content-control-to-prevent-user-editing-and-enforce-read-only-behavior-in-the-final.cs`
  - Task: Lock a content control to prevent user editing and enforce read‑only behavior in the final document.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `set-the-placeholder-text-color-inside-a-content-control-to-match-the-document-theme.cs`
  - Task: Set the placeholder text color inside a content control to match the document theme.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `replace-placeholder-text-in-a-content-control-with-values-from-a-dictionary-of-user-inputs.cs`
  - Task: Replace placeholder text in a content control with values from a dictionary of user inputs.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `replace-the-contents-of-a-rich-text-content-control-with-formatted-html-retrieved-from-a-w.cs`
  - Task: Replace the contents of a rich text content control with formatted HTML retrieved from a web service.
  - Workflow: native-sdt-api
  - Outputs: html
  - Selected engine: verified
- `programmatically-clear-the-contents-of-a-content-control-without-deleting-the-control-itse.cs`
  - Task: Programmatically clear the contents of a content control without deleting the control itself.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `programmatically-duplicate-a-content-control-and-insert-the-copy-at-a-different-location-i.cs`
  - Task: Programmatically duplicate a content control and insert the copy at a different location in the document.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `remove-all-picture-content-controls-from-a-document-and-replace-them-with-inline-images.cs`
  - Task: Remove all picture content controls from a document and replace them with inline images.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `detect-and-list-any-nested-content-controls-within-a-repeating-section-for-structural-insp.cs`
  - Task: Detect and list any nested content controls within a repeating section for structural inspection.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `iterate-through-all-content-controls-in-a-document-and-generate-a-summary-report-of-their.cs`
  - Task: Iterate through all content controls in a document and generate a summary report of their types.
  - Workflow: xml-json-export
  - Outputs: docx
  - Selected engine: verified
- `retrieve-the-inner-xml-of-a-content-control-and-transform-it-using-an-xslt-stylesheet.cs`
  - Task: Retrieve the inner XML of a content control and transform it using an XSLT stylesheet.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: verified
- `serialize-the-xml-mapping-of-all-content-controls-to-an-external-xsd-schema-file.cs`
  - Task: Serialize the XML mapping of all content controls to an external XSD schema file.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: verified
- `export-the-contents-of-all-checkbox-content-controls-to-a-csv-file-for-data-analysis.cs`
  - Task: Export the contents of all checkbox content controls to a CSV file for data analysis.
  - Workflow: xml-json-export
  - Outputs: csv
  - Selected engine: verified
- `convert-a-docx-document-containing-content-controls-to-pdf-while-preserving-control-placeh.cs`
  - Task: Convert a DOCX document containing content controls to PDF while preserving control placeholders.
  - Workflow: input-bootstrap
  - Outputs: pdf, docx
  - Selected engine: verified
- `generate-a-pdf-a-compliant-document-from-a-word-file-while-keeping-content-control-tags-in.cs`
  - Task: Generate a PDF/A compliant document from a Word file while keeping content control tags intact.
  - Workflow: native-sdt-api
  - Outputs: pdf
  - Selected engine: verified
- `export-a-document-containing-content-controls-to-xps-format-while-preserving-control-bound.cs`
  - Task: Export a document containing content controls to XPS format while preserving control boundaries.
  - Workflow: xml-json-export
  - Outputs: xps
  - Selected engine: verified
- `convert-a-document-with-content-controls-to-html-while-maintaining-control-attributes-as-d.cs`
  - Task: Convert a document with content controls to HTML while maintaining control attributes as data‑attributes.
  - Workflow: native-sdt-api
  - Outputs: html
  - Selected engine: verified
- `merge-multiple-word-documents-preserving-existing-content-controls-and-updating-their-ids.cs`
  - Task: Merge multiple Word documents, preserving existing content controls and updating their IDs.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified
- `extract-all-repeating-section-content-controls-from-a-word-file-and-serialize-each-instanc.cs`
  - Task: Extract all repeating section content controls from a Word file and serialize each instance to JSON.
  - Workflow: xml-json-export
  - Outputs: json
  - Selected engine: verified
- `implement-error-handling-for-missing-xml-nodes-when-binding-data-to-a-content-control.cs`
  - Task: Implement error handling for missing XML nodes when binding data to a content control.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: verified
- `use-a-content-control-to-store-custom-metadata-and-extract-it-for-indexing-in-a-search-eng.cs`
  - Task: Use a content control to store custom metadata and extract it for indexing in a search engine.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: verified

## Common failure patterns and preferred agent fixes

- **Invented SDT builder helpers**
  - Symptom: Code uses unsupported helpers such as StartStructuredDocumentTag or unsupported InsertStructuredDocumentTag overloads.
  - Preferred fix: Create StructuredDocumentTag nodes directly and insert them into valid parent nodes.

- **Invalid SDT insertion location**
  - Symptom: Runtime error such as 'Cannot insert a node of this type at this location'.
  - Preferred fix: Use valid SdtType and MarkupLevel combinations and insert SDTs into supported block or inline containers.

- **Wrong drawing or JSON library**
  - Symptom: Examples use System.Drawing or unsupported JSON serialization assumptions.
  - Preferred fix: Use Aspose.Drawing when needed and Newtonsoft.Json for JSON tasks.

- **Invented repeating-section members**
  - Symptom: Code assumes convenience members such as SdtContent or unsupported repeating-section helpers.
  - Preferred fix: Enumerate actual StructuredDocumentTag nodes and inspect their child nodes through the normal document tree.

- **Nullable warnings**
  - Symptom: CS8600, CS8602, or CS8604 around maybe-null nodes, paragraphs, or lookups.
  - Preferred fix: Use nullable locals and guard maybe-null values before dereference or assignment.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`
- Additional package: `Newtonsoft.Json`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
dotnet add package Newtonsoft.Json
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\content-control\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer direct Aspose.Words APIs over speculative shortcuts.
