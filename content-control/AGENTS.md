---
name: content-control
description: Verified C# examples for content-control scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Content Control

## Purpose

Provide an AI coding assistant with clear, standards-compliant patterns for creating, configuring, and processing content controls in Aspose.Words, ensuring that generated code compiles, runs, and respects the library's API surface.

## Non-negotiable conventions

- Do not reference Aspose.Words members that do not exist in the current version, such as SdtType.DateTime or StructuredDocumentTag.DateValue.
- Avoid invented helper methods like StartStructuredDocumentTag or InsertStructuredDocumentTag overloads that are not part of the API.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Input Bootstrap workflow**: 3 examples
- **Native Sdt API workflow**: 22 examples
- **Xml Json Export workflow**: 10 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Ensure that every API call matches the signature defined in Aspose.Words 26.11.0; compilation errors are the highest priority to address.

## File-to-task reference

- `insert-a-plain-text-content-control-at-a-specific-bookmark-in-a-docx-document.cs`
  - Task: Insert a plain text content control at a specific bookmark in a DOCX document.
  - Workflow: input-bootstrap
  - Outputs: docx
  - Selected engine: mcp
- `add-a-picture-content-control-that-references-an-external-image-file-and-embed-it-on-save.cs`
  - Task: Add a picture content control that references an external image file and embed it on save.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `batch-process-a-folder-of-word-files-inserting-a-header-content-control-with-document-meta.cs`
  - Task: Batch process a folder of Word files, inserting a header content control with document metadata.
  - Workflow: input-bootstrap
  - Outputs: docx
  - Selected engine: mcp
- `load-a-doc-file-add-a-date-picker-content-control-and-save-the-result-as-docx.cs`
  - Task: Load a DOC file, add a date picker content control, and save the result as DOCX.
  - Workflow: native-sdt-api
  - Outputs: docx, doc
  - Selected engine: mcp
- `use-a-content-control-to-embed-an-ole-object-and-ensure-it-renders-correctly-after-convers.cs`
  - Task: Use a content control to embed an OLE object and ensure it renders correctly after conversion to PDF.
  - Workflow: native-sdt-api
  - Outputs: pdf
  - Selected engine: existing_repo
- `use-a-content-control-to-embed-a-hyperlink-and-verify-its-target-url-after-document-conver.cs`
  - Task: Use a content control to embed a hyperlink and verify its target URL after document conversion.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: llm
- `create-a-repeating-section-content-control-that-repeats-a-table-row-for-each-item-in-a-col.cs`
  - Task: Create a repeating section content control that repeats a table row for each item in a collection.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `create-a-content-control-that-repeats-a-paragraph-for-each-entry-in-a-json-array-during-im.cs`
  - Task: Create a content control that repeats a paragraph for each entry in a JSON array during import.
  - Workflow: xml-json-export
  - Outputs: json
  - Selected engine: mcp
- `bind-a-dropdown-list-content-control-to-an-xml-data-source-and-populate-options-dynamicall.cs`
  - Task: Bind a dropdown list content control to an XML data source and populate options dynamically.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: existing_repo
- `apply-custom-xml-mapping-to-a-plain-text-content-control-to-synchronize-with-external-data.cs`
  - Task: Apply custom XML mapping to a plain text content control to synchronize with external data fields.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: existing_repo
- `apply-a-custom-style-to-the-text-inside-a-rich-text-content-control-programmatically.cs`
  - Task: Apply a custom style to the text inside a rich text content control programmatically.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-set-the-title-and-tag-properties-of-a-content-control-for-later-identific.cs`
  - Task: Programmatically set the title and tag properties of a content control for later identification.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `update-the-tag-of-all-content-controls-in-a-document-to-follow-a-standardized-naming-conve.cs`
  - Task: Update the tag of all content controls in a document to follow a standardized naming convention.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `configure-a-content-control-to-allow-only-numeric-input-and-enforce-validation-during-edit.cs`
  - Task: Configure a content control to allow only numeric input and enforce validation during editing.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: llm
- `validate-that-required-content-controls-contain-non-empty-text-before-saving-the-document.cs`
  - Task: Validate that required content controls contain non-empty text before saving the document.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `lock-a-content-control-to-prevent-user-editing-and-enforce-read-only-behavior-in-the-final.cs`
  - Task: Lock a content control to prevent user editing and enforce read-only behavior in the final document.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `set-the-placeholder-text-color-inside-a-content-control-to-match-the-document-theme.cs`
  - Task: Set the placeholder text color inside a content control to match the document theme.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: existing_repo
- `replace-placeholder-text-in-a-content-control-with-values-from-a-dictionary-of-user-inputs.cs`
  - Task: Replace placeholder text in a content control with values from a dictionary of user inputs.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `replace-the-contents-of-a-rich-text-content-control-with-formatted-html-retrieved-from-a-w.cs`
  - Task: Replace the contents of a rich text content control with formatted HTML retrieved from a web service.
  - Workflow: native-sdt-api
  - Outputs: html
  - Selected engine: mcp
- `programmatically-clear-the-contents-of-a-content-control-without-deleting-the-control-itse.cs`
  - Task: Programmatically clear the contents of a content control without deleting the control itself.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-duplicate-a-content-control-and-insert-the-copy-at-a-different-location-i.cs`
  - Task: Programmatically duplicate a content control and insert the copy at a different location in the document.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `remove-all-picture-content-controls-from-a-document-and-replace-them-with-inline-images.cs`
  - Task: Remove all picture content controls from a document and replace them with inline images.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `detect-and-list-any-nested-content-controls-within-a-repeating-section-for-structural-insp.cs`
  - Task: Detect and list any nested content controls within a repeating section for structural inspection.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `iterate-through-all-content-controls-in-a-document-and-generate-a-summary-report-of-their.cs`
  - Task: Iterate through all content controls in a document and generate a summary report of their types.
  - Workflow: xml-json-export
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-the-inner-xml-of-a-content-control-and-transform-it-using-an-xslt-stylesheet.cs`
  - Task: Retrieve the inner XML of a content control and transform it using an XSLT stylesheet.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: mcp
- `serialize-the-xml-mapping-of-all-content-controls-to-an-external-xsd-schema-file.cs`
  - Task: Serialize the XML mapping of all content controls to an external XSD schema file.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: existing_repo
- `export-the-contents-of-all-checkbox-content-controls-to-a-csv-file-for-data-analysis.cs`
  - Task: Export the contents of all checkbox content controls to a CSV file for data analysis.
  - Workflow: xml-json-export
  - Outputs: csv
  - Selected engine: mcp
- `convert-a-docx-document-containing-content-controls-to-pdf-while-preserving-control-placeh.cs`
  - Task: Convert a DOCX document containing content controls to PDF while preserving control placeholders.
  - Workflow: input-bootstrap
  - Outputs: pdf, docx
  - Selected engine: llm
- `generate-a-pdf-a-compliant-document-from-a-word-file-while-keeping-content-control-tags-in.cs`
  - Task: Generate a PDF/A compliant document from a Word file while keeping content control tags intact.
  - Workflow: native-sdt-api
  - Outputs: pdf
  - Selected engine: mcp
- `export-a-document-containing-content-controls-to-xps-format-while-preserving-control-bound.cs`
  - Task: Export a document containing content controls to XPS format while preserving control boundaries.
  - Workflow: xml-json-export
  - Outputs: xps
  - Selected engine: mcp
- `convert-a-document-with-content-controls-to-html-while-maintaining-control-attributes-as-d.cs`
  - Task: Convert a document with content controls to HTML while maintaining control attributes as data-attributes.
  - Workflow: native-sdt-api
  - Outputs: html
  - Selected engine: llm
- `merge-multiple-word-documents-preserving-existing-content-controls-and-updating-their-ids.cs`
  - Task: Merge multiple Word documents, preserving existing content controls and updating their IDs.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `extract-all-repeating-section-content-controls-from-a-word-file-and-serialize-each-instanc.cs`
  - Task: Extract all repeating section content controls from a Word file and serialize each instance to JSON.
  - Workflow: xml-json-export
  - Outputs: json
  - Selected engine: mcp
- `implement-error-handling-for-missing-xml-nodes-when-binding-data-to-a-content-control.cs`
  - Task: Implement error handling for missing XML nodes when binding data to a content control.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: existing_repo
- `use-a-content-control-to-store-custom-metadata-and-extract-it-for-indexing-in-a-search-eng.cs`
  - Task: Use a content control to store custom metadata and extract it for indexing in a search engine.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Use of non-existent API members**
  - Symptom: Compilation fails with CS0117 or CS1061 because the code references members such as SdtType.DateTime, StructuredDocumentTag.DateValue, or PdfSaveOptions.ShowPlaceholderText.
  - Root cause: The example assumes newer or custom extensions that are not present in the documented Aspose.Words API.
  - Preferred fix: Replace the missing members with supported alternatives, for example use a plain-text SDT to hold a formatted date string and set PdfSaveOptions.ShowPlaceholderText via the appropriate property if available.
  - Prevention: Consult the official Aspose.Words API reference for each version and limit code to members that exist in the declared package version.

- **Invalid method arguments**
  - Symptom: Build errors CS1501, CS1503, or CS1729 occur when calling methods like InsertOleObject or CustomXmlPart constructors with the wrong number or type of parameters.
  - Root cause: The code supplies arguments of incorrect types (e.g., string instead of CustomXmlPart) or omits required parameters.
  - Preferred fix: Pass arguments that match the method signature; for InsertOleObject provide a Stream for the OLE data, a display name string, a boolean for icon display, and optionally a Stream for the icon image.
  - Prevention: Reference method overload documentation and use IntelliSense to verify parameter types before coding.

- **Invalid node insertion location**
  - Symptom: Runtime throws ArgumentException with the message "Cannot insert a node of this type at this location" when a control is appended to an unsupported parent.
  - Root cause: An SDT is inserted directly into a CompositeNode that does not accept that node type, such as inserting a block-level SDT into a Run.
  - Preferred fix: Insert block-level SDTs into Section.Body or Paragraph, and inline SDTs into Paragraph or Run containers. Use DocumentBuilder.MoveTo to position the cursor correctly before insertion.
  - Prevention: Validate the target node's NodeType against the SDT's MarkupLevel before calling AppendChild or InsertAfter.

- **Missing building block for placeholder name**
  - Symptom: InvalidOperationException is thrown when setting StructuredDocumentTag.PlaceholderName without an existing building block of that name.
  - Root cause: PlaceholderName requires a BuildingBlock entry in the document's glossary, which the code does not create.
  - Preferred fix: Create a BuildingBlock with the desired name, add it to the document's GlossaryDocument, and then assign PlaceholderName to that block.
  - Prevention: Always add the required building block before referencing it as a placeholder.

- **Incorrect theme or color API usage**
  - Symptom: Compilation fails with CS0246 or CS1061 because types like Theme, ThemeColor, or ThemeColorScheme are not defined, and color conversion errors arise between Aspose.Drawing.Color and System.Drawing.Color.
  - Root cause: The code attempts to use theme APIs that belong to a different library version or mixes drawing namespaces.
  - Preferred fix: Use Aspose.Words.Themes.Theme for theme access and Aspose.Drawing.Color for color values; avoid System.Drawing.Color unless explicitly required.
  - Prevention: Reference the correct namespace (Aspose.Words.Themes) and keep drawing operations within Aspose.Drawing.

- **Missing type imports for table structures**
  - Symptom: CS0246 errors indicate that Table, Row, or Cell types cannot be found during compilation.
  - Root cause: The required Aspose.Words.Tables namespace is not imported or the project lacks a reference to the assembly containing those types.
  - Preferred fix: Add "using Aspose.Words.Tables;" at the top of the file and ensure the Aspose.Words package version includes the Tables namespace.
  - Prevention: Verify that all table-related types are covered by the declared required_namespaces list.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
dotnet add package Aspose.Drawing.Common --version 26.10.0
dotnet add package Newtonsoft.Json
```

```powershell
Copy-Item ..\content-control\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Periodically review the Aspose.Words release notes to update the strict_api_bans list when new members are added or deprecated.
