---
name: content-control
description: Verified C# examples for content-control scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Content Control

## Purpose

Provide an AI coding assistant with clear guidance and reusable patterns for creating, configuring, and validating content controls in Word documents using only the official Aspose.Words API surface.

## Non-negotiable conventions

- Do not reference helper methods, properties, or overloads that are not part of the Aspose.Words SDK.
- Do not treat a generic Node as a StructuredDocumentTag unless it was created as such.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Input Bootstrap workflow**: 3 examples
- **Native Sdt API workflow**: 22 examples
- **Xml Json Export workflow**: 10 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Verify that every StructuredDocumentTag uses a supported SdtType enum value.
2. Confirm that insertion points are compatible with the control's MarkupLevel (inline vs. block).
3. Ensure that any XML mapping references a valid CustomXmlPart and XPath expression.

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
  - Selected engine: llm
- `batch-process-a-folder-of-word-files-inserting-a-header-content-control-with-document-meta.cs`
  - Task: Batch process a folder of Word files, inserting a header content control with document metadata.
  - Workflow: input-bootstrap
  - Outputs: docx
  - Selected engine: mcp
- `load-a-doc-file-add-a-date-picker-content-control-and-save-the-result-as-docx.cs`
  - Task: Load a DOC file, add a date picker content control, and save the result as DOCX.
  - Workflow: native-sdt-api
  - Outputs: docx, doc
  - Selected engine: llm
- `use-a-content-control-to-embed-an-ole-object-and-ensure-it-renders-correctly-after-convers.cs`
  - Task: Use a content control to embed an OLE object and ensure it renders correctly after conversion to PDF.
  - Workflow: native-sdt-api
  - Outputs: pdf
  - Selected engine: existing_repo
- `use-a-content-control-to-embed-a-hyperlink-and-verify-its-target-url-after-document-conver.cs`
  - Task: Use a content control to embed a hyperlink and verify its target URL after document conversion.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: mcp
- `create-a-repeating-section-content-control-that-repeats-a-table-row-for-each-item-in-a-col.cs`
  - Task: Create a repeating section content control that repeats a table row for each item in a collection.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: llm
- `create-a-content-control-that-repeats-a-paragraph-for-each-entry-in-a-json-array-during-im.cs`
  - Task: Create a content control that repeats a paragraph for each entry in a JSON array during import.
  - Workflow: xml-json-export
  - Outputs: json
  - Selected engine: mcp
- `bind-a-dropdown-list-content-control-to-an-xml-data-source-and-populate-options-dynamicall.cs`
  - Task: Bind a dropdown list content control to an XML data source and populate options dynamically.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: mcp
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
  - Selected engine: existing_repo
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
  - Selected engine: existing_repo
- `serialize-the-xml-mapping-of-all-content-controls-to-an-external-xsd-schema-file.cs`
  - Task: Serialize the XML mapping of all content controls to an external XSD schema file.
  - Workflow: xml-json-export
  - Outputs: xml
  - Selected engine: mcp
- `export-the-contents-of-all-checkbox-content-controls-to-a-csv-file-for-data-analysis.cs`
  - Task: Export the contents of all checkbox content controls to a CSV file for data analysis.
  - Workflow: xml-json-export
  - Outputs: csv
  - Selected engine: mcp
- `convert-a-docx-document-containing-content-controls-to-pdf-while-preserving-control-placeh.cs`
  - Task: Convert a DOCX document containing content controls to PDF while preserving control placeholders.
  - Workflow: input-bootstrap
  - Outputs: pdf, docx
  - Selected engine: mcp
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
  - Selected engine: mcp
- `merge-multiple-word-documents-preserving-existing-content-controls-and-updating-their-ids.cs`
  - Task: Merge multiple Word documents, preserving existing content controls and updating their IDs.
  - Workflow: native-sdt-api
  - Outputs: docx
  - Selected engine: llm
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
  - Selected engine: llm

## Common failure patterns and preferred agent fixes

- **Invented SDT helper methods**
  - Symptom: Compilation fails because methods such as StartStructuredDocumentTag or InsertStructuredDocumentTag do not exist.
  - Root cause: The code assumes convenience APIs that are not provided by Aspose.Words.
  - Preferred fix: Instantiate StructuredDocumentTag directly and insert it with DocumentBuilder.InsertNode or Node.AppendChild.
  - Prevention: Consult the official API reference for StructuredDocumentTag constructors and insertion methods before writing helper wrappers.

- **Invalid insertion location**
  - Symptom: Runtime error "Cannot insert a node of this type at this location" is thrown when adding a control.
  - Root cause: The control's MarkupLevel does not match the target container (e.g., inserting an inline SDT into a HeaderFooter collection).
  - Preferred fix: Select a parent node that matches the control's MarkupLevel: use Paragraph or Run for Inline, and Header/Footer or Body for Block.
  - Prevention: Validate the parent node type with Node.NodeType before calling InsertNode or AppendChild.

- **Ambiguous type references**
  - Symptom: Compiler reports ambiguous references between Aspose.Drawing.Font and Aspose.Words.Font, or between Color types.
  - Root cause: Both Aspose.Drawing and Aspose.Words define types with identical names and the code does not qualify them.
  - Preferred fix: Fully qualify the type (e.g., Aspose.Drawing.Font) or add an alias using directive.
  - Prevention: Prefer explicit namespace prefixes for types that exist in multiple Aspose assemblies.

- **Unsupported SdtType members**
  - Symptom: CS0117 errors indicate that SdtType.DateTime or similar members are missing.
  - Root cause: The version of Aspose.Words used does not expose dedicated date-time SDT types.
  - Preferred fix: Use a plain-text StructuredDocumentTag and populate it with a formatted date string.
  - Prevention: Check the SDK release notes for available SdtType values and avoid hard-coding members that may be removed.

- **Incorrect InsertOleObject overload**
  - Symptom: CS1501 errors show that no InsertOleObject overload matches the supplied argument list.
  - Root cause: The code passes a file path string where a Stream is required, and includes parameters that do not exist.
  - Preferred fix: Open the OLE file as a FileStream and call InsertOleObject(stream, displayName, isIcon, iconStream) with the correct signatures.
  - Prevention: Reference the method signature in the Aspose.Words documentation and pass only the supported parameters.

- **Missing SaveFormat constants**
  - Symptom: CS0117 errors occur because SaveFormat.Xml or SaveFormat.OpenXml are not defined.
  - Root cause: The SDK provides SaveFormat.Docx, Html, Pdf, etc., but not an XML constant.
  - Preferred fix: Save the document using SaveFormat.Docx or Html and, if raw XML is needed, use doc.Save(stream, SaveOptions.CreateSaveOptions(SaveFormat.OpenXml)).
  - Prevention: Use only the SaveFormat members listed in the official API reference.

- **Incorrect CustomXmlPart construction**
  - Symptom: CS1503 errors indicate that a string is being passed where a CustomXmlPart instance is required.
  - Root cause: The code attempts to assign raw XML text directly to a property expecting a CustomXmlPart object.
  - Preferred fix: Create the part with doc.CustomXmlParts.Add(id, xmlString) and then use the returned CustomXmlPart for mapping.
  - Prevention: Always obtain a CustomXmlPart via the Add method before using it in XmlMapping.SetMapping.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.10.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
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

- Review the Aspose.Words release notes quarterly and update example code whenever new SdtType values or insertion APIs are added or deprecated.
