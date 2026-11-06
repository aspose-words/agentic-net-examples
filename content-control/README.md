# Content Control Examples for Aspose.Words for .NET

The Content Control category demonstrates how to manipulate Aspose.Words StructuredDocumentTag objects for a variety of scenarios, including plain-text, picture, date, OLE, hyperlink, repeating-section, dropdown, and custom XML-mapped controls. The verified examples show end-to-end implementations that create sample inputs, configure supported properties, and persist the results without relying on nonexistent APIs.

## Snapshot

- Category: Content Control
- Slug: content-control
- Total examples: 35
- Publish-ready successful examples: 35 / 35
- Source run: 20261102_031946_08fb5b
- Input Bootstrap examples: 3
- Native Sdt API examples: 22
- Xml Json Export examples: 10

## Category rules that shaped these examples

- Do not reference helper methods, properties, or overloads that are not part of the Aspose.Words SDK.
- Do not treat a generic Node as a StructuredDocumentTag unless it was created as such.
- Create StructuredDocumentTag instances directly with the constructor (doc, SdtType, MarkupLevel) and add child runs, paragraphs, tables, or shapes as needed.
- Use Aspose.Drawing for image generation and Newtonsoft.Json for JSON handling; avoid System.Drawing and undocumented JSON utilities.
- Initialize all non-nullable reference-type members (e.g., Title, Tag) when constructing a control.
- Declare variables as nullable when a lookup may return null and guard them with explicit null checks before dereferencing.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.10.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
dotnet add package Aspose.Drawing.Common --version 26.10.0
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
dotnet add package Aspose.Words --version 26.10.0

# PowerShell example
Copy-Item ..\content-control\insert-a-plain-text-content-control-at-a-specific-bookmark-in-a-docx-document.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `insert-a-plain-text-content-control-at-a-specific-bookmark-in-a-docx-document.cs` | Insert a plain text content control at a specific bookmark in a DOCX document. | input-bootstrap | docx | mcp |
| 2 | `add-a-picture-content-control-that-references-an-external-image-file-and-embed-it-on-save.cs` | Add a picture content control that references an external image file and embed it on save. | native-sdt-api | docx | llm |
| 3 | `batch-process-a-folder-of-word-files-inserting-a-header-content-control-with-document-meta.cs` | Batch process a folder of Word files, inserting a header content control with document metadata. | input-bootstrap | docx | mcp |
| 4 | `load-a-doc-file-add-a-date-picker-content-control-and-save-the-result-as-docx.cs` | Load a DOC file, add a date picker content control, and save the result as DOCX. | native-sdt-api | docx, doc | llm |
| 5 | `use-a-content-control-to-embed-an-ole-object-and-ensure-it-renders-correctly-after-convers.cs` | Use a content control to embed an OLE object and ensure it renders correctly after conversion to PDF. | native-sdt-api | pdf | existing_repo |
| 6 | `use-a-content-control-to-embed-a-hyperlink-and-verify-its-target-url-after-document-conver.cs` | Use a content control to embed a hyperlink and verify its target URL after document conversion. | native-sdt-api | docx | mcp |
| 7 | `create-a-repeating-section-content-control-that-repeats-a-table-row-for-each-item-in-a-col.cs` | Create a repeating section content control that repeats a table row for each item in a collection. | native-sdt-api | docx | llm |
| 8 | `create-a-content-control-that-repeats-a-paragraph-for-each-entry-in-a-json-array-during-im.cs` | Create a content control that repeats a paragraph for each entry in a JSON array during import. | xml-json-export | json | mcp |
| 9 | `bind-a-dropdown-list-content-control-to-an-xml-data-source-and-populate-options-dynamicall.cs` | Bind a dropdown list content control to an XML data source and populate options dynamically. | xml-json-export | xml | mcp |
| 10 | `apply-custom-xml-mapping-to-a-plain-text-content-control-to-synchronize-with-external-data.cs` | Apply custom XML mapping to a plain text content control to synchronize with external data fields. | xml-json-export | xml | existing_repo |
| 11 | `apply-a-custom-style-to-the-text-inside-a-rich-text-content-control-programmatically.cs` | Apply a custom style to the text inside a rich text content control programmatically. | native-sdt-api | docx | mcp |
| 12 | `programmatically-set-the-title-and-tag-properties-of-a-content-control-for-later-identific.cs` | Programmatically set the title and tag properties of a content control for later identification. | native-sdt-api | docx | mcp |
| 13 | `update-the-tag-of-all-content-controls-in-a-document-to-follow-a-standardized-naming-conve.cs` | Update the tag of all content controls in a document to follow a standardized naming convention. | native-sdt-api | docx | mcp |
| 14 | `configure-a-content-control-to-allow-only-numeric-input-and-enforce-validation-during-edit.cs` | Configure a content control to allow only numeric input and enforce validation during editing. | native-sdt-api | docx | existing_repo |
| 15 | `validate-that-required-content-controls-contain-non-empty-text-before-saving-the-document.cs` | Validate that required content controls contain non-empty text before saving the document. | native-sdt-api | docx | mcp |
| 16 | `lock-a-content-control-to-prevent-user-editing-and-enforce-read-only-behavior-in-the-final.cs` | Lock a content control to prevent user editing and enforce read-only behavior in the final document. | native-sdt-api | docx | mcp |
| 17 | `set-the-placeholder-text-color-inside-a-content-control-to-match-the-document-theme.cs` | Set the placeholder text color inside a content control to match the document theme. | native-sdt-api | docx | existing_repo |
| 18 | `replace-placeholder-text-in-a-content-control-with-values-from-a-dictionary-of-user-inputs.cs` | Replace placeholder text in a content control with values from a dictionary of user inputs. | native-sdt-api | docx | mcp |
| 19 | `replace-the-contents-of-a-rich-text-content-control-with-formatted-html-retrieved-from-a-w.cs` | Replace the contents of a rich text content control with formatted HTML retrieved from a web service. | native-sdt-api | html | mcp |
| 20 | `programmatically-clear-the-contents-of-a-content-control-without-deleting-the-control-itse.cs` | Programmatically clear the contents of a content control without deleting the control itself. | native-sdt-api | docx | mcp |
| 21 | `programmatically-duplicate-a-content-control-and-insert-the-copy-at-a-different-location-i.cs` | Programmatically duplicate a content control and insert the copy at a different location in the document. | native-sdt-api | docx | mcp |
| 22 | `remove-all-picture-content-controls-from-a-document-and-replace-them-with-inline-images.cs` | Remove all picture content controls from a document and replace them with inline images. | native-sdt-api | docx | mcp |
| 23 | `detect-and-list-any-nested-content-controls-within-a-repeating-section-for-structural-insp.cs` | Detect and list any nested content controls within a repeating section for structural inspection. | native-sdt-api | docx | mcp |
| 24 | `iterate-through-all-content-controls-in-a-document-and-generate-a-summary-report-of-their.cs` | Iterate through all content controls in a document and generate a summary report of their types. | xml-json-export | docx | mcp |
| 25 | `retrieve-the-inner-xml-of-a-content-control-and-transform-it-using-an-xslt-stylesheet.cs` | Retrieve the inner XML of a content control and transform it using an XSLT stylesheet. | xml-json-export | xml | existing_repo |
| 26 | `serialize-the-xml-mapping-of-all-content-controls-to-an-external-xsd-schema-file.cs` | Serialize the XML mapping of all content controls to an external XSD schema file. | xml-json-export | xml | mcp |
| 27 | `export-the-contents-of-all-checkbox-content-controls-to-a-csv-file-for-data-analysis.cs` | Export the contents of all checkbox content controls to a CSV file for data analysis. | xml-json-export | csv | mcp |
| 28 | `convert-a-docx-document-containing-content-controls-to-pdf-while-preserving-control-placeh.cs` | Convert a DOCX document containing content controls to PDF while preserving control placeholders. | input-bootstrap | pdf, docx | mcp |
| 29 | `generate-a-pdf-a-compliant-document-from-a-word-file-while-keeping-content-control-tags-in.cs` | Generate a PDF/A compliant document from a Word file while keeping content control tags intact. | native-sdt-api | pdf | mcp |
| 30 | `export-a-document-containing-content-controls-to-xps-format-while-preserving-control-bound.cs` | Export a document containing content controls to XPS format while preserving control boundaries. | xml-json-export | xps | mcp |
| 31 | `convert-a-document-with-content-controls-to-html-while-maintaining-control-attributes-as-d.cs` | Convert a document with content controls to HTML while maintaining control attributes as data-attributes. | native-sdt-api | html | mcp |
| 32 | `merge-multiple-word-documents-preserving-existing-content-controls-and-updating-their-ids.cs` | Merge multiple Word documents, preserving existing content controls and updating their IDs. | native-sdt-api | docx | llm |
| 33 | `extract-all-repeating-section-content-controls-from-a-word-file-and-serialize-each-instanc.cs` | Extract all repeating section content controls from a Word file and serialize each instance to JSON. | xml-json-export | json | mcp |
| 34 | `implement-error-handling-for-missing-xml-nodes-when-binding-data-to-a-content-control.cs` | Implement error handling for missing XML nodes when binding data to a content control. | xml-json-export | xml | existing_repo |
| 35 | `use-a-content-control-to-store-custom-metadata-and-extract-it-for-indexing-in-a-search-eng.cs` | Use a content control to store custom metadata and extract it for indexing in a search engine. | native-sdt-api | docx | llm |

## Common failure patterns seen during generation and how they were corrected

### Invented SDT helper methods

- Symptom: Compilation fails because methods such as StartStructuredDocumentTag or InsertStructuredDocumentTag do not exist.
- Root cause: The code assumes convenience APIs that are not provided by Aspose.Words.
- Preferred fix: Instantiate StructuredDocumentTag directly and insert it with DocumentBuilder.InsertNode or Node.AppendChild.
- Prevention: Consult the official API reference for StructuredDocumentTag constructors and insertion methods before writing helper wrappers.

### Invalid insertion location

- Symptom: Runtime error "Cannot insert a node of this type at this location" is thrown when adding a control.
- Root cause: The control's MarkupLevel does not match the target container (e.g., inserting an inline SDT into a HeaderFooter collection).
- Preferred fix: Select a parent node that matches the control's MarkupLevel: use Paragraph or Run for Inline, and Header/Footer or Body for Block.
- Prevention: Validate the parent node type with Node.NodeType before calling InsertNode or AppendChild.

### Ambiguous type references

- Symptom: Compiler reports ambiguous references between Aspose.Drawing.Font and Aspose.Words.Font, or between Color types.
- Root cause: Both Aspose.Drawing and Aspose.Words define types with identical names and the code does not qualify them.
- Preferred fix: Fully qualify the type (e.g., Aspose.Drawing.Font) or add an alias using directive.
- Prevention: Prefer explicit namespace prefixes for types that exist in multiple Aspose assemblies.

### Unsupported SdtType members

- Symptom: CS0117 errors indicate that SdtType.DateTime or similar members are missing.
- Root cause: The version of Aspose.Words used does not expose dedicated date-time SDT types.
- Preferred fix: Use a plain-text StructuredDocumentTag and populate it with a formatted date string.
- Prevention: Check the SDK release notes for available SdtType values and avoid hard-coding members that may be removed.

### Incorrect InsertOleObject overload

- Symptom: CS1501 errors show that no InsertOleObject overload matches the supplied argument list.
- Root cause: The code passes a file path string where a Stream is required, and includes parameters that do not exist.
- Preferred fix: Open the OLE file as a FileStream and call InsertOleObject(stream, displayName, isIcon, iconStream) with the correct signatures.
- Prevention: Reference the method signature in the Aspose.Words documentation and pass only the supported parameters.

### Missing SaveFormat constants

- Symptom: CS0117 errors occur because SaveFormat.Xml or SaveFormat.OpenXml are not defined.
- Root cause: The SDK provides SaveFormat.Docx, Html, Pdf, etc., but not an XML constant.
- Preferred fix: Save the document using SaveFormat.Docx or Html and, if raw XML is needed, use doc.Save(stream, SaveOptions.CreateSaveOptions(SaveFormat.OpenXml)).
- Prevention: Use only the SaveFormat members listed in the official API reference.

### Incorrect CustomXmlPart construction

- Symptom: CS1503 errors indicate that a string is being passed where a CustomXmlPart instance is required.
- Root cause: The code attempts to assign raw XML text directly to a property expecting a CustomXmlPart object.
- Preferred fix: Create the part with doc.CustomXmlParts.Add(id, xmlString) and then use the returned CustomXmlPart for mapping.
- Prevention: Always obtain a CustomXmlPart via the Add method before using it in XmlMapping.SetMapping.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Review the Aspose.Words release notes quarterly and update example code whenever new SdtType values or insertion APIs are added or deprecated.
