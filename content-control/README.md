# Content Control Examples for Aspose.Words for .NET

The Content Control category demonstrates how to manipulate structured document tags (SDTs) in Aspose.Words using supported APIs. Verified examples include inserting plain-text, rich-text, picture, and repeating-section controls, binding them to custom XML, and converting documents while preserving control metadata.

## Snapshot

- Category: Content Control
- Slug: content-control
- Total examples: 35
- Publish-ready successful examples: 35 / 35
- Source run: 20261124_190740_1c00f5
- Input Bootstrap examples: 3
- Native Sdt API examples: 22
- Xml Json Export examples: 10

## Category rules that shaped these examples

- Do not reference Aspose.Words members that do not exist in the current version, such as SdtType.DateTime or StructuredDocumentTag.DateValue.
- Avoid invented helper methods like StartStructuredDocumentTag or InsertStructuredDocumentTag overloads that are not part of the API.
- Instantiate StructuredDocumentTag with a valid SdtType and MarkupLevel, then append child nodes (Run, Paragraph, Table, etc.) as needed.
- Use DocumentBuilder to move the cursor into a paragraph inside the SDT before inserting images, OLE objects, or hyperlinks.
- Initialize all non-nullable reference-type properties (e.g., Title, Tag) when creating a control.
- Declare nullable locals for values that may be absent and guard them with null checks before dereferencing.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.11.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
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
dotnet add package Aspose.Words --version 26.11.0

# PowerShell example
Copy-Item ..\content-control\insert-a-plain-text-content-control-at-a-specific-bookmark-in-a-docx-document.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `insert-a-plain-text-content-control-at-a-specific-bookmark-in-a-docx-document.cs` | Insert a plain text content control at a specific bookmark in a DOCX document. | input-bootstrap | docx | mcp |
| 2 | `add-a-picture-content-control-that-references-an-external-image-file-and-embed-it-on-save.cs` | Add a picture content control that references an external image file and embed it on save. | native-sdt-api | docx | mcp |
| 3 | `batch-process-a-folder-of-word-files-inserting-a-header-content-control-with-document-meta.cs` | Batch process a folder of Word files, inserting a header content control with document metadata. | input-bootstrap | docx | mcp |
| 4 | `load-a-doc-file-add-a-date-picker-content-control-and-save-the-result-as-docx.cs` | Load a DOC file, add a date picker content control, and save the result as DOCX. | native-sdt-api | docx, doc | mcp |
| 5 | `use-a-content-control-to-embed-an-ole-object-and-ensure-it-renders-correctly-after-convers.cs` | Use a content control to embed an OLE object and ensure it renders correctly after conversion to PDF. | native-sdt-api | pdf | existing_repo |
| 6 | `use-a-content-control-to-embed-a-hyperlink-and-verify-its-target-url-after-document-conver.cs` | Use a content control to embed a hyperlink and verify its target URL after document conversion. | native-sdt-api | docx | llm |
| 7 | `create-a-repeating-section-content-control-that-repeats-a-table-row-for-each-item-in-a-col.cs` | Create a repeating section content control that repeats a table row for each item in a collection. | native-sdt-api | docx | mcp |
| 8 | `create-a-content-control-that-repeats-a-paragraph-for-each-entry-in-a-json-array-during-im.cs` | Create a content control that repeats a paragraph for each entry in a JSON array during import. | xml-json-export | json | mcp |
| 9 | `bind-a-dropdown-list-content-control-to-an-xml-data-source-and-populate-options-dynamicall.cs` | Bind a dropdown list content control to an XML data source and populate options dynamically. | xml-json-export | xml | existing_repo |
| 10 | `apply-custom-xml-mapping-to-a-plain-text-content-control-to-synchronize-with-external-data.cs` | Apply custom XML mapping to a plain text content control to synchronize with external data fields. | xml-json-export | xml | existing_repo |
| 11 | `apply-a-custom-style-to-the-text-inside-a-rich-text-content-control-programmatically.cs` | Apply a custom style to the text inside a rich text content control programmatically. | native-sdt-api | docx | mcp |
| 12 | `programmatically-set-the-title-and-tag-properties-of-a-content-control-for-later-identific.cs` | Programmatically set the title and tag properties of a content control for later identification. | native-sdt-api | docx | mcp |
| 13 | `update-the-tag-of-all-content-controls-in-a-document-to-follow-a-standardized-naming-conve.cs` | Update the tag of all content controls in a document to follow a standardized naming convention. | native-sdt-api | docx | mcp |
| 14 | `configure-a-content-control-to-allow-only-numeric-input-and-enforce-validation-during-edit.cs` | Configure a content control to allow only numeric input and enforce validation during editing. | native-sdt-api | docx | llm |
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
| 25 | `retrieve-the-inner-xml-of-a-content-control-and-transform-it-using-an-xslt-stylesheet.cs` | Retrieve the inner XML of a content control and transform it using an XSLT stylesheet. | xml-json-export | xml | mcp |
| 26 | `serialize-the-xml-mapping-of-all-content-controls-to-an-external-xsd-schema-file.cs` | Serialize the XML mapping of all content controls to an external XSD schema file. | xml-json-export | xml | existing_repo |
| 27 | `export-the-contents-of-all-checkbox-content-controls-to-a-csv-file-for-data-analysis.cs` | Export the contents of all checkbox content controls to a CSV file for data analysis. | xml-json-export | csv | mcp |
| 28 | `convert-a-docx-document-containing-content-controls-to-pdf-while-preserving-control-placeh.cs` | Convert a DOCX document containing content controls to PDF while preserving control placeholders. | input-bootstrap | pdf, docx | llm |
| 29 | `generate-a-pdf-a-compliant-document-from-a-word-file-while-keeping-content-control-tags-in.cs` | Generate a PDF/A compliant document from a Word file while keeping content control tags intact. | native-sdt-api | pdf | mcp |
| 30 | `export-a-document-containing-content-controls-to-xps-format-while-preserving-control-bound.cs` | Export a document containing content controls to XPS format while preserving control boundaries. | xml-json-export | xps | mcp |
| 31 | `convert-a-document-with-content-controls-to-html-while-maintaining-control-attributes-as-d.cs` | Convert a document with content controls to HTML while maintaining control attributes as data-attributes. | native-sdt-api | html | llm |
| 32 | `merge-multiple-word-documents-preserving-existing-content-controls-and-updating-their-ids.cs` | Merge multiple Word documents, preserving existing content controls and updating their IDs. | native-sdt-api | docx | mcp |
| 33 | `extract-all-repeating-section-content-controls-from-a-word-file-and-serialize-each-instanc.cs` | Extract all repeating section content controls from a Word file and serialize each instance to JSON. | xml-json-export | json | mcp |
| 34 | `implement-error-handling-for-missing-xml-nodes-when-binding-data-to-a-content-control.cs` | Implement error handling for missing XML nodes when binding data to a content control. | xml-json-export | xml | existing_repo |
| 35 | `use-a-content-control-to-store-custom-metadata-and-extract-it-for-indexing-in-a-search-eng.cs` | Use a content control to store custom metadata and extract it for indexing in a search engine. | native-sdt-api | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Use of non-existent API members

- Symptom: Compilation fails with CS0117 or CS1061 because the code references members such as SdtType.DateTime, StructuredDocumentTag.DateValue, or PdfSaveOptions.ShowPlaceholderText.
- Root cause: The example assumes newer or custom extensions that are not present in the documented Aspose.Words API.
- Preferred fix: Replace the missing members with supported alternatives, for example use a plain-text SDT to hold a formatted date string and set PdfSaveOptions.ShowPlaceholderText via the appropriate property if available.
- Prevention: Consult the official Aspose.Words API reference for each version and limit code to members that exist in the declared package version.

### Invalid method arguments

- Symptom: Build errors CS1501, CS1503, or CS1729 occur when calling methods like InsertOleObject or CustomXmlPart constructors with the wrong number or type of parameters.
- Root cause: The code supplies arguments of incorrect types (e.g., string instead of CustomXmlPart) or omits required parameters.
- Preferred fix: Pass arguments that match the method signature; for InsertOleObject provide a Stream for the OLE data, a display name string, a boolean for icon display, and optionally a Stream for the icon image.
- Prevention: Reference method overload documentation and use IntelliSense to verify parameter types before coding.

### Invalid node insertion location

- Symptom: Runtime throws ArgumentException with the message "Cannot insert a node of this type at this location" when a control is appended to an unsupported parent.
- Root cause: An SDT is inserted directly into a CompositeNode that does not accept that node type, such as inserting a block-level SDT into a Run.
- Preferred fix: Insert block-level SDTs into Section.Body or Paragraph, and inline SDTs into Paragraph or Run containers. Use DocumentBuilder.MoveTo to position the cursor correctly before insertion.
- Prevention: Validate the target node's NodeType against the SDT's MarkupLevel before calling AppendChild or InsertAfter.

### Missing building block for placeholder name

- Symptom: InvalidOperationException is thrown when setting StructuredDocumentTag.PlaceholderName without an existing building block of that name.
- Root cause: PlaceholderName requires a BuildingBlock entry in the document's glossary, which the code does not create.
- Preferred fix: Create a BuildingBlock with the desired name, add it to the document's GlossaryDocument, and then assign PlaceholderName to that block.
- Prevention: Always add the required building block before referencing it as a placeholder.

### Incorrect theme or color API usage

- Symptom: Compilation fails with CS0246 or CS1061 because types like Theme, ThemeColor, or ThemeColorScheme are not defined, and color conversion errors arise between Aspose.Drawing.Color and System.Drawing.Color.
- Root cause: The code attempts to use theme APIs that belong to a different library version or mixes drawing namespaces.
- Preferred fix: Use Aspose.Words.Themes.Theme for theme access and Aspose.Drawing.Color for color values; avoid System.Drawing.Color unless explicitly required.
- Prevention: Reference the correct namespace (Aspose.Words.Themes) and keep drawing operations within Aspose.Drawing.

### Missing type imports for table structures

- Symptom: CS0246 errors indicate that Table, Row, or Cell types cannot be found during compilation.
- Root cause: The required Aspose.Words.Tables namespace is not imported or the project lacks a reference to the assembly containing those types.
- Preferred fix: Add "using Aspose.Words.Tables;" at the top of the file and ensure the Aspose.Words package version includes the Tables namespace.
- Prevention: Verify that all table-related types are covered by the declared required_namespaces list.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Periodically review the Aspose.Words release notes to update the strict_api_bans list when new members are added or deprecated.
