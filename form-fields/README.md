# Form Fields Examples for Aspose.Words for .NET

The Form Fields category covers the creation, manipulation, and protection of legacy form fields in Word documents using Aspose.Words. Verified examples demonstrate inserting text inputs, check boxes, combo boxes, updating values, and applying form-only protection with robust validation.

## Snapshot

- Category: Form Fields
- Slug: form-fields
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261102_031946_08fb5b
- Text Form Field examples: 7
- General Form Field examples: 17
- Dropdown Form Field examples: 4
- Inspection And Reporting examples: 2

## Category rules that shaped these examples

- Never assume that a form field already exists in the target document; always insert or verify it explicitly.
- Do not use StructuredDocumentTag or other content-control APIs when the task requires legacy form fields.
- Avoid inventing helper methods that are not part of the Aspose.Words API.
- Use DocumentBuilder.InsertTextInput, InsertCheckBox, and InsertComboBox to create legacy fields.
- Access existing fields through Document.Range.FormFields and use Result, Checked, or DropDownItems according to the field type.
- Validate that each field was added and that its properties match the expected defaults before saving.
- Check for null after retrieving a FormField from the collection before dereferencing it.
- Guard against missing items in a combo box by ensuring the items array contains at least one element.
- Validate numeric arguments such as maxLength, selectedIndex, and size to stay within the method's contract.

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

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\form-fields\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `form-fields/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0

# PowerShell example
Copy-Item ..\form-fields\insert-a-text-input-form-field-with-placeholder-text-using-documentbuilder-in-a-new-docume.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `insert-a-text-input-form-field-with-placeholder-text-using-documentbuilder-in-a-new-docume.cs` | Insert a text input form field with placeholder text using DocumentBuilder in a new document. | text-form-field | docx | mcp |
| 2 | `add-a-check-box-form-field-set-default-checked-state-and-custom-size-via-documentbuilder.cs` | Add a check box form field, set default checked state and custom size via DocumentBuilder. | general-form-field | docx | mcp |
| 3 | `create-a-combo-box-form-field-containing-three-items-and-set-default-selected-index-using.cs` | Create a combo box form field containing three items and set default selected index using DocumentBuilder. | dropdown-form-field | docx | mcp |
| 4 | `specify-a-name-when-inserting-a-text-input-field-to-automatically-generate-a-matching-book.cs` | Specify a name when inserting a text input field to automatically generate a matching bookmark. | text-form-field | docx | mcp |
| 5 | `insert-a-text-input-field-with-maximum-length-of-50-characters-and-numeric-format-using-do.cs` | Insert a text input field with maximum length of 50 characters and numeric format using DocumentBuilder. | text-form-field | docx | llm |
| 6 | `insert-a-text-input-field-with-custom-date-format-and-default-current-date-value-using-doc.cs` | Insert a text input field with custom date format and default current date value using DocumentBuilder. | text-form-field | docx | mcp |
| 7 | `implement-a-reusable-method-that-adds-a-combo-box-form-field-with-customizable-items-and-d.cs` | Implement a reusable method that adds a combo box form field with customizable items and default index. | dropdown-form-field | docx | mcp |
| 8 | `assign-a-unique-name-to-each-inserted-form-field-to-ensure-distinct-automatic-bookmarks.cs` | Assign a unique name to each inserted form field to ensure distinct automatic bookmarks. | general-form-field | docx | llm |
| 9 | `batch-insert-multiple-text-input-form-fields-into-a-template-using-a-loop-over-field-defin.cs` | Batch insert multiple text input form fields into a template using a loop over field definitions. | text-form-field | docx | mcp |
| 10 | `load-a-docx-file-change-combo-box-selections-programmatically-and-save-the-modified-docume.cs` | Load a DOCX file, change combo box selections programmatically, and save the modified document. | dropdown-form-field | docx | mcp |
| 11 | `update-result-values-of-several-check-box-form-fields-based-on-external-json-configuration.cs` | Update result values of several check box form fields based on external JSON configuration data. | inspection-and-reporting | json | mcp |
| 12 | `set-the-result-property-of-a-text-input-form-field-to-a-predefined-string-value.cs` | Set the Result property of a text input form field to a predefined string value. | text-form-field | docx | mcp |
| 13 | `read-the-result-property-of-a-check-box-form-field-to-determine-whether-it-is-checked.cs` | Read the Result property of a check box form field to determine whether it is checked. | general-form-field | docx | mcp |
| 14 | `toggle-the-checked-state-of-a-specific-check-box-form-field-based-on-external-configuratio.cs` | Toggle the checked state of a specific check box form field based on external configuration settings. | general-form-field | docx | mcp |
| 15 | `change-the-size-of-an-existing-check-box-form-field-programmatically-to-improve-visual-con.cs` | Change the size of an existing check box form field programmatically to improve visual consistency. | general-form-field | docx | mcp |
| 16 | `retrieve-a-form-field-by-its-name-from-a-loaded-document-and-read-its-result-property.cs` | Retrieve a form field by its name from a loaded document and read its Result property. | general-form-field | docx | llm |
| 17 | `access-a-form-field-by-index-from-the-formfields-collection-and-modify-its-result-value.cs` | Access a form field by index from the FormFields collection and modify its Result value. | general-form-field | docx | mcp |
| 18 | `use-formfield-type-enumeration-to-differentiate-between-text-input-check-box-and-combo-box.cs` | Use FormField.Type enumeration to differentiate between text input, check box, and combo box fields. | dropdown-form-field | docx | mcp |
| 19 | `iterate-over-all-form-fields-in-a-document-and-list-each-field-s-name-and-type.cs` | Iterate over all form fields in a document and list each field's name and type. | inspection-and-reporting | docx | existing_repo |
| 20 | `count-the-number-of-each-form-field-type-by-iterating-through-the-formfields-collection.cs` | Count the number of each form field type by iterating through the FormFields collection. | general-form-field | docx | mcp |
| 21 | `iterate-through-form-fields-and-log-each-field-s-result-value-for-debugging-purposes.cs` | Iterate through form fields and log each field's Result value for debugging purposes. | general-form-field | docx | mcp |
| 22 | `extract-automatically-generated-bookmark-names-for-all-form-fields-and-store-them-in-a-loo.cs` | Extract automatically generated bookmark names for all form fields and store them in a lookup dictionary. | general-form-field | docx | llm |
| 23 | `navigate-to-a-form-field-using-its-automatically-created-bookmark-and-extract-surrounding.cs` | Navigate to a form field using its automatically created bookmark and extract surrounding paragraph text. | general-form-field | docx | existing_repo |
| 24 | `protect-a-word-document-with-allowonlyformfields-option-allowing-only-form-fields-to-be-ed.cs` | Protect a Word document with AllowOnlyFormFields option, allowing only form fields to be edited. | general-form-field | docx | mcp |
| 25 | `apply-document-protection-levels-ensuring-only-form-fields-remain-editable-while-other-sec.cs` | Apply document protection levels ensuring only form fields remain editable while other sections are read-only. | general-form-field | docx | existing_repo |
| 26 | `verify-that-protected-document-still-permits-editing-of-form-fields-after-saving-and-reope.cs` | Verify that protected document still permits editing of form fields after saving and reopening. | general-form-field | docx | llm |
| 27 | `after-inserting-fields-protect-the-document-and-then-verify-that-non-field-content-cannot.cs` | After inserting fields, protect the document and then verify that non-field content cannot be edited. | general-form-field | docx | llm |
| 28 | `disable-all-check-box-form-fields-programmatically-when-a-specific-document-condition-is-m.cs` | Disable all check box form fields programmatically when a specific document condition is met. | general-form-field | docx | mcp |
| 29 | `implement-error-handling-for-attempts-to-insert-a-form-field-with-an-empty-name-logging-a.cs` | Implement error handling for attempts to insert a form field with an empty name, logging a warning. | general-form-field | docx | llm |
| 30 | `build-a-console-application-that-reads-a-csv-file-and-populates-corresponding-text-input-f.cs` | Build a console application that reads a CSV file and populates corresponding text input fields in a template. | text-form-field | csv | mcp |

## Common failure patterns seen during generation and how they were corrected

### Incorrect API signatures

- Symptom: Compilation errors such as CS1503 and CS7036 indicate mismatched argument types or missing required parameters.
- Root cause: The code supplies arguments in the wrong order, uses wrong data types, or omits required parameters like maxLength for InsertTextInput or selectedIndex for InsertComboBox.
- Preferred fix: Consult the Aspose.Words documentation and pass arguments exactly as defined: name (string), TextFormFieldType, format (string), default text (string), maxLength (int) for InsertTextInput; name, isChecked (bool), size (int) for InsertCheckBox; name, items (string[]), selectedIndex (int) for InsertComboBox.
- Prevention: Write a small wrapper that calls the overload with named parameters and compile-time checks to ensure all required arguments are provided.

### Missing or incorrect namespaces

- Symptom: CS0246 and CS0234 errors show that types like FormField, TextFormFieldType, and ProtectionType cannot be resolved.
- Root cause: The source file lacks the required using directives or references the wrong Aspose.Words sub-namespaces (e.g., Aspose.Words.Building, Aspose.Words.Forms).
- Preferred fix: Add using Aspose.Words; using Aspose.Words.Fields; and ensure the project references Aspose.Words.dll version 26.10.0.
- Prevention: Include a standard header block with the necessary using statements in every form-field example template.

### Mixing legacy fields with content controls

- Symptom: Examples that import StructuredDocumentTag or other content-control classes fail to compile or behave unexpectedly.
- Root cause: Legacy form fields and content controls are distinct APIs; the task explicitly requires legacy fields.
- Preferred fix: Remove all StructuredDocumentTag usage and rely solely on DocumentBuilder legacy insertion methods and the FormFields collection.
- Prevention: Document the distinction in the coding guidelines and enforce a lint rule that flags any usage of StructuredDocumentTag in form-field samples.

### Assuming pre-existing fields

- Symptom: Runtime code attempts to read or modify a field that was never inserted, leading to null-reference exceptions.
- Root cause: The sample code starts by loading a document and directly accesses FormFields without first creating them.
- Preferred fix: Bootstrap a minimal document that inserts the required fields before any read or update logic.
- Prevention: Include a step in the workflow that always creates a fresh document or explicitly checks for field existence and throws a clear error if missing.

### Invalid protection constants

- Symptom: CS0117 errors indicate that ProtectionType.Forms does not exist.
- Root cause: The correct enum value is ProtectionType.AllowOnlyFormFields; using an outdated or misspelled constant causes compilation failure.
- Preferred fix: Replace ProtectionType.Forms with ProtectionType.AllowOnlyFormFields when calling Document.Protect.
- Prevention: Reference the official enum list in the documentation and add a unit test that verifies document protection can be applied without errors.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Update the example repository whenever Aspose.Words introduces new overloads or deprecates legacy form-field APIs to keep the samples current.
