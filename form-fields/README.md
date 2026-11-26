# Form Fields Examples for Aspose.Words for .NET

The Form Fields category covers creation, manipulation, and inspection of legacy form fields in Word documents using Aspose.Words 26.11.0. Verified examples demonstrate inserting text inputs, check boxes, combo boxes, batch processing, JSON-driven updates, and document protection while adhering to the prescribed workflow.

## Snapshot

- Category: Form Fields
- Slug: form-fields
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261124_190740_1c00f5
- Text Form Field examples: 7
- General Form Field examples: 17
- Dropdown Form Field examples: 4
- Inspection And Reporting examples: 2

## Category rules that shaped these examples

- Do not assume that form fields already exist in a document before accessing them.
- Do not use StructuredDocumentTag or content-control APIs when the task targets legacy form fields.
- Do not reference helper types that are not part of Aspose.Words, such as TextFormFieldType from non-existent namespaces.
- Create form fields with DocumentBuilder.InsertTextInput, InsertCheckBox, and InsertComboBox, supplying all required parameters.
- Retrieve and modify fields via Document.Range.FormFields, using Result for text fields, Checked for check boxes, and Result for combo boxes.
- Assign a distinct name to each field; Aspose.Words automatically creates a bookmark with the same identifier.
- Check that a FormField retrieved from the collection is not null before dereferencing its properties.
- Guard against missing values when parsing external data, such as JSON, by validating deserialization results.
- Avoid implicit conversions that could produce CS8600, CS8602, or CS8604 warnings by using explicit casts or matching types.

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
dotnet add package Aspose.Words --version 26.11.0

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
| 5 | `insert-a-text-input-field-with-maximum-length-of-50-characters-and-numeric-format-using-do.cs` | Insert a text input field with maximum length of 50 characters and numeric format using DocumentBuilder. | text-form-field | docx | mcp |
| 6 | `insert-a-text-input-field-with-custom-date-format-and-default-current-date-value-using-doc.cs` | Insert a text input field with custom date format and default current date value using DocumentBuilder. | text-form-field | docx | mcp |
| 7 | `implement-a-reusable-method-that-adds-a-combo-box-form-field-with-customizable-items-and-d.cs` | Implement a reusable method that adds a combo box form field with customizable items and default index. | dropdown-form-field | docx | mcp |
| 8 | `assign-a-unique-name-to-each-inserted-form-field-to-ensure-distinct-automatic-bookmarks.cs` | Assign a unique name to each inserted form field to ensure distinct automatic bookmarks. | general-form-field | docx | mcp |
| 9 | `batch-insert-multiple-text-input-form-fields-into-a-template-using-a-loop-over-field-defin.cs` | Batch insert multiple text input form fields into a template using a loop over field definitions. | text-form-field | docx | llm |
| 10 | `load-a-docx-file-change-combo-box-selections-programmatically-and-save-the-modified-docume.cs` | Load a DOCX file, change combo box selections programmatically, and save the modified document. | dropdown-form-field | docx | mcp |
| 11 | `update-result-values-of-several-check-box-form-fields-based-on-external-json-configuration.cs` | Update result values of several check box form fields based on external JSON configuration data. | inspection-and-reporting | json | mcp |
| 12 | `set-the-result-property-of-a-text-input-form-field-to-a-predefined-string-value.cs` | Set the Result property of a text input form field to a predefined string value. | text-form-field | docx | mcp |
| 13 | `read-the-result-property-of-a-check-box-form-field-to-determine-whether-it-is-checked.cs` | Read the Result property of a check box form field to determine whether it is checked. | general-form-field | docx | mcp |
| 14 | `toggle-the-checked-state-of-a-specific-check-box-form-field-based-on-external-configuratio.cs` | Toggle the checked state of a specific check box form field based on external configuration settings. | general-form-field | docx | mcp |
| 15 | `change-the-size-of-an-existing-check-box-form-field-programmatically-to-improve-visual-con.cs` | Change the size of an existing check box form field programmatically to improve visual consistency. | general-form-field | docx | existing_repo |
| 16 | `retrieve-a-form-field-by-its-name-from-a-loaded-document-and-read-its-result-property.cs` | Retrieve a form field by its name from a loaded document and read its Result property. | general-form-field | docx | mcp |
| 17 | `access-a-form-field-by-index-from-the-formfields-collection-and-modify-its-result-value.cs` | Access a form field by index from the FormFields collection and modify its Result value. | general-form-field | docx | existing_repo |
| 18 | `use-formfield-type-enumeration-to-differentiate-between-text-input-check-box-and-combo-box.cs` | Use FormField.Type enumeration to differentiate between text input, check box, and combo box fields. | dropdown-form-field | docx | llm |
| 19 | `iterate-over-all-form-fields-in-a-document-and-list-each-field-s-name-and-type.cs` | Iterate over all form fields in a document and list each field's name and type. | inspection-and-reporting | docx | llm |
| 20 | `count-the-number-of-each-form-field-type-by-iterating-through-the-formfields-collection.cs` | Count the number of each form field type by iterating through the FormFields collection. | general-form-field | docx | mcp |
| 21 | `iterate-through-form-fields-and-log-each-field-s-result-value-for-debugging-purposes.cs` | Iterate through form fields and log each field's Result value for debugging purposes. | general-form-field | docx | mcp |
| 22 | `extract-automatically-generated-bookmark-names-for-all-form-fields-and-store-them-in-a-loo.cs` | Extract automatically generated bookmark names for all form fields and store them in a lookup dictionary. | general-form-field | docx | mcp |
| 23 | `navigate-to-a-form-field-using-its-automatically-created-bookmark-and-extract-surrounding.cs` | Navigate to a form field using its automatically created bookmark and extract surrounding paragraph text. | general-form-field | docx | llm |
| 24 | `protect-a-word-document-with-allowonlyformfields-option-allowing-only-form-fields-to-be-ed.cs` | Protect a Word document with AllowOnlyFormFields option, allowing only form fields to be edited. | general-form-field | docx | existing_repo |
| 25 | `apply-document-protection-levels-ensuring-only-form-fields-remain-editable-while-other-sec.cs` | Apply document protection levels ensuring only form fields remain editable while other sections are read-only. | general-form-field | docx | mcp |
| 26 | `verify-that-protected-document-still-permits-editing-of-form-fields-after-saving-and-reope.cs` | Verify that protected document still permits editing of form fields after saving and reopening. | general-form-field | docx | llm |
| 27 | `after-inserting-fields-protect-the-document-and-then-verify-that-non-field-content-cannot.cs` | After inserting fields, protect the document and then verify that non-field content cannot be edited. | general-form-field | docx | llm |
| 28 | `disable-all-check-box-form-fields-programmatically-when-a-specific-document-condition-is-m.cs` | Disable all check box form fields programmatically when a specific document condition is met. | general-form-field | docx | mcp |
| 29 | `implement-error-handling-for-attempts-to-insert-a-form-field-with-an-empty-name-logging-a.cs` | Implement error handling for attempts to insert a form field with an empty name, logging a warning. | general-form-field | docx | llm |
| 30 | `build-a-console-application-that-reads-a-csv-file-and-populates-corresponding-text-input-f.cs` | Build a console application that reads a CSV file and populates corresponding text input fields in a template. | text-form-field | csv | llm |

## Common failure patterns seen during generation and how they were corrected

### Incorrect overload usage for Insert* methods

- Symptom: Compilation fails with CS1503, CS1501, CS7036, or CS1509 errors because arguments of the wrong type or count are supplied.
- Root cause: The code calls InsertTextInput, InsertCheckBox, or InsertComboBox with parameters that do not match the documented signatures, such as passing a bool where TextFormFieldType is required or omitting the maxLength argument.
- Preferred fix: Use the exact overloads: InsertTextInput(string name, TextFormFieldType type, string format, string defaultText, int maxLength); InsertCheckBox(string name, bool isChecked, int size); InsertComboBox(string name, string[] items, int selectedIndex).
- Prevention: Consult the Aspose.Words API reference for each method and match parameter types and order precisely.

### Missing required namespaces or referencing non-existent types

- Symptom: Compilation produces CS0103, CS0246, or CS0234 errors indicating that identifiers like TextFormFieldType, FormFieldCollection, or Aspose.Words.Building are unknown.
- Root cause: The source file lacks using directives for Aspose.Words.Fields and other necessary namespaces, or it references types that are not part of the installed Aspose.Words version.
- Preferred fix: Add using Aspose.Words.Fields; and ensure the project references Aspose.Words 26.11.0, which defines TextFormFieldType, FormField, and related classes.
- Prevention: Always include the required namespaces listed in the category metadata before writing code.

### Mixing legacy form fields with content controls

- Symptom: Examples combine StructuredDocumentTag with InsertTextInput or InsertCheckBox, leading to logical errors or unsupported API calls.
- Root cause: A misunderstanding of the distinction between legacy form fields and modern content controls causes inappropriate API usage.
- Preferred fix: Restrict the solution to legacy form-field APIs (DocumentBuilder.Insert* and Range.FormFields) unless the task explicitly demands content controls.
- Prevention: Follow the canonical pattern that separates legacy field handling from StructuredDocumentTag usage.

### Insufficient validation of field existence and value changes

- Symptom: Code saves the document without confirming that a form field was created or that its Result/Checked property was updated, potentially leaving the document in an incorrect state.
- Root cause: Missing null checks and post-assignment assertions after accessing FormFields.
- Preferred fix: After retrieving a field, verify it is not null, then compare the property value after assignment; throw an exception or log an error if the check fails.
- Prevention: Incorporate explicit existence checks and value verification steps as part of every example.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Keep the example code aligned with the Aspose.Words 26.11.0 API surface; update signatures and namespace imports if a newer version is adopted.
