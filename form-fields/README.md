# Form Fields Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Form Fields category. Each file is a standalone console example selected from the verified 26.5.0 run.

## Snapshot

- Category: Form Fields
- Slug: form-fields
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20260619_131835_59df5f
- Dropdown Form Field examples: 4
- General Form Field examples: 17
- Inspection And Reporting examples: 2
- Text Form Field examples: 7

## Category rules that shaped these examples

- Do not assume form fields already exist in a document.
- Do not mix StructuredDocumentTag unless the task explicitly requires content controls.
- Do not invent unsupported form-field helper APIs.
- Do not skip validation of field existence when the task expects an existing field.
- Use DocumentBuilder.InsertTextInput, InsertCheckBox, and InsertComboBox to create legacy form fields.
- Access legacy fields through Document.Range.FormFields.
- Use FormField.Result for text-based values and FormField.Checked for checkboxes.
- Bootstrap local source documents whenever the task implies an existing file or pre-existing fields.
- Check for null before accessing form fields or lookup results.
- Avoid CS8600, CS8602, and CS8604 by guarding values before dereference.
- Do not rely on maybe-null field lookups without validation.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.5.0

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

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
dotnet add package Aspose.Words --version 26.5.0

# PowerShell example
Copy-Item ..\form-fields\insert-a-text-input-form-field-with-placeholder-text-using-documentbuilder-in-a-new-docume.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `insert-a-text-input-form-field-with-placeholder-text-using-documentbuilder-in-a-new-docume.cs` | Insert a text input form field with placeholder text using DocumentBuilder in a new document. | Text Form Field | docx | mcp |
| 2 | `add-a-check-box-form-field-set-default-checked-state-and-custom-size-via-documentbuilder.cs` | Add a check box form field, set default checked state and custom size via DocumentBuilder. | General Form Field | docx | mcp |
| 3 | `create-a-combo-box-form-field-containing-three-items-and-set-default-selected-index-using.cs` | Create a combo box form field containing three items and set default selected index using DocumentBuilder. | Dropdown Form Field | docx | mcp |
| 4 | `specify-a-name-when-inserting-a-text-input-field-to-automatically-generate-a-matching-book.cs` | Specify a name when inserting a text input field to automatically generate a matching bookmark. | Text Form Field | docx | mcp |
| 5 | `insert-a-text-input-field-with-maximum-length-of-50-characters-and-numeric-format-using-do.cs` | Insert a text input field with maximum length of 50 characters and numeric format using DocumentBuilder. | Text Form Field | docx | mcp |
| 6 | `insert-a-text-input-field-with-custom-date-format-and-default-current-date-value-using-doc.cs` | Insert a text input field with custom date format and default current date value using DocumentBuilder. | Text Form Field | docx | mcp |
| 7 | `implement-a-reusable-method-that-adds-a-combo-box-form-field-with-customizable-items-and-d.cs` | Implement a reusable method that adds a combo box form field with customizable items and default index. | Dropdown Form Field | docx | mcp |
| 8 | `assign-a-unique-name-to-each-inserted-form-field-to-ensure-distinct-automatic-bookmarks.cs` | Assign a unique name to each inserted form field to ensure distinct automatic bookmarks. | General Form Field | docx | mcp |
| 9 | `batch-insert-multiple-text-input-form-fields-into-a-template-using-a-loop-over-field-defin.cs` | Batch insert multiple text input form fields into a template using a loop over field definitions. | Text Form Field | docx | mcp |
| 10 | `load-a-docx-file-change-combo-box-selections-programmatically-and-save-the-modified-docume.cs` | Load a DOCX file, change combo box selections programmatically, and save the modified document. | Dropdown Form Field | docx | mcp |
| 11 | `update-result-values-of-several-check-box-form-fields-based-on-external-json-configuration.cs` | Update result values of several check box form fields based on external JSON configuration data. | Inspection And Reporting | json | mcp |
| 12 | `set-the-result-property-of-a-text-input-form-field-to-a-predefined-string-value.cs` | Set the Result property of a text input form field to a predefined string value. | Text Form Field | docx | mcp |
| 13 | `read-the-result-property-of-a-check-box-form-field-to-determine-whether-it-is-checked.cs` | Read the Result property of a check box form field to determine whether it is checked. | General Form Field | docx | mcp |
| 14 | `toggle-the-checked-state-of-a-specific-check-box-form-field-based-on-external-configuratio.cs` | Toggle the checked state of a specific check box form field based on external configuration settings. | General Form Field | docx | mcp |
| 15 | `change-the-size-of-an-existing-check-box-form-field-programmatically-to-improve-visual-con.cs` | Change the size of an existing check box form field programmatically to improve visual consistency. | General Form Field | docx | mcp |
| 16 | `retrieve-a-form-field-by-its-name-from-a-loaded-document-and-read-its-result-property.cs` | Retrieve a form field by its name from a loaded document and read its Result property. | General Form Field | docx | mcp |
| 17 | `access-a-form-field-by-index-from-the-formfields-collection-and-modify-its-result-value.cs` | Access a form field by index from the FormFields collection and modify its Result value. | General Form Field | docx | mcp |
| 18 | `use-formfield-type-enumeration-to-differentiate-between-text-input-check-box-and-combo-box.cs` | Use FormField.Type enumeration to differentiate between text input, check box, and combo box fields. | Dropdown Form Field | docx | mcp |
| 19 | `iterate-over-all-form-fields-in-a-document-and-list-each-field-s-name-and-type.cs` | Iterate over all form fields in a document and list each field's name and type. | Inspection And Reporting | docx | mcp |
| 20 | `count-the-number-of-each-form-field-type-by-iterating-through-the-formfields-collection.cs` | Count the number of each form field type by iterating through the FormFields collection. | General Form Field | docx | mcp |
| 21 | `iterate-through-form-fields-and-log-each-field-s-result-value-for-debugging-purposes.cs` | Iterate through form fields and log each field's Result value for debugging purposes. | General Form Field | docx | mcp |
| 22 | `extract-automatically-generated-bookmark-names-for-all-form-fields-and-store-them-in-a-loo.cs` | Extract automatically generated bookmark names for all form fields and store them in a lookup dictionary. | General Form Field | docx | mcp |
| 23 | `navigate-to-a-form-field-using-its-automatically-created-bookmark-and-extract-surrounding.cs` | Navigate to a form field using its automatically created bookmark and extract surrounding paragraph text. | General Form Field | docx | mcp |
| 24 | `protect-a-word-document-with-allowonlyformfields-option-allowing-only-form-fields-to-be-ed.cs` | Protect a Word document with AllowOnlyFormFields option, allowing only form fields to be edited. | General Form Field | docx | mcp |
| 25 | `apply-document-protection-levels-ensuring-only-form-fields-remain-editable-while-other-sec.cs` | Apply document protection levels ensuring only form fields remain editable while other sections are read-only. | General Form Field | docx | mcp |
| 26 | `verify-that-protected-document-still-permits-editing-of-form-fields-after-saving-and-reope.cs` | Verify that protected document still permits editing of form fields after saving and reopening. | General Form Field | docx | mcp |
| 27 | `after-inserting-fields-protect-the-document-and-then-verify-that-non-field-content-cannot.cs` | After inserting fields, protect the document and then verify that non-field content cannot be edited. | General Form Field | docx | mcp |
| 28 | `disable-all-check-box-form-fields-programmatically-when-a-specific-document-condition-is-m.cs` | Disable all check box form fields programmatically when a specific document condition is met. | General Form Field | docx | mcp |
| 29 | `implement-error-handling-for-attempts-to-insert-a-form-field-with-an-empty-name-logging-a.cs` | Implement error handling for attempts to insert a form field with an empty name, logging a warning. | General Form Field | docx | mcp |
| 30 | `build-a-console-application-that-reads-a-csv-file-and-populates-corresponding-text-input-f.cs` | Build a console application that reads a CSV file and populates corresponding text input fields in a template. | Text Form Field | csv | mcp |

## Common failure patterns seen during generation and how they were corrected

### Mixing legacy form fields with content controls

- Symptom: Examples use StructuredDocumentTag when the task is about legacy form fields.
- Fix: Use DocumentBuilder legacy form-field APIs and Document.Range.FormFields unless the task explicitly requires content controls.

### Wrong field access pattern

- Symptom: Code assumes unsupported helpers or reads the wrong property for the field type.
- Fix: Access legacy fields through Range.FormFields and use Result, Checked, or DropDownItems according to the field type.

### Missing bootstrap input

- Symptom: Examples assume an existing document with form fields already exists.
- Fix: Create a local sample document and seed the required form fields before reading or updating them.

### Weak validation

- Symptom: The example saves output but does not verify that the form field exists or that its value changed.
- Fix: Validate field existence and the final assigned value before finishing.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.5.0 run.
- Preserve file-to-task traceability when updating this folder.
- Keep examples standalone and bootstrap local inputs inside the example whenever external sources are mentioned.
