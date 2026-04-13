# Form Fields Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Form Fields category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Form Fields**
- Slug: **form-fields**
- Total examples: **30**
- Publish-ready successful examples: **30 / 30**
- Text form-field examples: **7**
- Dropdown form-field examples: **4**
- Inspection / reporting examples: **2**
- General form-field examples: **17**

## Category rules that shaped these examples

- Use native Aspose.Words legacy form-field APIs directly.
- Use DocumentBuilder.InsertTextInput, InsertCheckBox, and InsertComboBox to create fields.
- Access legacy fields through Document.Range.FormFields.
- Do not mix StructuredDocumentTag unless the task explicitly requires content controls.
- Create realistic local sample inputs whenever the task mentions an existing document or pre-existing fields.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

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
dotnet add package Aspose.Words --version 26.3.0

# PowerShell example
Copy-Item ..\form-fields\insert-a-text-input-form-field-with-placeholder-text-using-documentbuilder-in-a-new-docume.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `insert-a-text-input-form-field-with-placeholder-text-using-documentbuilder-in-a-new-docume.cs` | Insert a text input form field with placeholder text using DocumentBuilder in a new document. | text-form-field | docx | verified |
| 2 | `add-a-check-box-form-field-set-default-checked-state-and-custom-size-via-documentbuilder.cs` | Add a check box form field, set default checked state and custom size via DocumentBuilder. | general-form-field | docx | verified |
| 3 | `create-a-combo-box-form-field-containing-three-items-and-set-default-selected-index-using.cs` | Create a combo box form field containing three items and set default selected index using DocumentBuilder. | dropdown-form-field | docx | verified |
| 4 | `specify-a-name-when-inserting-a-text-input-field-to-automatically-generate-a-matching-book.cs` | Specify a name when inserting a text input field to automatically generate a matching bookmark. | text-form-field | docx | verified |
| 5 | `insert-a-text-input-field-with-maximum-length-of-50-characters-and-numeric-format-using-do.cs` | Insert a text input field with maximum length of 50 characters and numeric format using DocumentBuilder. | text-form-field | docx | verified |
| 6 | `insert-a-text-input-field-with-custom-date-format-and-default-current-date-value-using-doc.cs` | Insert a text input field with custom date format and default current date value using DocumentBuilder. | text-form-field | docx | verified |
| 7 | `implement-a-reusable-method-that-adds-a-combo-box-form-field-with-customizable-items-and-d.cs` | Implement a reusable method that adds a combo box form field with customizable items and default index. | dropdown-form-field | docx | verified |
| 8 | `assign-a-unique-name-to-each-inserted-form-field-to-ensure-distinct-automatic-bookmarks.cs` | Assign a unique name to each inserted form field to ensure distinct automatic bookmarks. | general-form-field | docx | verified |
| 9 | `batch-insert-multiple-text-input-form-fields-into-a-template-using-a-loop-over-field-defin.cs` | Batch insert multiple text input form fields into a template using a loop over field definitions. | text-form-field | docx | verified |
| 10 | `load-a-docx-file-change-combo-box-selections-programmatically-and-save-the-modified-docume.cs` | Load a DOCX file, change combo box selections programmatically, and save the modified document. | dropdown-form-field | docx | verified |
| 11 | `update-result-values-of-several-check-box-form-fields-based-on-external-json-configuration.cs` | Update result values of several check box form fields based on external JSON configuration data. | inspection-and-reporting | json | verified |
| 12 | `set-the-result-property-of-a-text-input-form-field-to-a-predefined-string-value.cs` | Set the Result property of a text input form field to a predefined string value. | text-form-field | docx | verified |
| 13 | `read-the-result-property-of-a-check-box-form-field-to-determine-whether-it-is-checked.cs` | Read the Result property of a check box form field to determine whether it is checked. | general-form-field | docx | verified |
| 14 | `toggle-the-checked-state-of-a-specific-check-box-form-field-based-on-external-configuratio.cs` | Toggle the checked state of a specific check box form field based on external configuration settings. | general-form-field | docx | verified |
| 15 | `change-the-size-of-an-existing-check-box-form-field-programmatically-to-improve-visual-con.cs` | Change the size of an existing check box form field programmatically to improve visual consistency. | general-form-field | docx | verified |
| 16 | `retrieve-a-form-field-by-its-name-from-a-loaded-document-and-read-its-result-property.cs` | Retrieve a form field by its name from a loaded document and read its Result property. | general-form-field | docx | verified |
| 17 | `access-a-form-field-by-index-from-the-formfields-collection-and-modify-its-result-value.cs` | Access a form field by index from the FormFields collection and modify its Result value. | general-form-field | docx | verified |
| 18 | `use-formfield-type-enumeration-to-differentiate-between-text-input-check-box-and-combo-box.cs` | Use FormField.Type enumeration to differentiate between text input, check box, and combo box fields. | dropdown-form-field | docx | verified |
| 19 | `iterate-over-all-form-fields-in-a-document-and-list-each-field-s-name-and-type.cs` | Iterate over all form fields in a document and list each field's name and type. | inspection-and-reporting | docx | verified |
| 20 | `count-the-number-of-each-form-field-type-by-iterating-through-the-formfields-collection.cs` | Count the number of each form field type by iterating through the FormFields collection. | general-form-field | docx | verified |
| 21 | `iterate-through-form-fields-and-log-each-field-s-result-value-for-debugging-purposes.cs` | Iterate through form fields and log each field's Result value for debugging purposes. | general-form-field | docx | verified |
| 22 | `extract-automatically-generated-bookmark-names-for-all-form-fields-and-store-them-in-a-loo.cs` | Extract automatically generated bookmark names for all form fields and store them in a lookup dictionary. | general-form-field | docx | verified |
| 23 | `navigate-to-a-form-field-using-its-automatically-created-bookmark-and-extract-surrounding.cs` | Navigate to a form field using its automatically created bookmark and extract surrounding paragraph text. | general-form-field | docx | verified |
| 24 | `protect-a-word-document-with-allowonlyformfields-option-allowing-only-form-fields-to-be-ed.cs` | Protect a Word document with AllowOnlyFormFields option, allowing only form fields to be edited. | general-form-field | docx | verified |
| 25 | `apply-document-protection-levels-ensuring-only-form-fields-remain-editable-while-other-sec.cs` | Apply document protection levels ensuring only form fields remain editable while other sections are read‑only. | general-form-field | docx | verified |
| 26 | `verify-that-protected-document-still-permits-editing-of-form-fields-after-saving-and-reope.cs` | Verify that protected document still permits editing of form fields after saving and reopening. | general-form-field | docx | verified |
| 27 | `after-inserting-fields-protect-the-document-and-then-verify-that-non-field-content-cannot.cs` | After inserting fields, protect the document and then verify that non‑field content cannot be edited. | general-form-field | docx | verified |
| 28 | `disable-all-check-box-form-fields-programmatically-when-a-specific-document-condition-is-m.cs` | Disable all check box form fields programmatically when a specific document condition is met. | general-form-field | docx | verified |
| 29 | `implement-error-handling-for-attempts-to-insert-a-form-field-with-an-empty-name-logging-a.cs` | Implement error handling for attempts to insert a form field with an empty name, logging a warning. | general-form-field | docx | verified |
| 30 | `build-a-console-application-that-reads-a-csv-file-and-populates-corresponding-text-input-f.cs` | Build a console application that reads a CSV file and populates corresponding text input fields in a template. | text-form-field | csv | verified |

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

## Notes for maintainers

- This category is now **100% publish-ready** for the current run.
- Preserve file-to-task traceability when updating this folder.
- For future updates, keep the examples standalone and continue bootstrapping local inputs inside the example whenever external sources are mentioned.
