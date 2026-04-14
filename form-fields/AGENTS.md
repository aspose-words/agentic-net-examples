---
name: form-fields
description: Verified C# examples for form field scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Form Fields

## Purpose

This folder is a live, curated example set for form field scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free creation, reading, updating, validation, and reporting of legacy form fields using direct Aspose.Words APIs.

## Non-negotiable conventions

- Use native Aspose.Words legacy form-field APIs directly.
- Use DocumentBuilder.InsertTextInput, InsertCheckBox, and InsertComboBox to create fields.
- Access existing legacy fields through Document.Range.FormFields.
- Do not mix StructuredDocumentTag unless the task explicitly requires content controls.
- Bootstrap local source documents whenever the task implies an existing file or pre-existing fields.
- Guard maybe-null values to avoid nullable-reference warnings such as CS8600, CS8602, and CS8604.

## Recommended workflow selection

- **Text form-field workflow**: 7 examples
- **Dropdown form-field workflow**: 4 examples
- **Inspection / reporting workflow**: 2 examples
- **General form-field workflow**: 17 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. Source documents must be bootstrapped locally whenever the task implies an existing input.
3. Legacy form fields must be created, found, or updated through the correct Aspose.Words APIs.
4. Requested report or output files must actually be written.
5. Field existence and final values must be validated when the task expects a changed result.

## File-to-task reference

- `insert-a-text-input-form-field-with-placeholder-text-using-documentbuilder-in-a-new-docume.cs`
  - Task: Insert a text input form field with placeholder text using DocumentBuilder in a new document.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: verified
- `add-a-check-box-form-field-set-default-checked-state-and-custom-size-via-documentbuilder.cs`
  - Task: Add a check box form field, set default checked state and custom size via DocumentBuilder.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `create-a-combo-box-form-field-containing-three-items-and-set-default-selected-index-using.cs`
  - Task: Create a combo box form field containing three items and set default selected index using DocumentBuilder.
  - Workflow: dropdown-form-field
  - Outputs: docx
  - Selected engine: verified
- `specify-a-name-when-inserting-a-text-input-field-to-automatically-generate-a-matching-book.cs`
  - Task: Specify a name when inserting a text input field to automatically generate a matching bookmark.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: verified
- `insert-a-text-input-field-with-maximum-length-of-50-characters-and-numeric-format-using-do.cs`
  - Task: Insert a text input field with maximum length of 50 characters and numeric format using DocumentBuilder.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: verified
- `insert-a-text-input-field-with-custom-date-format-and-default-current-date-value-using-doc.cs`
  - Task: Insert a text input field with custom date format and default current date value using DocumentBuilder.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: verified
- `implement-a-reusable-method-that-adds-a-combo-box-form-field-with-customizable-items-and-d.cs`
  - Task: Implement a reusable method that adds a combo box form field with customizable items and default index.
  - Workflow: dropdown-form-field
  - Outputs: docx
  - Selected engine: verified
- `assign-a-unique-name-to-each-inserted-form-field-to-ensure-distinct-automatic-bookmarks.cs`
  - Task: Assign a unique name to each inserted form field to ensure distinct automatic bookmarks.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `batch-insert-multiple-text-input-form-fields-into-a-template-using-a-loop-over-field-defin.cs`
  - Task: Batch insert multiple text input form fields into a template using a loop over field definitions.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: verified
- `load-a-docx-file-change-combo-box-selections-programmatically-and-save-the-modified-docume.cs`
  - Task: Load a DOCX file, change combo box selections programmatically, and save the modified document.
  - Workflow: dropdown-form-field
  - Outputs: docx
  - Selected engine: verified
- `update-result-values-of-several-check-box-form-fields-based-on-external-json-configuration.cs`
  - Task: Update result values of several check box form fields based on external JSON configuration data.
  - Workflow: inspection-and-reporting
  - Outputs: json
  - Selected engine: verified
- `set-the-result-property-of-a-text-input-form-field-to-a-predefined-string-value.cs`
  - Task: Set the Result property of a text input form field to a predefined string value.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: verified
- `read-the-result-property-of-a-check-box-form-field-to-determine-whether-it-is-checked.cs`
  - Task: Read the Result property of a check box form field to determine whether it is checked.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `toggle-the-checked-state-of-a-specific-check-box-form-field-based-on-external-configuratio.cs`
  - Task: Toggle the checked state of a specific check box form field based on external configuration settings.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `change-the-size-of-an-existing-check-box-form-field-programmatically-to-improve-visual-con.cs`
  - Task: Change the size of an existing check box form field programmatically to improve visual consistency.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `retrieve-a-form-field-by-its-name-from-a-loaded-document-and-read-its-result-property.cs`
  - Task: Retrieve a form field by its name from a loaded document and read its Result property.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `access-a-form-field-by-index-from-the-formfields-collection-and-modify-its-result-value.cs`
  - Task: Access a form field by index from the FormFields collection and modify its Result value.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `use-formfield-type-enumeration-to-differentiate-between-text-input-check-box-and-combo-box.cs`
  - Task: Use FormField.Type enumeration to differentiate between text input, check box, and combo box fields.
  - Workflow: dropdown-form-field
  - Outputs: docx
  - Selected engine: verified
- `iterate-over-all-form-fields-in-a-document-and-list-each-field-s-name-and-type.cs`
  - Task: Iterate over all form fields in a document and list each field's name and type.
  - Workflow: inspection-and-reporting
  - Outputs: docx
  - Selected engine: verified
- `count-the-number-of-each-form-field-type-by-iterating-through-the-formfields-collection.cs`
  - Task: Count the number of each form field type by iterating through the FormFields collection.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `iterate-through-form-fields-and-log-each-field-s-result-value-for-debugging-purposes.cs`
  - Task: Iterate through form fields and log each field's Result value for debugging purposes.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `extract-automatically-generated-bookmark-names-for-all-form-fields-and-store-them-in-a-loo.cs`
  - Task: Extract automatically generated bookmark names for all form fields and store them in a lookup dictionary.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `navigate-to-a-form-field-using-its-automatically-created-bookmark-and-extract-surrounding.cs`
  - Task: Navigate to a form field using its automatically created bookmark and extract surrounding paragraph text.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `protect-a-word-document-with-allowonlyformfields-option-allowing-only-form-fields-to-be-ed.cs`
  - Task: Protect a Word document with AllowOnlyFormFields option, allowing only form fields to be edited.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `apply-document-protection-levels-ensuring-only-form-fields-remain-editable-while-other-sec.cs`
  - Task: Apply document protection levels ensuring only form fields remain editable while other sections are read‑only.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `verify-that-protected-document-still-permits-editing-of-form-fields-after-saving-and-reope.cs`
  - Task: Verify that protected document still permits editing of form fields after saving and reopening.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `after-inserting-fields-protect-the-document-and-then-verify-that-non-field-content-cannot.cs`
  - Task: After inserting fields, protect the document and then verify that non‑field content cannot be edited.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `disable-all-check-box-form-fields-programmatically-when-a-specific-document-condition-is-m.cs`
  - Task: Disable all check box form fields programmatically when a specific document condition is met.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `implement-error-handling-for-attempts-to-insert-a-form-field-with-an-empty-name-logging-a.cs`
  - Task: Implement error handling for attempts to insert a form field with an empty name, logging a warning.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: verified
- `build-a-console-application-that-reads-a-csv-file-and-populates-corresponding-text-input-f.cs`
  - Task: Build a console application that reads a CSV file and populates corresponding text input fields in a template.
  - Workflow: text-form-field
  - Outputs: csv
  - Selected engine: verified

## Common failure patterns and preferred agent fixes

- **Mixing legacy form fields with content controls**
  - Symptom: Examples use StructuredDocumentTag when the task is about legacy form fields.
  - Preferred fix: Use DocumentBuilder legacy form-field APIs and Document.Range.FormFields unless the task explicitly requires content controls.

- **Wrong field access pattern**
  - Symptom: Code assumes unsupported helpers or reads the wrong property for the field type.
  - Preferred fix: Access legacy fields through Range.FormFields and use Result, Checked, or DropDownItems according to the field type.

- **Missing bootstrap input**
  - Symptom: Examples assume an existing document with form fields already exists.
  - Preferred fix: Create a local sample document and seed the required form fields before reading or updating them.

- **Weak validation**
  - Symptom: The example saves output but does not verify that the form field exists or that its value changed.
  - Preferred fix: Validate field existence and the final assigned value before finishing.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.3.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\form-fields\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer direct Aspose.Words form-field APIs over speculative shortcuts.
