---
name: form-fields
description: Verified C# examples for form-fields scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Form Fields

## Purpose

Provide an AI coding agent with concise guidance to generate correct, compile-time-safe C# snippets that create and manage Aspose.Words legacy form fields, validate their existence, and apply document protection without mixing unsupported APIs or content controls.

## Non-negotiable conventions

- Do not assume that form fields already exist in a document before accessing them.
- Do not use StructuredDocumentTag or content-control APIs when the task targets legacy form fields.
- Do not reference helper types that are not part of Aspose.Words, such as TextFormFieldType from non-existent namespaces.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Text Form Field workflow**: 7 examples
- **General Form Field workflow**: 17 examples
- **Dropdown Form Field workflow**: 4 examples
- **Inspection And Reporting workflow**: 2 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Validate that each inserted form field exists in the FormFields collection, has the expected name, and that its Result or Checked property reflects the intended value before the document is saved.

## File-to-task reference

- `insert-a-text-input-form-field-with-placeholder-text-using-documentbuilder-in-a-new-docume.cs`
  - Task: Insert a text input form field with placeholder text using DocumentBuilder in a new document.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: mcp
- `add-a-check-box-form-field-set-default-checked-state-and-custom-size-via-documentbuilder.cs`
  - Task: Add a check box form field, set default checked state and custom size via DocumentBuilder.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `create-a-combo-box-form-field-containing-three-items-and-set-default-selected-index-using.cs`
  - Task: Create a combo box form field containing three items and set default selected index using DocumentBuilder.
  - Workflow: dropdown-form-field
  - Outputs: docx
  - Selected engine: mcp
- `specify-a-name-when-inserting-a-text-input-field-to-automatically-generate-a-matching-book.cs`
  - Task: Specify a name when inserting a text input field to automatically generate a matching bookmark.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-text-input-field-with-maximum-length-of-50-characters-and-numeric-format-using-do.cs`
  - Task: Insert a text input field with maximum length of 50 characters and numeric format using DocumentBuilder.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: mcp
- `insert-a-text-input-field-with-custom-date-format-and-default-current-date-value-using-doc.cs`
  - Task: Insert a text input field with custom date format and default current date value using DocumentBuilder.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: mcp
- `implement-a-reusable-method-that-adds-a-combo-box-form-field-with-customizable-items-and-d.cs`
  - Task: Implement a reusable method that adds a combo box form field with customizable items and default index.
  - Workflow: dropdown-form-field
  - Outputs: docx
  - Selected engine: mcp
- `assign-a-unique-name-to-each-inserted-form-field-to-ensure-distinct-automatic-bookmarks.cs`
  - Task: Assign a unique name to each inserted form field to ensure distinct automatic bookmarks.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `batch-insert-multiple-text-input-form-fields-into-a-template-using-a-loop-over-field-defin.cs`
  - Task: Batch insert multiple text input form fields into a template using a loop over field definitions.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: llm
- `load-a-docx-file-change-combo-box-selections-programmatically-and-save-the-modified-docume.cs`
  - Task: Load a DOCX file, change combo box selections programmatically, and save the modified document.
  - Workflow: dropdown-form-field
  - Outputs: docx
  - Selected engine: mcp
- `update-result-values-of-several-check-box-form-fields-based-on-external-json-configuration.cs`
  - Task: Update result values of several check box form fields based on external JSON configuration data.
  - Workflow: inspection-and-reporting
  - Outputs: json
  - Selected engine: mcp
- `set-the-result-property-of-a-text-input-form-field-to-a-predefined-string-value.cs`
  - Task: Set the Result property of a text input form field to a predefined string value.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: mcp
- `read-the-result-property-of-a-check-box-form-field-to-determine-whether-it-is-checked.cs`
  - Task: Read the Result property of a check box form field to determine whether it is checked.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `toggle-the-checked-state-of-a-specific-check-box-form-field-based-on-external-configuratio.cs`
  - Task: Toggle the checked state of a specific check box form field based on external configuration settings.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `change-the-size-of-an-existing-check-box-form-field-programmatically-to-improve-visual-con.cs`
  - Task: Change the size of an existing check box form field programmatically to improve visual consistency.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: existing_repo
- `retrieve-a-form-field-by-its-name-from-a-loaded-document-and-read-its-result-property.cs`
  - Task: Retrieve a form field by its name from a loaded document and read its Result property.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `access-a-form-field-by-index-from-the-formfields-collection-and-modify-its-result-value.cs`
  - Task: Access a form field by index from the FormFields collection and modify its Result value.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: existing_repo
- `use-formfield-type-enumeration-to-differentiate-between-text-input-check-box-and-combo-box.cs`
  - Task: Use FormField.Type enumeration to differentiate between text input, check box, and combo box fields.
  - Workflow: dropdown-form-field
  - Outputs: docx
  - Selected engine: llm
- `iterate-over-all-form-fields-in-a-document-and-list-each-field-s-name-and-type.cs`
  - Task: Iterate over all form fields in a document and list each field's name and type.
  - Workflow: inspection-and-reporting
  - Outputs: docx
  - Selected engine: llm
- `count-the-number-of-each-form-field-type-by-iterating-through-the-formfields-collection.cs`
  - Task: Count the number of each form field type by iterating through the FormFields collection.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `iterate-through-form-fields-and-log-each-field-s-result-value-for-debugging-purposes.cs`
  - Task: Iterate through form fields and log each field's Result value for debugging purposes.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `extract-automatically-generated-bookmark-names-for-all-form-fields-and-store-them-in-a-loo.cs`
  - Task: Extract automatically generated bookmark names for all form fields and store them in a lookup dictionary.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `navigate-to-a-form-field-using-its-automatically-created-bookmark-and-extract-surrounding.cs`
  - Task: Navigate to a form field using its automatically created bookmark and extract surrounding paragraph text.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: llm
- `protect-a-word-document-with-allowonlyformfields-option-allowing-only-form-fields-to-be-ed.cs`
  - Task: Protect a Word document with AllowOnlyFormFields option, allowing only form fields to be edited.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: existing_repo
- `apply-document-protection-levels-ensuring-only-form-fields-remain-editable-while-other-sec.cs`
  - Task: Apply document protection levels ensuring only form fields remain editable while other sections are read-only.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `verify-that-protected-document-still-permits-editing-of-form-fields-after-saving-and-reope.cs`
  - Task: Verify that protected document still permits editing of form fields after saving and reopening.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: llm
- `after-inserting-fields-protect-the-document-and-then-verify-that-non-field-content-cannot.cs`
  - Task: After inserting fields, protect the document and then verify that non-field content cannot be edited.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: llm
- `disable-all-check-box-form-fields-programmatically-when-a-specific-document-condition-is-m.cs`
  - Task: Disable all check box form fields programmatically when a specific document condition is met.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `implement-error-handling-for-attempts-to-insert-a-form-field-with-an-empty-name-logging-a.cs`
  - Task: Implement error handling for attempts to insert a form field with an empty name, logging a warning.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: llm
- `build-a-console-application-that-reads-a-csv-file-and-populates-corresponding-text-input-f.cs`
  - Task: Build a console application that reads a CSV file and populates corresponding text input fields in a template.
  - Workflow: text-form-field
  - Outputs: csv
  - Selected engine: llm

## Common failure patterns and preferred agent fixes

- **Incorrect overload usage for Insert* methods**
  - Symptom: Compilation fails with CS1503, CS1501, CS7036, or CS1509 errors because arguments of the wrong type or count are supplied.
  - Root cause: The code calls InsertTextInput, InsertCheckBox, or InsertComboBox with parameters that do not match the documented signatures, such as passing a bool where TextFormFieldType is required or omitting the maxLength argument.
  - Preferred fix: Use the exact overloads: InsertTextInput(string name, TextFormFieldType type, string format, string defaultText, int maxLength); InsertCheckBox(string name, bool isChecked, int size); InsertComboBox(string name, string[] items, int selectedIndex).
  - Prevention: Consult the Aspose.Words API reference for each method and match parameter types and order precisely.

- **Missing required namespaces or referencing non-existent types**
  - Symptom: Compilation produces CS0103, CS0246, or CS0234 errors indicating that identifiers like TextFormFieldType, FormFieldCollection, or Aspose.Words.Building are unknown.
  - Root cause: The source file lacks using directives for Aspose.Words.Fields and other necessary namespaces, or it references types that are not part of the installed Aspose.Words version.
  - Preferred fix: Add using Aspose.Words.Fields; and ensure the project references Aspose.Words 26.11.0, which defines TextFormFieldType, FormField, and related classes.
  - Prevention: Always include the required namespaces listed in the category metadata before writing code.

- **Mixing legacy form fields with content controls**
  - Symptom: Examples combine StructuredDocumentTag with InsertTextInput or InsertCheckBox, leading to logical errors or unsupported API calls.
  - Root cause: A misunderstanding of the distinction between legacy form fields and modern content controls causes inappropriate API usage.
  - Preferred fix: Restrict the solution to legacy form-field APIs (DocumentBuilder.Insert* and Range.FormFields) unless the task explicitly demands content controls.
  - Prevention: Follow the canonical pattern that separates legacy field handling from StructuredDocumentTag usage.

- **Insufficient validation of field existence and value changes**
  - Symptom: Code saves the document without confirming that a form field was created or that its Result/Checked property was updated, potentially leaving the document in an incorrect state.
  - Root cause: Missing null checks and post-assignment assertions after accessing FormFields.
  - Preferred fix: After retrieving a field, verify it is not null, then compare the property value after assignment; throw an exception or log an error if the check fails.
  - Prevention: Incorporate explicit existence checks and value verification steps as part of every example.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
```

```powershell
Copy-Item ..\form-fields\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Keep the example code aligned with the Aspose.Words 26.11.0 API surface; update signatures and namespace imports if a newer version is adopted.
