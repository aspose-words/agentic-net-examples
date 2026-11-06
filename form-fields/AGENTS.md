---
name: form-fields
description: Verified C# examples for form-fields scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Form Fields

## Purpose

Provide concise, reliable guidance for an AI coding agent to generate correct C# snippets that create, read, update, and protect Aspose.Words legacy form fields while avoiding common compilation and runtime pitfalls.

## Non-negotiable conventions

- Never assume that a form field already exists in the target document; always insert or verify it explicitly.
- Do not use StructuredDocumentTag or other content-control APIs when the task requires legacy form fields.
- Avoid inventing helper methods that are not part of the Aspose.Words API.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Text Form Field workflow**: 7 examples
- **General Form Field workflow**: 17 examples
- **Dropdown Form Field workflow**: 4 examples
- **Inspection And Reporting workflow**: 2 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Always verify that the field name exists in Document.Range.FormFields and that its type matches the intended operation before reading or writing any property.

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
  - Selected engine: llm
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
  - Selected engine: llm
- `batch-insert-multiple-text-input-form-fields-into-a-template-using-a-loop-over-field-defin.cs`
  - Task: Batch insert multiple text input form fields into a template using a loop over field definitions.
  - Workflow: text-form-field
  - Outputs: docx
  - Selected engine: mcp
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
  - Selected engine: mcp
- `retrieve-a-form-field-by-its-name-from-a-loaded-document-and-read-its-result-property.cs`
  - Task: Retrieve a form field by its name from a loaded document and read its Result property.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: llm
- `access-a-form-field-by-index-from-the-formfields-collection-and-modify-its-result-value.cs`
  - Task: Access a form field by index from the FormFields collection and modify its Result value.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `use-formfield-type-enumeration-to-differentiate-between-text-input-check-box-and-combo-box.cs`
  - Task: Use FormField.Type enumeration to differentiate between text input, check box, and combo box fields.
  - Workflow: dropdown-form-field
  - Outputs: docx
  - Selected engine: mcp
- `iterate-over-all-form-fields-in-a-document-and-list-each-field-s-name-and-type.cs`
  - Task: Iterate over all form fields in a document and list each field's name and type.
  - Workflow: inspection-and-reporting
  - Outputs: docx
  - Selected engine: existing_repo
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
  - Selected engine: llm
- `navigate-to-a-form-field-using-its-automatically-created-bookmark-and-extract-surrounding.cs`
  - Task: Navigate to a form field using its automatically created bookmark and extract surrounding paragraph text.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: existing_repo
- `protect-a-word-document-with-allowonlyformfields-option-allowing-only-form-fields-to-be-ed.cs`
  - Task: Protect a Word document with AllowOnlyFormFields option, allowing only form fields to be edited.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: mcp
- `apply-document-protection-levels-ensuring-only-form-fields-remain-editable-while-other-sec.cs`
  - Task: Apply document protection levels ensuring only form fields remain editable while other sections are read-only.
  - Workflow: general-form-field
  - Outputs: docx
  - Selected engine: existing_repo
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
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Incorrect API signatures**
  - Symptom: Compilation errors such as CS1503 and CS7036 indicate mismatched argument types or missing required parameters.
  - Root cause: The code supplies arguments in the wrong order, uses wrong data types, or omits required parameters like maxLength for InsertTextInput or selectedIndex for InsertComboBox.
  - Preferred fix: Consult the Aspose.Words documentation and pass arguments exactly as defined: name (string), TextFormFieldType, format (string), default text (string), maxLength (int) for InsertTextInput; name, isChecked (bool), size (int) for InsertCheckBox; name, items (string[]), selectedIndex (int) for InsertComboBox.
  - Prevention: Write a small wrapper that calls the overload with named parameters and compile-time checks to ensure all required arguments are provided.

- **Missing or incorrect namespaces**
  - Symptom: CS0246 and CS0234 errors show that types like FormField, TextFormFieldType, and ProtectionType cannot be resolved.
  - Root cause: The source file lacks the required using directives or references the wrong Aspose.Words sub-namespaces (e.g., Aspose.Words.Building, Aspose.Words.Forms).
  - Preferred fix: Add using Aspose.Words; using Aspose.Words.Fields; and ensure the project references Aspose.Words.dll version 26.10.0.
  - Prevention: Include a standard header block with the necessary using statements in every form-field example template.

- **Mixing legacy fields with content controls**
  - Symptom: Examples that import StructuredDocumentTag or other content-control classes fail to compile or behave unexpectedly.
  - Root cause: Legacy form fields and content controls are distinct APIs; the task explicitly requires legacy fields.
  - Preferred fix: Remove all StructuredDocumentTag usage and rely solely on DocumentBuilder legacy insertion methods and the FormFields collection.
  - Prevention: Document the distinction in the coding guidelines and enforce a lint rule that flags any usage of StructuredDocumentTag in form-field samples.

- **Assuming pre-existing fields**
  - Symptom: Runtime code attempts to read or modify a field that was never inserted, leading to null-reference exceptions.
  - Root cause: The sample code starts by loading a document and directly accesses FormFields without first creating them.
  - Preferred fix: Bootstrap a minimal document that inserts the required fields before any read or update logic.
  - Prevention: Include a step in the workflow that always creates a fresh document or explicitly checks for field existence and throws a clear error if missing.

- **Invalid protection constants**
  - Symptom: CS0117 errors indicate that ProtectionType.Forms does not exist.
  - Root cause: The correct enum value is ProtectionType.AllowOnlyFormFields; using an outdated or misspelled constant causes compilation failure.
  - Preferred fix: Replace ProtectionType.Forms with ProtectionType.AllowOnlyFormFields when calling Document.Protect.
  - Prevention: Reference the official enum list in the documentation and add a unit test that verifies document protection can be applied without errors.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.10.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
```

```powershell
Copy-Item ..\form-fields\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Update the example repository whenever Aspose.Words introduces new overloads or deprecates legacy form-field APIs to keep the samples current.
