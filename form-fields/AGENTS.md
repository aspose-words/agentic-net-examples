---
name: form-fields
description: C# examples for form-fields using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - form-fields

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **form-fields** category.
This folder contains standalone C# examples for form-fields operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **form-fields**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using System;` (30/30 files) ← category-specific
- `using Aspose.Words;` (30/30 files)
- `using Aspose.Words.Fields;` (30/30 files)
- `using System.Collections.Generic;` (6/30 files)
- `using System.IO;` (3/30 files)
- `using Aspose.Words.Saving;` (1/30 files)
- `using Aspose.Words.Reporting;` (1/30 files)
- `using System.Text.Json;` (1/30 files)
- `using System.Diagnostics;` (1/30 files)

## Common Code Pattern

Most files follow this pattern:

```csharp
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);
// ... operations ...
doc.Save("output.docx");
```

## Files in this folder

| File | Key APIs | Description |
|------|----------|-------------|
| [access-form-field-index-formfields-collection-modify-it...](./access-form-field-index-formfields-collection-modify-its-result-value.cs) | `Document`, `Aspose`, `Input` | Access form field index formfields collection modify its result value |
| [add-check-box-form-field-set-default-checked-state-cust...](./add-check-box-form-field-set-default-checked-state-custom-size-via-documentbuilder.cs) | `Document`, `DocumentBuilder`, `Aspose` | Add check box form field set default checked state custom size via documentbu... |
| [after-inserting-fields-protect-document-then-verify-tha...](./after-inserting-fields-protect-document-then-verify-that-non-field-content-cannot-be.cs) | `Document`, `DocumentBuilder`, `Aspose` | After inserting fields protect document then verify that non field content ca... |
| [apply-document-protection-levels-ensuring-only-form-fie...](./apply-document-protection-levels-ensuring-only-form-fields-remain-editable-while.cs) | `Document`, `DocumentBuilder`, `Aspose` | Apply document protection levels ensuring only form fields remain editable while |
| [assign-unique-name-each-inserted-form-field-ensure-dist...](./assign-unique-name-each-inserted-form-field-ensure-distinct-automatic-bookmarks.cs) | `Document`, `DocumentBuilder`, `Aspose` | Assign unique name each inserted form field ensure distinct automatic bookmarks |
| [batch-insert-multiple-text-input-form-fields-template-l...](./batch-insert-multiple-text-input-form-fields-template-loop-over-field-definitions.cs) | `TextFormFieldType`, `Aspose`, `Document` | Batch insert multiple text input form fields template loop over field definit... |
| [build-console-application-that-reads-csv-file-populates...](./build-console-application-that-reads-csv-file-populates-corresponding-text-input.cs) | `Aspose`, `Document`, `CsvDataLoadOptions` | Build console application that reads csv file populates corresponding text input |
| [change-size-existing-check-box-form-field-programmatica...](./change-size-existing-check-box-form-field-programmatically-improve-visual-consistency.cs) | `Document`, `Aspose`, `Words` | Change size existing check box form field programmatically improve visual con... |
| [combo-box-form-field-containing-three-items-set-default...](./combo-box-form-field-containing-three-items-set-default-selected-index-documentbuilder.cs) | `Document`, `DocumentBuilder`, `Aspose` | Combo box form field containing three items set default selected index docume... |
| [count-number-each-form-field-type-iterating-through-for...](./count-number-each-form-field-type-iterating-through-formfields-collection.cs) | `Document`, `DocumentBuilder`, `Aspose` | Count number each form field type iterating through formfields collection |
| [disable-all-check-box-form-fields-programmatically-when...](./disable-all-check-box-form-fields-programmatically-when-specific-document-condition.cs) | `Document`, `Aspose`, `Words` | Disable all check box form fields programmatically when specific document con... |
| [docx-file-change-combo-box-selections-programmatically-...](./docx-file-change-combo-box-selections-programmatically-modified-document.cs) | `Document`, `Aspose`, `Words` | Docx file change combo box selections programmatically modified document |
| [extract-automatically-bookmark-names-all-form-fields-st...](./extract-automatically-bookmark-names-all-form-fields-store-them-lookup-dictionary.cs) | `Aspose`, `Document`, `Words` | Extract automatically bookmark names all form fields store them lookup dictio... |
| [formfield-type-enumeration-differentiate-between-text-i...](./formfield-type-enumeration-differentiate-between-text-input-check-box-combo-box-fields.cs) | `BreakType`, `FieldType`, `Document` | Formfield type enumeration differentiate between text input check box combo b... |
| [implement-error-handling-attempts-insert-form-field-emp...](./implement-error-handling-attempts-insert-form-field-empty-name-logging-warning.cs) | `Aspose`, `Document`, `DocumentBuilder` | Implement error handling attempts insert form field empty name logging warning |
| [implement-reusable-method-that-adds-combo-box-form-fiel...](./implement-reusable-method-that-adds-combo-box-form-field-customizable-items-default.cs) | `Aspose`, `Document`, `DocumentBuilder` | Implement reusable method that adds combo box form field customizable items d... |
| [insert-text-input-field-custom-date-format-default-curr...](./insert-text-input-field-custom-date-format-default-current-date-value-documentbuilder.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert text input field custom date format default current date value documen... |
| [insert-text-input-field-maximum-length-50-characters-nu...](./insert-text-input-field-maximum-length-50-characters-numeric-format-documentbuilder.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert text input field maximum length 50 characters numeric format documentb... |
| [insert-text-input-form-field-placeholder-text-documentb...](./insert-text-input-form-field-placeholder-text-documentbuilder-new-document.cs) | `Document`, `DocumentBuilder`, `Aspose` | Insert text input form field placeholder text documentbuilder new document |
| [iterate-over-all-form-fields-document-list-each-field-s...](./iterate-over-all-form-fields-document-list-each-field-s-name-type.cs) | `Document`, `Aspose`, `System` | Iterate over all form fields document list each field s name type |
| [iterate-through-form-fields-log-each-field-s-result-val...](./iterate-through-form-fields-log-each-field-s-result-value-debugging-purposes.cs) | `Document`, `Aspose`, `Words` | Iterate through form fields log each field s result value debugging purposes |
| [navigate-form-field-its-automatically-created-bookmark-...](./navigate-form-field-its-automatically-created-bookmark-extract-surrounding-paragraph.cs) | `Document`, `DocumentBuilder`, `Aspose` | Navigate form field its automatically created bookmark extract surrounding pa... |
| [protect-word-document-allowonlyformfields-option-allowi...](./protect-word-document-allowonlyformfields-option-allowing-only-form-fields-be-edited.cs) | `Document`, `DocumentBuilder`, `Aspose` | Protect word document allowonlyformfields option allowing only form fields be... |
| [read-result-property-check-box-form-field-determine-whe...](./read-result-property-check-box-form-field-determine-whether-it-is-checked.cs) | `Document`, `Aspose`, `Words` | Read result property check box form field determine whether it is checked |
| [retrieve-form-field-its-name-loaded-document-read-its-r...](./retrieve-form-field-its-name-loaded-document-read-its-result-property.cs) | `Document`, `Aspose`, `Words` | Retrieve form field its name loaded document read its result property |
| [set-result-property-text-input-form-field-predefined-st...](./set-result-property-text-input-form-field-predefined-string-value.cs) | `Document`, `DocumentBuilder`, `Aspose` | Set result property text input form field predefined string value |
| [specify-name-when-inserting-text-input-field-automatica...](./specify-name-when-inserting-text-input-field-automatically-matching-bookmark.cs) | `Aspose`, `Document`, `DocumentBuilder` | Specify name when inserting text input field automatically matching bookmark |
| [toggle-checked-state-specific-check-box-form-field-base...](./toggle-checked-state-specific-check-box-form-field-based-external-configuration.cs) | `Document`, `Aspose`, `Words` | Toggle checked state specific check box form field based external configuration |
| [update-result-values-several-check-box-form-fields-base...](./update-result-values-several-check-box-form-fields-based-external-json-configuration.cs) | `System`, `Document`, `Aspose` | Update result values several check box form fields based external json config... |
| [verify-that-protected-document-still-permits-editing-fo...](./verify-that-protected-document-still-permits-editing-form-fields-after-reopening.cs) | `Document`, `Debug`, `DocumentBuilder` | Verify that protected document still permits editing form fields after reopening |

## Category Statistics
- Total examples: 30

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for form-fields patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082405`
<!-- AUTOGENERATED:END -->
