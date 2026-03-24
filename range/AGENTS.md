---
name: range
description: C# examples for range using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - range

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **range** category.
This folder contains standalone C# examples for range operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **range**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (29/29 files) ← category-specific
- `using System;` (26/29 files)
- `using System.Collections.Generic;` (5/29 files)
- `using Aspose.Words.Fields;` (4/29 files)
- `using System.IO;` (4/29 files)
- `using Aspose.Words.Replacing;` (4/29 files)
- `using Aspose.Words.Saving;` (3/29 files)
- `using System.Text;` (2/29 files)
- `using Aspose.Words.Markup;` (1/29 files)
- `using System.Text.RegularExpressions;` (1/29 files)

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
| [add-new-bookmark-at-start-document-range-bookmarks](./add-new-bookmark-at-start-document-range-bookmarks.cs) | `Document`, `DocumentBuilder`, `Range` | Add new bookmark at start document range bookmarks |
| [append-text-end-range](./append-text-end-range.cs) | `DocumentBuilder`, `FirstSection`, `Body` | Append text end range |
| [checkbox-form-field-inside-specific-range-set-its-defau...](./checkbox-form-field-inside-specific-range-set-its-default-state.cs) | `Document`, `DocumentBuilder`, `CheckboxFormField` | Checkbox form field inside specific range set its default state |
| [clear-text-specific-bookmark-s-range-without-deleting-b...](./clear-text-specific-bookmark-s-range-without-deleting-bookmark.cs) | `Document`, `Input`, `Range` | Clear text specific bookmark s range without deleting bookmark |
| [copy-paragraph-s-range-content-string-variable-further-...](./copy-paragraph-s-range-content-string-variable-further-processing.cs) | `Document`, `Range`, `Input` | Copy paragraph s range content string variable further processing |
| [delete-all-characters-document-s-body-calling-delete-do...](./delete-all-characters-document-s-body-calling-delete-doc-range.cs) | `Document`, `Range` | Delete all characters document s body calling delete doc range |
| [delete-all-form-fields-within-document-iterating-over-r...](./delete-all-form-fields-within-document-iterating-over-range-formfields-calling-remove.cs) | `Document`, `Input`, `Range` | Delete all form fields within document iterating over range formfields callin... |
| [document-after-removing-all-content-its-range-empty-tem...](./document-after-removing-all-content-its-range-empty-template.cs) | `Document`, `Range`, `EmptyTemplate` | Document after removing all content its range empty template |
| [export-extracted-plain-text-range-txt-file-while-preser...](./export-extracted-plain-text-range-txt-file-while-preserving-line-breaks.cs) | `Document`, `TxtSaveOptions`, `InputDocument` | Export extracted plain text range txt file while preserving line breaks |
| [extract-plain-text-each-section-via-each-section-s-rang...](./extract-plain-text-each-section-via-each-section-s-range-text-property.cs) | `Document`, `Range`, `Sections` | Extract plain text each section via each section s range text property |
| [extract-plain-unformatted-text-document-range-text-prop...](./extract-plain-unformatted-text-document-range-text-property.cs) | `Document`, `Range`, `Text` | Extract plain unformatted text document range text property |
| [implement-batch-process-that-clears-content-multiple-do...](./implement-batch-process-that-clears-content-multiple-documents-doc-range-delete.cs) | `Document`, `DocumentBuilder`, `BatchClearer` | Implement batch process that clears content multiple documents doc range delete |
| [implement-script-that-removes-all-bookmarks-form-fields...](./implement-script-that-removes-all-bookmarks-form-fields-document-range-before.cs) | `Document`, `Range`, `Input` | Implement script that removes all bookmarks form fields document range before |
| [insert-new-text-at-beginning-range](./insert-new-text-at-beginning-range.cs) | `DocumentBuilder`, `Document`, `Run` | Insert new text at beginning range |
| [iterate-over-bookmarks-range-modify-their-names](./iterate-over-bookmarks-range-modify-their-names.cs) | `Document`, `Input`, `Range` | Iterate over bookmarks range modify their names |
| [iterate-over-form-fields-range-list-their-names-types](./iterate-over-form-fields-range-list-their-names-types.cs) | `Document`, `Collections` | Iterate over form fields range list their names types |
| [iterate-through-each-bookmark-range-output-its-name](./iterate-through-each-bookmark-range-output-its-name.cs) | `Document`, `Collections`, `Range` | Iterate through each bookmark range output its name |
| [log-names-all-bookmarks-found-range-debugging](./log-names-all-bookmarks-found-range-debugging.cs) | `Document`, `Collections` | Log names all bookmarks found range debugging |
| [perform-case-insensitive-search-within-range-collect-ma...](./perform-case-insensitive-search-within-range-collect-matching-paragraph-indices.cs) | `Document`, `CollectParagraphIndicesCallback`, `FindReplaceOptions` | Perform case insensitive search within range collect matching paragraph indices |
| [plain-text-version-document-extracting-each-section-s-r...](./plain-text-version-document-extracting-each-section-s-range-text.cs) | `Document`, `DocumentBuilder`, `StringBuilder` | Plain text version document extracting each section s range text |
| [range-object-extract-plain-text-header-footer-indexing](./range-object-extract-plain-text-header-footer-indexing.cs) | `Document`, `FirstSection`, `HeaderFooterType` | Range object extract plain text header footer indexing |
| [remove-specific-bookmark-locating-its-range-calling-remove](./remove-specific-bookmark-locating-its-range-calling-remove.cs) | `Document`, `Input`, `Range` | Remove specific bookmark locating its range calling remove |
| [replace-text-within-range-assigning-new-string-range-te...](./replace-text-within-range-assigning-new-string-range-text-property.cs) | `Document`, `DocumentBuilder`, `Range` | Replace text within range assigning new string range text property |
| [retrieve-count-bookmarks-within-range](./retrieve-count-bookmarks-within-range.cs) | `Document`, `Bookmarks`, `Sample` | Retrieve count bookmarks within range |
| [retrieve-count-form-fields-within-range](./retrieve-count-form-fields-within-range.cs) | `Document`, `FirstSection`, `FormFields` | Retrieve count form fields within range |
| [search-specific-phrase-within-range-replace-it-another-...](./search-specific-phrase-within-range-replace-it-another-string.cs) | `Document`, `DocumentBuilder`, `FirstSection` | Search specific phrase within range replace it another string |
| [summary-report-bookmark-names-their-corresponding-text-...](./summary-report-bookmark-names-their-corresponding-text-extracted-ranges.cs) | `Document`, `BookmarkReport`, `DocumentBuilder` | Summary report bookmark names their corresponding text extracted ranges |
| [update-value-text-input-form-field-located-given-range](./update-value-text-input-form-field-located-given-range.cs) | `Document`, `Range` | Update value text input form field located given range |
| [validate-that-range-contains-no-form-fields-before-perf...](./validate-that-range-contains-no-form-fields-before-performing-text-replacement.cs) | `Document`, `Range`, `FormFields` | Validate that range contains no form fields before performing text replacement |

## Category Statistics
- Total examples: 29

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for range patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
