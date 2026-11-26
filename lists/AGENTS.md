---
name: lists
description: Verified C# examples for lists scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Lists

## Purpose

Provide an AI coding assistant with clear guidance on how to generate and modify Word lists reliably, ensuring that only supported Aspose.Words APIs are used and that list formatting is applied and cleared correctly.

## Non-negotiable conventions

- Do not reference APIs that are not present in Aspose.Words, such as ListFormat.IncreaseIndent, DecreaseIndent, List.Clone, or ListNumberStyle.
- Do not omit the assignment of builder.ListFormat.List before writing list items.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Numbered List Workflow workflow**: 2 examples
- **Bulleted List Workflow workflow**: 1 examples
- **Nested List Workflow workflow**: 17 examples
- **List Formatting Workflow workflow**: 5 examples
- **General List Workflow workflow**: 4 examples
- **Numbering Control Workflow workflow**: 1 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Ensure that every list sequence sets builder.ListFormat.List, adjusts ListLevelNumber as needed, and calls RemoveNumbers() when the list ends.

## File-to-task reference

- `create-a-default-numbered-list-in-a-word-document-using-documentbuilder-listformat-applynu.cs`
  - Task: Create a default numbered list in a Word document using DocumentBuilder.ListFormat.ApplyNumberDefault method.
  - Workflow: numbered-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-default-bulleted-list-to-paragraphs-with-documentbuilder-listformat-applybulletdef.cs`
  - Task: Apply a default bulleted list to paragraphs with DocumentBuilder.ListFormat.ApplyBulletDefault method for formatting.
  - Workflow: bulleted-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `set-listlevelnumber-to-three-before-adding-a-paragraph-to-create-a-third-level-list-item-i.cs`
  - Task: Set ListLevelNumber to three before adding a paragraph to create a third-level list item in the document.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `increase-list-indent-programmatically-using-documentbuilder-listformat-increaseindent-to-m.cs`
  - Task: Increase list indent programmatically using DocumentBuilder.ListFormat.IncreaseIndent to move paragraph deeper within the current list structure.
  - Workflow: list-formatting-workflow
  - Outputs: docx
  - Selected engine: mcp
- `decrease-list-indent-with-documentbuilder-listformat-decreaseindent-to-promote-paragraph-t.cs`
  - Task: Decrease list indent with DocumentBuilder.ListFormat.DecreaseIndent to promote paragraph to higher level in the outline.
  - Workflow: list-formatting-workflow
  - Outputs: docx
  - Selected engine: mcp
- `remove-numbering-from-a-paragraph-by-calling-documentbuilder-listformat-removenumbers-meth.cs`
  - Task: Remove numbering from a paragraph by calling DocumentBuilder.ListFormat.RemoveNumbers method to stop list formatting.
  - Workflow: numbered-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-custom-list-object-configure-its-levels-and-add-it-to-document-lists-collection.cs`
  - Task: Create a custom List object, configure its levels, and add it to Document.Lists collection.
  - Workflow: nested-list-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `add-a-new-list-to-the-document-s-listcollection-and-assign-it-to-multiple-paragraphs-to-sh.cs`
  - Task: Add a new List to the document's ListCollection and assign it to multiple paragraphs to share formatting.
  - Workflow: general-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `assign-an-existing-list-to-a-paragraph-by-setting-paragraph-listformat-list-property-to-ap.cs`
  - Task: Assign an existing List to a paragraph by setting Paragraph.ListFormat.List property to apply the desired list style.
  - Workflow: list-formatting-workflow
  - Outputs: docx
  - Selected engine: mcp
- `clone-an-existing-list-modify-its-level-start-values-and-apply-the-cloned-list-to-new-sect.cs`
  - Task: Clone an existing List, modify its level start values, and apply the cloned list to new sections.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `set-numberstyle-of-a-listlevel-to-upperroman-to-generate-roman-numeral-list-items.cs`
  - Task: Set NumberStyle of a ListLevel to UpperRoman to generate Roman numeral list items.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: llm
- `apply-a-custom-bullet-character-to-a-list-by-setting-listlevel-numberstyle-to-bullet-and-d.cs`
  - Task: Apply a custom bullet character to a list by setting ListLevel.NumberStyle to Bullet and defining BulletChar.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `define-startingnumber-for-a-listlevel-to-begin-numbering-at-five-instead-of-default-one.cs`
  - Task: Define StartingNumber for a ListLevel to begin numbering at five instead of default one.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `restart-numbering-for-each-new-section-by-resetting-listlevel-startingnumber-before-applyi.cs`
  - Task: Restart numbering for each new section by resetting ListLevel.StartingNumber before applying the list.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-numbered-list-that-restarts-numbering-at-each-new-chapter-by-resetting-startingnu.cs`
  - Task: Create a numbered list that restarts numbering at each new chapter by resetting StartingNumber on chapter start.
  - Workflow: numbering-control-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-nine-level-hierarchical-list-by-defining-properties-for-each-listlevel-up-to-leve.cs`
  - Task: Create a nine-level hierarchical list by defining properties for each ListLevel up to level nine.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `create-a-multi-level-list-with-alternating-bullet-and-number-styles-by-configuring-each-li.cs`
  - Task: Create a multi-level list with alternating bullet and number styles by configuring each ListLevel individually.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: llm
- `adjust-listlevel-indentation-by-setting-listlevel-indentation-to-36-points-for-proper-alig.cs`
  - Task: Adjust ListLevel indentation by setting ListLevel.Indentation to 36 points for proper alignment.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-custom-tab-stop-to-list-levels-to-align-text-after-numbers-using-listlevel-tabposi.cs`
  - Task: Apply a custom tab stop to list levels to align text after numbers using ListLevel.TabPosition property.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `configure-tabposition-of-a-listlevel-to-72-points-to-align-text-after-list-number.cs`
  - Task: Configure TabPosition of a ListLevel to 72 points to align text after list number.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `validate-that-each-list-in-a-document-does-not-exceed-nine-levels-to-comply-with-api-const.cs`
  - Task: Validate that each list in a document does not exceed nine levels to comply with API constraints.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `implement-error-handling-for-attempts-to-create-more-than-nine-list-levels-catching-the-re.cs`
  - Task: Implement error handling for attempts to create more than nine list levels, catching the resulting exception.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `ensure-ooxmlcompliance-is-set-higher-than-ecma376-before-saving-document-to-retain-custom.cs`
  - Task: Ensure OoxmlCompliance is set higher than Ecma376 before saving document to retain custom list settings.
  - Workflow: general-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `check-document-ooxmlcompliance-property-and-set-to-higher-than-ecma376-when-modifying-list.cs`
  - Task: Check document OoxmlCompliance property and set to higher than Ecma376 when modifying list definitions.
  - Workflow: general-list-workflow
  - Outputs: docx
  - Selected engine: existing_repo
- `retrieve-a-specific-list-from-document-lists-by-its-id-and-adjust-its-level-properties-pro.cs`
  - Task: Retrieve a specific List from Document.Lists by its ID and adjust its level properties programmatically.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `iterate-through-document-lists-collection-to-modify-all-list-definitions-uniformly-across.cs`
  - Task: Iterate through Document.Lists collection to modify all list definitions uniformly across the document.
  - Workflow: general-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `apply-a-uniform-list-style-to-all-lists-by-iterating-over-document-lists-and-updating-each.cs`
  - Task: Apply a uniform list style to all lists by iterating over Document.Lists and updating each level.
  - Workflow: list-formatting-workflow
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-increase-list-level-depth-for-a-paragraph-using-documentbuilder-listforma.cs`
  - Task: Programmatically increase list level depth for a paragraph using DocumentBuilder.ListFormat.IncreaseIndent within a loop.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-decrease-list-level-depth-for-a-paragraph-using-documentbuilder-listforma.cs`
  - Task: Programmatically decrease list level depth for a paragraph using DocumentBuilder.ListFormat.DecreaseIndent inside a conditional.
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-remove-list-formatting-from-selected-paragraphs-while-preserving-their-in.cs`
  - Task: Programmatically remove list formatting from selected paragraphs while preserving their indentation using ListFormat.RemoveNumbers.
  - Workflow: list-formatting-workflow
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Use of non-existent API members**
  - Symptom: Compilation fails with CS1061 because members such as IncreaseIndent, DecreaseIndent, Indentation, StartingNumber, BulletChar, or List.Clone are not found.
  - Root cause: The code references API members that were removed or never existed in the Aspose.Words 26.11.0 library.
  - Preferred fix: Replace the missing members with supported alternatives: control nesting via builder.ListFormat.ListLevelNumber and adjust positioning with ListLevel.NumberPosition and TextPosition; use ListLevel.StartAt to set the start number.
  - Prevention: Consult the official Aspose.Words API reference and limit code to members that appear in the current version's documentation.

- **Incorrect enum or property names**
  - Symptom: Compilation fails with CS0103 or CS0234 when using ListNumberStyle, NumberStyle.Decimal, or ListLevel.StartingNumber, which are undefined.
  - Root cause: The code uses outdated or misspelled identifiers; the correct enum is NumberStyle and the correct start-number property is ListLevel.StartAt.
  - Preferred fix: Replace ListNumberStyle with NumberStyle, use NumberStyle.Arabic, UppercaseRoman, Bullet, etc., and set ListLevel.StartAt instead of StartingNumber.
  - Prevention: Reference the latest enum definitions and property names from the Aspose.Words.Lists namespace.

- **Unsupported OoxmlCompliance usage**
  - Symptom: Compilation fails with CS1061 because Document, CompatibilityOptions, or OoxmlSaveOptions do not expose an OoxmlCompliance member.
  - Root cause: The OoxmlCompliance property is not part of the Aspose.Words 26.11.0 API surface.
  - Preferred fix: Remove any attempts to set OoxmlCompliance; rely on the default compliance level or configure compatibility via Document.CompatibilityOptions if needed.
  - Prevention: Verify property availability in the target Aspose.Words version before using it in code.

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
Copy-Item ..\lists\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Periodically review the Aspose.Words release notes to update the list of supported APIs and adjust the canonical patterns accordingly.
