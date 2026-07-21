---
name: lists
description: Verified C# examples for lists scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Lists

## Purpose

This folder is a live, curated example set for Lists scenarios. Treat every `.cs` file as a standalone console application. The goal is correct, warning-free examples that use documented Aspose APIs and match the original task intent.

## Non-negotiable conventions

- Create lists with doc.Lists.Add(ListTemplate).
- Apply a list through builder.ListFormat.List before writing list items.
- Use builder.ListFormat.ListLevelNumber for nested list levels.
- Call builder.ListFormat.RemoveNumbers() after finishing a list.
- Do not invent unsupported list APIs.

## Recommended workflow selection

- Numbered List Workflow workflow: 2 examples
- Bulleted List Workflow workflow: 1 examples
- Nested List Workflow workflow: 17 examples
- List Formatting Workflow workflow: 5 examples
- General List Workflow workflow: 4 examples
- Numbering Control Workflow workflow: 1 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. API usage must be supported by the configured package versions.
3. Exported outputs must actually be written by the example.
4. Validation scenarios must inspect only the behavior requested by the task.
5. Examples that depend on files, folders, streams, images, or data should bootstrap those inputs locally during the example run.

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
  - Selected engine: mcp
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
  - Selected engine: mcp
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
  - Selected engine: mcp
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
  - Selected engine: mcp
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

- **Invented list APIs**
  - Symptom: Code uses unsupported helpers such as CreateList or AddListItem.
  - Preferred fix: Use doc.Lists.Add(ListTemplate) and builder.ListFormat.List only.

- **List formatting not applied**
  - Symptom: Paragraphs are written but they are not actual list items.
  - Preferred fix: Assign builder.ListFormat.List before writing each list sequence.

- **Nested level drift**
  - Symptom: Nested items appear at the wrong level or never return to the parent level.
  - Preferred fix: Set builder.ListFormat.ListLevelNumber explicitly before each nested level transition.

- **List mode not cleared**
  - Symptom: Subsequent paragraphs continue as list items unintentionally.
  - Preferred fix: Call builder.ListFormat.RemoveNumbers() immediately after the list is complete.

## Build and run contract

- Target framework: `net8.0`
- Package: `Aspose.Words` `26.6.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.6.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\lists\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer documented Aspose APIs over speculative shortcuts.
