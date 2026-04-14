---
name: lists
description: Verified C# examples for list scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - Lists

## Purpose

This folder is a live, curated example set for list scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free creation and formatting of numbered, bulleted, and nested lists using direct Aspose.Words APIs.

## Non-negotiable conventions

- Create lists with doc.Lists.Add(ListTemplate).
- Apply a list through builder.ListFormat.List before writing list items.
- Use builder.ListFormat.ListLevelNumber for nested list levels.
- Call builder.ListFormat.RemoveNumbers() after finishing a list.
- Do not invent unsupported list APIs.

## Recommended workflow selection

- **Numbered list workflow**: 2 examples
- **Bulleted list workflow**: 1 examples
- **Nested list workflow**: 17 examples
- **Numbering control workflow**: 1 examples
- **List formatting workflow**: 5 examples
- **General list workflow**: 4 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. List formatting must be applied to the intended paragraphs only.
3. Nested levels must be set explicitly when required.
4. List formatting must be removed after the list is complete.
5. Requested output files must actually be written.

## File-to-task reference

- `add-a-new-list-to-the-document-s-listcollection-and-assign-it-to-multiple-paragraphs-to-sh.cs`
  - Task: add a new list to the document s listcollection and assign it to multiple paragraphs to sh
  - Workflow: general-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `adjust-listlevel-indentation-by-setting-listlevel-indentation-to-36-points-for-proper-alig.cs`
  - Task: adjust listlevel indentation by setting listlevel indentation to 36 points for proper alig
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `apply-a-custom-bullet-character-to-a-list-by-setting-listlevel-numberstyle-to-bullet-and-d.cs`
  - Task: apply a custom bullet character to a list by setting listlevel numberstyle to bullet and d
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `apply-a-custom-tab-stop-to-list-levels-to-align-text-after-numbers-using-listlevel-tabposi.cs`
  - Task: apply a custom tab stop to list levels to align text after numbers using listlevel tabposi
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `apply-a-default-bulleted-list-to-paragraphs-with-documentbuilder-listformat-applybulletdef.cs`
  - Task: apply a default bulleted list to paragraphs with documentbuilder listformat applybulletdef
  - Workflow: bulleted-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `apply-a-uniform-list-style-to-all-lists-by-iterating-over-document-lists-and-updating-each.cs`
  - Task: apply a uniform list style to all lists by iterating over document lists and updating each
  - Workflow: list-formatting-workflow
  - Outputs: docx
  - Selected engine: verified
- `assign-an-existing-list-to-a-paragraph-by-setting-paragraph-listformat-list-property-to-ap.cs`
  - Task: assign an existing list to a paragraph by setting paragraph listformat list property to ap
  - Workflow: list-formatting-workflow
  - Outputs: docx
  - Selected engine: verified
- `check-document-ooxmlcompliance-property-and-set-to-higher-than-ecma376-when-modifying-list.cs`
  - Task: check document ooxmlcompliance property and set to higher than ecma376 when modifying list
  - Workflow: general-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `clone-an-existing-list-modify-its-level-start-values-and-apply-the-cloned-list-to-new-sect.cs`
  - Task: clone an existing list modify its level start values and apply the cloned list to new sect
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `configure-tabposition-of-a-listlevel-to-72-points-to-align-text-after-list-number.cs`
  - Task: configure tabposition of a listlevel to 72 points to align text after list number
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `create-a-custom-list-object-configure-its-levels-and-add-it-to-document-lists-collection.cs`
  - Task: create a custom list object configure its levels and add it to document lists collection
  - Workflow: nested-list-workflow
  - Outputs: doc
  - Selected engine: verified
- `create-a-default-numbered-list-in-a-word-document-using-documentbuilder-listformat-applynu.cs`
  - Task: create a default numbered list in a word document using documentbuilder listformat applynu
  - Workflow: numbered-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `create-a-multi-level-list-with-alternating-bullet-and-number-styles-by-configuring-each-li.cs`
  - Task: create a multi level list with alternating bullet and number styles by configuring each li
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `create-a-nine-level-hierarchical-list-by-defining-properties-for-each-listlevel-up-to-leve.cs`
  - Task: create a nine level hierarchical list by defining properties for each listlevel up to leve
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `create-a-numbered-list-that-restarts-numbering-at-each-new-chapter-by-resetting-startingnu.cs`
  - Task: create a numbered list that restarts numbering at each new chapter by resetting startingnu
  - Workflow: numbering-control-workflow
  - Outputs: docx
  - Selected engine: verified
- `decrease-list-indent-with-documentbuilder-listformat-decreaseindent-to-promote-paragraph-t.cs`
  - Task: decrease list indent with documentbuilder listformat decreaseindent to promote paragraph t
  - Workflow: list-formatting-workflow
  - Outputs: docx
  - Selected engine: verified
- `define-startingnumber-for-a-listlevel-to-begin-numbering-at-five-instead-of-default-one.cs`
  - Task: define startingnumber for a listlevel to begin numbering at five instead of default one
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `ensure-ooxmlcompliance-is-set-higher-than-ecma376-before-saving-document-to-retain-custom.cs`
  - Task: ensure ooxmlcompliance is set higher than ecma376 before saving document to retain custom
  - Workflow: general-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `implement-error-handling-for-attempts-to-create-more-than-nine-list-levels-catching-the-re.cs`
  - Task: implement error handling for attempts to create more than nine list levels catching the re
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `increase-list-indent-programmatically-using-documentbuilder-listformat-increaseindent-to-m.cs`
  - Task: increase list indent programmatically using documentbuilder listformat increaseindent to m
  - Workflow: list-formatting-workflow
  - Outputs: docx
  - Selected engine: verified
- `iterate-through-document-lists-collection-to-modify-all-list-definitions-uniformly-across.cs`
  - Task: iterate through document lists collection to modify all list definitions uniformly across
  - Workflow: general-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `programmatically-decrease-list-level-depth-for-a-paragraph-using-documentbuilder-listforma.cs`
  - Task: programmatically decrease list level depth for a paragraph using documentbuilder listforma
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `programmatically-increase-list-level-depth-for-a-paragraph-using-documentbuilder-listforma.cs`
  - Task: programmatically increase list level depth for a paragraph using documentbuilder listforma
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `programmatically-remove-list-formatting-from-selected-paragraphs-while-preserving-their-in.cs`
  - Task: programmatically remove list formatting from selected paragraphs while preserving their in
  - Workflow: list-formatting-workflow
  - Outputs: docx
  - Selected engine: verified
- `remove-numbering-from-a-paragraph-by-calling-documentbuilder-listformat-removenumbers-meth.cs`
  - Task: remove numbering from a paragraph by calling documentbuilder listformat removenumbers meth
  - Workflow: numbered-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `restart-numbering-for-each-new-section-by-resetting-listlevel-startingnumber-before-applyi.cs`
  - Task: restart numbering for each new section by resetting listlevel startingnumber before applyi
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `retrieve-a-specific-list-from-document-lists-by-its-id-and-adjust-its-level-properties-pro.cs`
  - Task: retrieve a specific list from document lists by its id and adjust its level properties pro
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `set-listlevelnumber-to-three-before-adding-a-paragraph-to-create-a-third-level-list-item-i.cs`
  - Task: set listlevelnumber to three before adding a paragraph to create a third level list item i
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `set-numberstyle-of-a-listlevel-to-upperroman-to-generate-roman-numeral-list-items.cs`
  - Task: set numberstyle of a listlevel to upperroman to generate roman numeral list items
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified
- `validate-that-each-list-in-a-document-does-not-exceed-nine-levels-to-comply-with-api-const.cs`
  - Task: validate that each list in a document does not exceed nine levels to comply with api const
  - Workflow: nested-list-workflow
  - Outputs: docx
  - Selected engine: verified

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
Copy-Item ..\lists\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```
