# Lists Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Lists category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **Lists**
- Slug: **lists**
- Total examples: **30**
- Publish-ready successful examples: **30 / 30**
- Numbered list examples: **2**
- Bulleted list examples: **1**
- Nested list examples: **17**
- Numbering control examples: **1**
- List formatting examples: **5**
- General list examples: **4**

## Category rules that shaped these examples

- Create lists with doc.Lists.Add(ListTemplate).
- Apply the list through builder.ListFormat.List before writing items.
- Use builder.ListFormat.ListLevelNumber for nested levels.
- Call builder.ListFormat.RemoveNumbers() after finishing a list.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
Copy-Item ..\lists\<example-file>.cs .\Program.cs
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```
## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0
```

### PowerShell example

```powershell
Copy-Item ..\lists\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `add-a-new-list-to-the-document-s-listcollection-and-assign-it-to-multiple-paragraphs-to-sh.cs` | add a new list to the document s listcollection and assign it to multiple paragraphs to sh | general-list-workflow | docx | verified |
| 2 | `adjust-listlevel-indentation-by-setting-listlevel-indentation-to-36-points-for-proper-alig.cs` | adjust listlevel indentation by setting listlevel indentation to 36 points for proper alig | nested-list-workflow | docx | verified |
| 3 | `apply-a-custom-bullet-character-to-a-list-by-setting-listlevel-numberstyle-to-bullet-and-d.cs` | apply a custom bullet character to a list by setting listlevel numberstyle to bullet and d | nested-list-workflow | docx | verified |
| 4 | `apply-a-custom-tab-stop-to-list-levels-to-align-text-after-numbers-using-listlevel-tabposi.cs` | apply a custom tab stop to list levels to align text after numbers using listlevel tabposi | nested-list-workflow | docx | verified |
| 5 | `apply-a-default-bulleted-list-to-paragraphs-with-documentbuilder-listformat-applybulletdef.cs` | apply a default bulleted list to paragraphs with documentbuilder listformat applybulletdef | bulleted-list-workflow | docx | verified |
| 6 | `apply-a-uniform-list-style-to-all-lists-by-iterating-over-document-lists-and-updating-each.cs` | apply a uniform list style to all lists by iterating over document lists and updating each | list-formatting-workflow | docx | verified |
| 7 | `assign-an-existing-list-to-a-paragraph-by-setting-paragraph-listformat-list-property-to-ap.cs` | assign an existing list to a paragraph by setting paragraph listformat list property to ap | list-formatting-workflow | docx | verified |
| 8 | `check-document-ooxmlcompliance-property-and-set-to-higher-than-ecma376-when-modifying-list.cs` | check document ooxmlcompliance property and set to higher than ecma376 when modifying list | general-list-workflow | docx | verified |
| 9 | `clone-an-existing-list-modify-its-level-start-values-and-apply-the-cloned-list-to-new-sect.cs` | clone an existing list modify its level start values and apply the cloned list to new sect | nested-list-workflow | docx | verified |
| 10 | `configure-tabposition-of-a-listlevel-to-72-points-to-align-text-after-list-number.cs` | configure tabposition of a listlevel to 72 points to align text after list number | nested-list-workflow | docx | verified |
| 11 | `create-a-custom-list-object-configure-its-levels-and-add-it-to-document-lists-collection.cs` | create a custom list object configure its levels and add it to document lists collection | nested-list-workflow | doc | verified |
| 12 | `create-a-default-numbered-list-in-a-word-document-using-documentbuilder-listformat-applynu.cs` | create a default numbered list in a word document using documentbuilder listformat applynu | numbered-list-workflow | docx | verified |
| 13 | `create-a-multi-level-list-with-alternating-bullet-and-number-styles-by-configuring-each-li.cs` | create a multi level list with alternating bullet and number styles by configuring each li | nested-list-workflow | docx | verified |
| 14 | `create-a-nine-level-hierarchical-list-by-defining-properties-for-each-listlevel-up-to-leve.cs` | create a nine level hierarchical list by defining properties for each listlevel up to leve | nested-list-workflow | docx | verified |
| 15 | `create-a-numbered-list-that-restarts-numbering-at-each-new-chapter-by-resetting-startingnu.cs` | create a numbered list that restarts numbering at each new chapter by resetting startingnu | numbering-control-workflow | docx | verified |
| 16 | `decrease-list-indent-with-documentbuilder-listformat-decreaseindent-to-promote-paragraph-t.cs` | decrease list indent with documentbuilder listformat decreaseindent to promote paragraph t | list-formatting-workflow | docx | verified |
| 17 | `define-startingnumber-for-a-listlevel-to-begin-numbering-at-five-instead-of-default-one.cs` | define startingnumber for a listlevel to begin numbering at five instead of default one | nested-list-workflow | docx | verified |
| 18 | `ensure-ooxmlcompliance-is-set-higher-than-ecma376-before-saving-document-to-retain-custom.cs` | ensure ooxmlcompliance is set higher than ecma376 before saving document to retain custom | general-list-workflow | docx | verified |
| 19 | `implement-error-handling-for-attempts-to-create-more-than-nine-list-levels-catching-the-re.cs` | implement error handling for attempts to create more than nine list levels catching the re | nested-list-workflow | docx | verified |
| 20 | `increase-list-indent-programmatically-using-documentbuilder-listformat-increaseindent-to-m.cs` | increase list indent programmatically using documentbuilder listformat increaseindent to m | list-formatting-workflow | docx | verified |
| 21 | `iterate-through-document-lists-collection-to-modify-all-list-definitions-uniformly-across.cs` | iterate through document lists collection to modify all list definitions uniformly across | general-list-workflow | docx | verified |
| 22 | `programmatically-decrease-list-level-depth-for-a-paragraph-using-documentbuilder-listforma.cs` | programmatically decrease list level depth for a paragraph using documentbuilder listforma | nested-list-workflow | docx | verified |
| 23 | `programmatically-increase-list-level-depth-for-a-paragraph-using-documentbuilder-listforma.cs` | programmatically increase list level depth for a paragraph using documentbuilder listforma | nested-list-workflow | docx | verified |
| 24 | `programmatically-remove-list-formatting-from-selected-paragraphs-while-preserving-their-in.cs` | programmatically remove list formatting from selected paragraphs while preserving their in | list-formatting-workflow | docx | verified |
| 25 | `remove-numbering-from-a-paragraph-by-calling-documentbuilder-listformat-removenumbers-meth.cs` | remove numbering from a paragraph by calling documentbuilder listformat removenumbers meth | numbered-list-workflow | docx | verified |
| 26 | `restart-numbering-for-each-new-section-by-resetting-listlevel-startingnumber-before-applyi.cs` | restart numbering for each new section by resetting listlevel startingnumber before applyi | nested-list-workflow | docx | verified |
| 27 | `retrieve-a-specific-list-from-document-lists-by-its-id-and-adjust-its-level-properties-pro.cs` | retrieve a specific list from document lists by its id and adjust its level properties pro | nested-list-workflow | docx | verified |
| 28 | `set-listlevelnumber-to-three-before-adding-a-paragraph-to-create-a-third-level-list-item-i.cs` | set listlevelnumber to three before adding a paragraph to create a third level list item i | nested-list-workflow | docx | verified |
| 29 | `set-numberstyle-of-a-listlevel-to-upperroman-to-generate-roman-numeral-list-items.cs` | set numberstyle of a listlevel to upperroman to generate roman numeral list items | nested-list-workflow | docx | verified |
| 30 | `validate-that-each-list-in-a-document-does-not-exceed-nine-levels-to-comply-with-api-const.cs` | validate that each list in a document does not exceed nine levels to comply with api const | nested-list-workflow | docx | verified |

## Notes for maintainers

- This category is now **100% publish-ready** for the current run.
