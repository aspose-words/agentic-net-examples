# Lists Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the Lists category. Each file is a standalone console example selected from the verified 26.5.0 run.

## Snapshot

- Category: Lists
- Slug: lists
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20260619_131835_59df5f
- Bulleted List Workflow examples: 1
- General List Workflow examples: 4
- List Formatting Workflow examples: 5
- Nested List Workflow examples: 17
- Numbered List Workflow examples: 2
- Numbering Control Workflow examples: 1

## Category rules that shaped these examples

- Do not invent unsupported list APIs.
- Do not skip assigning ListFormat.List before writing list items.
- Do not forget to remove list formatting.
- Create lists with doc.Lists.Add(ListTemplate).
- Apply the list through builder.ListFormat.List before writing list items.
- Use builder.ListFormat.ListLevelNumber for nested list levels.
- Call builder.ListFormat.RemoveNumbers() after finishing a list.
- Avoid maybe-null dereference when reusing list references.
- Guard document nodes before reading list-related properties.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.5.0

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\lists\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `lists/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

# PowerShell example
Copy-Item ..\lists\create-a-default-numbered-list-in-a-word-document-using-documentbuilder-listformat-applynu.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `create-a-default-numbered-list-in-a-word-document-using-documentbuilder-listformat-applynu.cs` | Create a default numbered list in a Word document using DocumentBuilder.ListFormat.ApplyNumberDefault method. | Numbered List Workflow | docx | mcp |
| 2 | `apply-a-default-bulleted-list-to-paragraphs-with-documentbuilder-listformat-applybulletdef.cs` | Apply a default bulleted list to paragraphs with DocumentBuilder.ListFormat.ApplyBulletDefault method for formatting. | Bulleted List Workflow | docx | mcp |
| 3 | `set-listlevelnumber-to-three-before-adding-a-paragraph-to-create-a-third-level-list-item-i.cs` | Set ListLevelNumber to three before adding a paragraph to create a third-level list item in the document. | Nested List Workflow | docx | mcp |
| 4 | `increase-list-indent-programmatically-using-documentbuilder-listformat-increaseindent-to-m.cs` | Increase list indent programmatically using DocumentBuilder.ListFormat.IncreaseIndent to move paragraph deeper within the current list structure. | List Formatting Workflow | docx | mcp |
| 5 | `decrease-list-indent-with-documentbuilder-listformat-decreaseindent-to-promote-paragraph-t.cs` | Decrease list indent with DocumentBuilder.ListFormat.DecreaseIndent to promote paragraph to higher level in the outline. | List Formatting Workflow | docx | mcp |
| 6 | `remove-numbering-from-a-paragraph-by-calling-documentbuilder-listformat-removenumbers-meth.cs` | Remove numbering from a paragraph by calling DocumentBuilder.ListFormat.RemoveNumbers method to stop list formatting. | Numbered List Workflow | docx | mcp |
| 7 | `create-a-custom-list-object-configure-its-levels-and-add-it-to-document-lists-collection.cs` | Create a custom List object, configure its levels, and add it to Document.Lists collection. | Nested List Workflow | doc | mcp |
| 8 | `add-a-new-list-to-the-document-s-listcollection-and-assign-it-to-multiple-paragraphs-to-sh.cs` | Add a new List to the document's ListCollection and assign it to multiple paragraphs to share formatting. | General List Workflow | docx | mcp |
| 9 | `assign-an-existing-list-to-a-paragraph-by-setting-paragraph-listformat-list-property-to-ap.cs` | Assign an existing List to a paragraph by setting Paragraph.ListFormat.List property to apply the desired list style. | List Formatting Workflow | docx | mcp |
| 10 | `clone-an-existing-list-modify-its-level-start-values-and-apply-the-cloned-list-to-new-sect.cs` | Clone an existing List, modify its level start values, and apply the cloned list to new sections. | Nested List Workflow | docx | mcp |
| 11 | `set-numberstyle-of-a-listlevel-to-upperroman-to-generate-roman-numeral-list-items.cs` | Set NumberStyle of a ListLevel to UpperRoman to generate Roman numeral list items. | Nested List Workflow | docx | mcp |
| 12 | `apply-a-custom-bullet-character-to-a-list-by-setting-listlevel-numberstyle-to-bullet-and-d.cs` | Apply a custom bullet character to a list by setting ListLevel.NumberStyle to Bullet and defining BulletChar. | Nested List Workflow | docx | mcp |
| 13 | `define-startingnumber-for-a-listlevel-to-begin-numbering-at-five-instead-of-default-one.cs` | Define StartingNumber for a ListLevel to begin numbering at five instead of default one. | Nested List Workflow | docx | mcp |
| 14 | `restart-numbering-for-each-new-section-by-resetting-listlevel-startingnumber-before-applyi.cs` | Restart numbering for each new section by resetting ListLevel.StartingNumber before applying the list. | Nested List Workflow | docx | mcp |
| 15 | `create-a-numbered-list-that-restarts-numbering-at-each-new-chapter-by-resetting-startingnu.cs` | Create a numbered list that restarts numbering at each new chapter by resetting StartingNumber on chapter start. | Numbering Control Workflow | docx | mcp |
| 16 | `create-a-nine-level-hierarchical-list-by-defining-properties-for-each-listlevel-up-to-leve.cs` | Create a nine-level hierarchical list by defining properties for each ListLevel up to level nine. | Nested List Workflow | docx | mcp |
| 17 | `create-a-multi-level-list-with-alternating-bullet-and-number-styles-by-configuring-each-li.cs` | Create a multi-level list with alternating bullet and number styles by configuring each ListLevel individually. | Nested List Workflow | docx | mcp |
| 18 | `adjust-listlevel-indentation-by-setting-listlevel-indentation-to-36-points-for-proper-alig.cs` | Adjust ListLevel indentation by setting ListLevel.Indentation to 36 points for proper alignment. | Nested List Workflow | docx | mcp |
| 19 | `apply-a-custom-tab-stop-to-list-levels-to-align-text-after-numbers-using-listlevel-tabposi.cs` | Apply a custom tab stop to list levels to align text after numbers using ListLevel.TabPosition property. | Nested List Workflow | docx | mcp |
| 20 | `configure-tabposition-of-a-listlevel-to-72-points-to-align-text-after-list-number.cs` | Configure TabPosition of a ListLevel to 72 points to align text after list number. | Nested List Workflow | docx | mcp |
| 21 | `validate-that-each-list-in-a-document-does-not-exceed-nine-levels-to-comply-with-api-const.cs` | Validate that each list in a document does not exceed nine levels to comply with API constraints. | Nested List Workflow | docx | mcp |
| 22 | `implement-error-handling-for-attempts-to-create-more-than-nine-list-levels-catching-the-re.cs` | Implement error handling for attempts to create more than nine list levels, catching the resulting exception. | Nested List Workflow | docx | mcp |
| 23 | `ensure-ooxmlcompliance-is-set-higher-than-ecma376-before-saving-document-to-retain-custom.cs` | Ensure OoxmlCompliance is set higher than Ecma376 before saving document to retain custom list settings. | General List Workflow | docx | mcp |
| 24 | `check-document-ooxmlcompliance-property-and-set-to-higher-than-ecma376-when-modifying-list.cs` | Check document OoxmlCompliance property and set to higher than Ecma376 when modifying list definitions. | General List Workflow | docx | mcp |
| 25 | `retrieve-a-specific-list-from-document-lists-by-its-id-and-adjust-its-level-properties-pro.cs` | Retrieve a specific List from Document.Lists by its ID and adjust its level properties programmatically. | Nested List Workflow | docx | mcp |
| 26 | `iterate-through-document-lists-collection-to-modify-all-list-definitions-uniformly-across.cs` | Iterate through Document.Lists collection to modify all list definitions uniformly across the document. | General List Workflow | docx | mcp |
| 27 | `apply-a-uniform-list-style-to-all-lists-by-iterating-over-document-lists-and-updating-each.cs` | Apply a uniform list style to all lists by iterating over Document.Lists and updating each level. | List Formatting Workflow | docx | mcp |
| 28 | `programmatically-increase-list-level-depth-for-a-paragraph-using-documentbuilder-listforma.cs` | Programmatically increase list level depth for a paragraph using DocumentBuilder.ListFormat.IncreaseIndent within a loop. | Nested List Workflow | docx | mcp |
| 29 | `programmatically-decrease-list-level-depth-for-a-paragraph-using-documentbuilder-listforma.cs` | Programmatically decrease list level depth for a paragraph using DocumentBuilder.ListFormat.DecreaseIndent inside a conditional. | Nested List Workflow | docx | mcp |
| 30 | `programmatically-remove-list-formatting-from-selected-paragraphs-while-preserving-their-in.cs` | Programmatically remove list formatting from selected paragraphs while preserving their indentation using ListFormat.RemoveNumbers. | List Formatting Workflow | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Invented list APIs

- Symptom: Code uses unsupported helpers such as CreateList or AddListItem.
- Fix: Use doc.Lists.Add(ListTemplate) and builder.ListFormat.List only.

### List formatting not applied

- Symptom: Paragraphs are written but they are not actual list items.
- Fix: Assign builder.ListFormat.List before writing each list sequence.

### Nested level drift

- Symptom: Nested items appear at the wrong level or never return to the parent level.
- Fix: Set builder.ListFormat.ListLevelNumber explicitly before each nested level transition.

### List mode not cleared

- Symptom: Subsequent paragraphs continue as list items unintentionally.
- Fix: Call builder.ListFormat.RemoveNumbers() immediately after the list is complete.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.5.0 run.
- Preserve file-to-task traceability when updating this folder.
- Keep examples standalone and bootstrap local inputs inside the example whenever external sources are mentioned.
