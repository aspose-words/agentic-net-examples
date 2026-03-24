---
name: lists
description: C# examples for lists using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - lists

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **lists** category.
This folder contains standalone C# examples for lists operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **lists**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (30/30 files) ← category-specific
- `using Aspose.Words.Lists;` (30/30 files)
- `using System;` (24/30 files)
- `using System.Drawing;` (11/30 files)
- `using Aspose.Words.Drawing;` (2/30 files)
- `using System.IO;` (2/30 files)
- `using Aspose.Words.Saving;` (2/30 files)
- `using Aspose.Words.Settings;` (1/30 files)

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
| [add-new-list-document-s-listcollection-assign-it-multip...](./add-new-list-document-s-listcollection-assign-it-multiple-paragraphs-share-formatting.cs) | `Document`, `DocumentBuilder`, `ListFormat` | Add new list document s listcollection assign it multiple paragraphs share fo... |
| [adjust-listlevel-indentation-setting-listlevel-indentat...](./adjust-listlevel-indentation-setting-listlevel-indentation-36-points-proper-alignment.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Adjust listlevel indentation setting listlevel indentation 36 points proper a... |
| [apply-custom-bullet-character-list-setting-listlevel-nu...](./apply-custom-bullet-character-list-setting-listlevel-numberstyle-bullet-defining.cs) | `Font`, `Document`, `DocumentBuilder` | Apply custom bullet character list setting listlevel numberstyle bullet defining |
| [apply-custom-tab-stop-list-levels-align-text-after-numb...](./apply-custom-tab-stop-list-levels-align-text-after-numbers-listlevel-tabposition.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Apply custom tab stop list levels align text after numbers listlevel tabposition |
| [apply-default-bulleted-list-paragraphs-documentbuilder-...](./apply-default-bulleted-list-paragraphs-documentbuilder-listformat-applybulletdefault.cs) | `Document`, `DocumentBuilder`, `ListFormat` | Apply default bulleted list paragraphs documentbuilder listformat applybullet... |
| [apply-uniform-list-style-all-lists-iterating-over-docum...](./apply-uniform-list-style-all-lists-iterating-over-document-lists-updating-each-level.cs) | `Font`, `Document`, `Input` | Apply uniform list style all lists iterating over document lists updating eac... |
| [assign-existing-list-paragraph-setting-paragraph-listfo...](./assign-existing-list-paragraph-setting-paragraph-listformat-list-property-apply.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Assign existing list paragraph setting paragraph listformat list property apply |
| [check-document-ooxmlcompliance-property-set-higher-than...](./check-document-ooxmlcompliance-property-set-higher-than-ecma376-when-modifying-list.cs) | `Document`, `DocumentBuilder`, `ListFormat` | Check document ooxmlcompliance property set higher than ecma376 when modifyin... |
| [clone-existing-list-modify-its-level-start-values-apply...](./clone-existing-list-modify-its-level-start-values-apply-cloned-list-new-sections.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Clone existing list modify its level start values apply cloned list new sections |
| [configure-tabposition-listlevel-72-points-align-text-af...](./configure-tabposition-listlevel-72-points-align-text-after-list-number.cs) | `Document`, `DocumentBuilder`, `ListFormat` | Configure tabposition listlevel 72 points align text after list number |
| [custom-list-object-configure-its-levels-add-it-document...](./custom-list-object-configure-its-levels-add-it-document-lists-collection.cs) | `ListFormat`, `Font`, `Document` | Custom list object configure its levels add it document lists collection |
| [decrease-list-indent-documentbuilder-listformat-decreas...](./decrease-list-indent-documentbuilder-listformat-decreaseindent-promote-paragraph.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Decrease list indent documentbuilder listformat decreaseindent promote paragraph |
| [default-numbered-list-word-document-documentbuilder-lis...](./default-numbered-list-word-document-documentbuilder-listformat-applynumberdefault.cs) | `Document`, `DocumentBuilder`, `ListFormat` | Default numbered list word document documentbuilder listformat applynumberdef... |
| [define-startingnumber-listlevel-begin-numbering-at-five...](./define-startingnumber-listlevel-begin-numbering-at-five-instead-default-one.cs) | `Document`, `DocumentBuilder`, `ListFormat` | Define startingnumber listlevel begin numbering at five instead default one |
| [ensure-ooxmlcompliance-is-set-higher-than-ecma376-befor...](./ensure-ooxmlcompliance-is-set-higher-than-ecma376-before-document-retain-custom-list.cs) | `Document`, `DocumentBuilder`, `OoxmlSaveOptions` | Ensure ooxmlcompliance is set higher than ecma376 before document retain cust... |
| [implement-error-handling-attempts-more-than-nine-list-l...](./implement-error-handling-attempts-more-than-nine-list-levels-catching-resulting.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Implement error handling attempts more than nine list levels catching resulting |
| [increase-list-indent-programmatically-documentbuilder-l...](./increase-list-indent-programmatically-documentbuilder-listformat-increaseindent-move.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Increase list indent programmatically documentbuilder listformat increaseinde... |
| [iterate-through-document-lists-collection-modify-all-li...](./iterate-through-document-lists-collection-modify-all-list-definitions-uniformly.cs) | `Document`, `Input`, `Font` | Iterate through document lists collection modify all list definitions uniformly |
| [multi-level-list-alternating-bullet-number-styles-confi...](./multi-level-list-alternating-bullet-number-styles-configuring-each-listlevel.cs) | `Font`, `ListFormat`, `Document` | Multi level list alternating bullet number styles configuring each listlevel |
| [nine-level-hierarchical-list-defining-properties-each-l...](./nine-level-hierarchical-list-defining-properties-each-listlevel-up-level-nine.cs) | `NumberStyle`, `Font`, `ListFormat` | Nine level hierarchical list defining properties each listlevel up level nine |
| [numbered-list-that-restarts-numbering-at-each-new-chapt...](./numbered-list-that-restarts-numbering-at-each-new-chapter-resetting-startingnumber.cs) | `Document`, `DocumentBuilder`, `ListFormat` | Numbered list that restarts numbering at each new chapter resetting startingn... |
| [programmatically-decrease-list-level-depth-paragraph-do...](./programmatically-decrease-list-level-depth-paragraph-documentbuilder-listformat.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Programmatically decrease list level depth paragraph documentbuilder listformat |
| [programmatically-increase-list-level-depth-paragraph-do...](./programmatically-increase-list-level-depth-paragraph-documentbuilder-listformat.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Programmatically increase list level depth paragraph documentbuilder listformat |
| [programmatically-remove-list-formatting-selected-paragr...](./programmatically-remove-list-formatting-selected-paragraphs-while-preserving-their.cs) | `Document`, `ListFormat`, `ParagraphFormat` | Programmatically remove list formatting selected paragraphs while preserving... |
| [remove-numbering-paragraph-calling-documentbuilder-list...](./remove-numbering-paragraph-calling-documentbuilder-listformat-removenumbers-method.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Remove numbering paragraph calling documentbuilder listformat removenumbers m... |
| [restart-numbering-each-new-section-resetting-listlevel-...](./restart-numbering-each-new-section-resetting-listlevel-startingnumber-before-applying.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Restart numbering each new section resetting listlevel startingnumber before... |
| [retrieve-specific-list-document-lists-its-id-adjust-its...](./retrieve-specific-list-document-lists-its-id-adjust-its-level-properties.cs) | `Document`, `DocumentBuilder`, `ListFormat` | Retrieve specific list document lists its id adjust its level properties |
| [set-listlevelnumber-three-before-adding-paragraph-third...](./set-listlevelnumber-three-before-adding-paragraph-third-level-list-item-document.cs) | `ListFormat`, `Document`, `DocumentBuilder` | Set listlevelnumber three before adding paragraph third level list item document |
| [set-numberstyle-listlevel-upperroman-roman-numeral-list...](./set-numberstyle-listlevel-upperroman-roman-numeral-list-items.cs) | `Document`, `DocumentBuilder`, `Lists` | Set numberstyle listlevel upperroman roman numeral list items |
| [validate-that-each-list-document-does-not-exceed-nine-l...](./validate-that-each-list-document-does-not-exceed-nine-levels-comply-api-constraints.cs) | `Document`, `Input`, `ListLevelCollection` | Validate that each list document does not exceed nine levels comply api const... |

## Category Statistics
- Total examples: 30

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for lists patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
