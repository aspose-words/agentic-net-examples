---
name: paragraphs
description: C# examples for paragraphs using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - paragraphs

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **paragraphs** category.
This folder contains standalone C# examples for paragraphs operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **paragraphs**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (33/33 files) ← category-specific
- `using System;` (29/33 files)
- `using System.Drawing;` (5/33 files)
- `using Aspose.Words.Tables;` (5/33 files)
- `using Aspose.Words.Saving;` (5/33 files)
- `using Aspose.Words.Drawing;` (4/33 files)
- `using System.IO;` (3/33 files)
- `using Aspose.Words.Loading;` (3/33 files)
- `using System.Text;` (2/33 files)
- `using System.Linq;` (2/33 files)
- `using System.Collections.Generic;` (2/33 files)
- `using Aspose.Words.Fields;` (1/33 files)

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
| [add-paragraph-containing-hyperlink-run-assign-built-hyp...](./add-paragraph-containing-hyperlink-run-assign-built-hyperlink-style-it.cs) | `Font`, `Document`, `DocumentBuilder` | Add paragraph containing hyperlink run assign built hyperlink style it |
| [adjust-paragraph-line-spacing-1-5-lines-setting-builder...](./adjust-paragraph-line-spacing-1-5-lines-setting-builder-currentparagraph.cs) | `Document`, `DocumentBuilder`, `CurrentParagraph` | Adjust paragraph line spacing 1 5 lines setting builder currentparagraph |
| [apply-built-heading1-style-current-paragraph-assigning-...](./apply-built-heading1-style-current-paragraph-assigning-paragraphformat.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Apply built heading1 style current paragraph assigning paragraphformat |
| [apply-custom-paragraph-style-named-mystyle-setting-para...](./apply-custom-paragraph-style-named-mystyle-setting-paragraphformat-stylename-property.cs) | `Font`, `Document`, `Paragraph` | Apply custom paragraph style named mystyle setting paragraphformat stylename... |
| [apply-right-left-paragraph-direction-setting-paragraphf...](./apply-right-left-paragraph-direction-setting-paragraphformat-bidi-property-true.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Apply right left paragraph direction setting paragraphformat bidi property true |
| [apply-specific-paragraph-style-all-heading-paragraphs-s...](./apply-specific-paragraph-style-all-heading-paragraphs-style-selector-loop-across.cs) | `Font`, `ParagraphFormat`, `Document` | Apply specific paragraph style all heading paragraphs style selector loop across |
| [convert-detected-plain-text-list-numbering-word-list-st...](./convert-detected-plain-text-list-numbering-word-list-structures-after-document.cs) | `Document`, `TxtLoadOptions`, `Encoding` | Convert detected plain text list numbering word list structures after document |
| [count-number-paragraphs-that-contain-given-keyword-outp...](./count-number-paragraphs-that-contain-given-keyword-output-total-count-result.cs) | `Document`, `NodeType`, `StringComparison` | Count number paragraphs that contain given keyword output total count result |
| [determine-whether-paragraph-resides-inside-table-checki...](./determine-whether-paragraph-resides-inside-table-checking-paragraph-isintable.cs) | `Document`, `DocumentBuilder`, `NodeType` | Determine whether paragraph resides inside table checking paragraph isintable |
| [enable-bi-directional-marks-before-plain-text-setting-d...](./enable-bi-directional-marks-before-plain-text-setting-document-addbidimarks-property.cs) | `Document`, `DocumentBuilder`, `TxtSaveOptions` | Enable bi directional marks before plain text setting document addbidimarks p... |
| [export-document-txt-format-while-preserving-headers-foo...](./export-document-txt-format-while-preserving-headers-footers-enabling.cs) | `Document`, `TxtSaveOptions`, `Input` | Export document txt format while preserving headers footers enabling |
| [export-document-txt-without-headers-footers-setting-exp...](./export-document-txt-without-headers-footers-setting-exportheadersfooters-option-false.cs) | `Document`, `DocumentBuilder`, `TxtSaveOptions` | Export document txt without headers footers setting exportheadersfooters opti... |
| [export-paragraphs-plain-text-line-numbers-prefixed-enab...](./export-paragraphs-plain-text-line-numbers-prefixed-enabling-line-numbering-option.cs) | `Document`, `DocumentBuilder`, `TxtSaveOptions` | Export paragraphs plain text line numbers prefixed enabling line numbering op... |
| [insert-empty-paragraph-after-bookmark-moving-builder-bo...](./insert-empty-paragraph-after-bookmark-moving-builder-bookmark-calling-insertparagraph.cs) | `Document`, `DocumentBuilder`, `Input` | Insert empty paragraph after bookmark moving builder bookmark calling insertp... |
| [insert-empty-paragraph-at-specific-node-documentbuilder...](./insert-empty-paragraph-at-specific-node-documentbuilder-insertparagraph-precise.cs) | `Document`, `DocumentBuilder`, `FirstSection` | Insert empty paragraph at specific node documentbuilder insertparagraph precise |
| [insert-new-paragraph-custom-text-documentbuilder-writel...](./insert-new-paragraph-custom-text-documentbuilder-writeln-inside-document-loop.cs) | `Document`, `DocumentBuilder`, `InsertedParagraphs` | Insert new paragraph custom text documentbuilder writeln inside document loop |
| [insert-paragraph-field-code-such-as-date-format-field-r...](./insert-paragraph-field-code-such-as-date-format-field-result-paragraph-formatting.cs) | `Document`, `DocumentBuilder`, `Font` | Insert paragraph field code such as date format field result paragraph format... |
| [insert-paragraph-page-break-before-it-setting-paragraph...](./insert-paragraph-page-break-before-it-setting-paragraphformat-pagebreakbefore.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Insert paragraph page break before it setting paragraphformat pagebreakbefore |
| [insert-style-separator-run-between-two-text-runs-combin...](./insert-style-separator-run-between-two-text-runs-combine-different-paragraph-styles.cs) | `Document`, `Font`, `DocumentBuilder` | Insert style separator run between two text runs combine different paragraph... |
| [merge-consecutive-paragraphs-that-share-identical-forma...](./merge-consecutive-paragraphs-that-share-identical-formatting-single-paragraph.cs) | `ParagraphFormat`, `ParagraphAlignment`, `Document` | Merge consecutive paragraphs that share identical formatting single paragraph |
| [modify-paragraph-spacing-before-after-setting-paragraph...](./modify-paragraph-spacing-before-after-setting-paragraphformat-spacebefore-spaceafter.cs) | `ParagraphFormat`, `Document`, `DocumentBuilder` | Modify paragraph spacing before after setting paragraphformat spacebefore spa... |
| [navigate-specific-paragraph-index-documentbuilder-movet...](./navigate-specific-paragraph-index-documentbuilder-movetoparagraph-before-applying.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Navigate specific paragraph index documentbuilder movetoparagraph before appl... |
| [plain-text-file-detectnumberingwithwhitespaces-enabled-...](./plain-text-file-detectnumberingwithwhitespaces-enabled-automatically-recognize-list.cs) | `Document`, `TxtLoadOptions`, `Lists` | Plain text file detectnumberingwithwhitespaces enabled automatically recogniz... |
| [remove-all-paragraphs-that-specific-style-identifier-su...](./remove-all-paragraphs-that-specific-style-identifier-such-as-styleidentifier-quote.cs) | `Document`, `ParagraphFormat`, `Input` | Remove all paragraphs that specific style identifier such as styleidentifier... |
| [replace-text-within-paragraph-while-preserving-its-orig...](./replace-text-within-paragraph-while-preserving-its-original-formatting-run-replace.cs) | `Document`, `Text`, `Input` | Replace text within paragraph while preserving its original formatting run re... |
| [search-document-runs-styleidentifier-styleseparator-ide...](./search-document-runs-styleidentifier-styleseparator-identify-existing-style.cs) | `Document`, `Collections`, `NodeType` | Search document runs styleidentifier styleseparator identify existing style |
| [set-first-line-indent-paragraph-half-inch-paragraphform...](./set-first-line-indent-paragraph-half-inch-paragraphformat-firstlineindent-property.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Set first line indent paragraph half inch paragraphformat firstlineindent pro... |
| [set-paragraph-alignment-center-modifying-builder-curren...](./set-paragraph-alignment-center-modifying-builder-currentparagraph-paragraphformat.cs) | `Document`, `DocumentBuilder`, `CurrentParagraph` | Set paragraph alignment center modifying builder currentparagraph paragraphfo... |
| [set-paragraph-outline-level-2-subheadings-ensure-proper...](./set-paragraph-outline-level-2-subheadings-ensure-proper-inclusion-table-contents.cs) | `ParagraphFormat`, `StyleIdentifier`, `Document` | Set paragraph outline level 2 subheadings ensure proper inclusion table contents |
| [set-paragraph-shading-background-color-light-gray-parag...](./set-paragraph-shading-background-color-light-gray-paragraphformat-shading.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Set paragraph shading background color light gray paragraphformat shading |
| ... | | *and 3 more files* |

## Category Statistics
- Total examples: 33

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for paragraphs patterns.


## Command Reference

### Build and Run

Files in this folder are standalone `.cs` examples. Run one example at a time by copying it into a temporary console project as `Program.cs`.

```bash
# Create a temporary console project from the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\paragraphs\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

### Notes for Agents and Developers

- Treat every `.cs` file in `paragraphs/` as a full console program, not a snippet.
- Run one file at a time by copying it to `Program.cs`.
- If a sample needs input documents, images, fonts, or data files, place them in the temporary project directory before running.
- See the root [AGENTS.md](../AGENTS.md) for repository-wide prerequisites, project file template, and testing guidance.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
