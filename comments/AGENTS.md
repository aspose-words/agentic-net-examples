---
name: comments
description: C# examples for comments using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - comments

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **comments** category.
This folder contains standalone C# examples for comments operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **comments**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (22/22 files) ← category-specific
- `using System;` (21/22 files)
- `using System.IO;` (8/22 files)
- `using System.Linq;` (5/22 files)
- `using Aspose.Words.Saving;` (4/22 files)
- `using Aspose.Words.Drawing;` (3/22 files)
- `using Aspose.Words.Tables;` (3/22 files)
- `using System.Drawing;` (2/22 files)
- `using Aspose.Words.Layout;` (2/22 files)
- `using System.Text;` (1/22 files)
- `using Aspose.Words.Notes;` (1/22 files)
- `using Aspose.Words.AI;` (1/22 files)

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
| [add-comment-containing-hyperlink-external-resource-veri...](./add-comment-containing-hyperlink-external-resource-verify-link-functions-pdf.cs) | `Font`, `Document`, `DocumentBuilder` | Add comment containing hyperlink external resource verify link functions pdf |
| [add-new-comment-specific-paragraph-word-document-as-docx](./add-new-comment-specific-paragraph-word-document-as-docx.cs) | `Document`, `DocumentBuilder`, `Comment` | Add new comment specific paragraph word document as docx |
| [apply-custom-style-all-comment-text-blocks-within-docum...](./apply-custom-style-all-comment-text-blocks-within-document-match-corporate-branding.cs) | `Font`, `Document`, `NodeType` | Apply custom style all comment text blocks within document match corporate br... |
| [batch-process-that-adds-standardized-disclaimer-comment...](./batch-process-that-adds-standardized-disclaimer-comment-every-document-directory.cs) | `Document`, `Comment`, `DocumentBuilder` | Batch process that adds standardized disclaimer comment every document directory |
| [comment-collection-events-trigger-custom-logging-whenev...](./comment-collection-events-trigger-custom-logging-whenever-comment-is-added-removed.cs) | `Comment`, `INodeChangingCallback`, `StringBuilder` | Comment collection events trigger custom logging whenever comment is added re... |
| [convert-doc-file-pdf-while-retaining-all-comment-annota...](./convert-doc-file-pdf-while-retaining-all-comment-annotations-visible-output.cs) | `Document`, `DocumentBuilder`, `Comment` | Convert doc file pdf while retaining all comment annotations visible output |
| [convert-document-comments-xps-format-ensuring-comments-...](./convert-document-comments-xps-format-ensuring-comments-appear-as-markup-annotations.cs) | `Document`, `DocumentBuilder`, `Comment` | Convert document comments xps format ensuring comments appear as markup annot... |
| [document-change-comment-author-names-uppercase-updated-...](./document-change-comment-author-names-uppercase-updated-file.cs) | `Document`, `NodeType`, `Input` | Document change comment author names uppercase updated file |
| [document-stream-add-comments-modified-document-back-mem...](./document-stream-add-comments-modified-document-back-memory-stream.cs) | `Document`, `DocumentBuilder`, `Comment` | Document stream add comments modified document back memory stream |
| [docx-file-enumerate-all-comments-print-each-author-text...](./docx-file-enumerate-all-comments-print-each-author-text-console.cs) | `Document`, `NodeType` | Docx file enumerate all comments print each author text console |
| [export-all-comments-docx-file-csv-file-author-date-text...](./export-all-comments-docx-file-csv-file-author-date-text-columns.cs) | `Document`, `StreamWriter`, `NodeType` | Export all comments docx file csv file author date text columns |
| [extract-comment-text-embed-it-as-footnotes-within-same-...](./extract-comment-text-embed-it-as-footnotes-within-same-document-alternative.cs) | `Document`, `DocumentBuilder`, `Comment` | Extract comment text embed it as footnotes within same document alternative |
| [filter-comments-author-export-only-those-comments-separ...](./filter-comments-author-export-only-those-comments-separate-word-document-review.cs) | `Document`, `DocumentBuilder`, `Comment` | Filter comments author export only those comments separate word document review |
| [import-comments-exported-xml-file-attach-them-appropria...](./import-comments-exported-xml-file-attach-them-appropriate-locations-new-document.cs) | `Document`, `DocumentBuilder`, `Comment` | Import comments exported xml file attach them appropriate locations new document |
| [iterate-through-comment-collection-remove-comments-olde...](./iterate-through-comment-collection-remove-comments-older-than-specified-date-threshold.cs) | `Document`, `Input`, `NodeType` | Iterate through comment collection remove comments older than specified date... |
| [multiple-word-documents-folder-aggregate-their-comments...](./multiple-word-documents-folder-aggregate-their-comments-summary-report.cs) | `Document`, `DocumentBuilder`, `SummarizeOptions` | Multiple word documents folder aggregate their comments summary report |
| [programmatically-accept-reject-comments-based-author-na...](./programmatically-accept-reject-comments-based-author-name-revised-document-version.cs) | `Document`, `Input`, `NodeType` | Programmatically accept reject comments based author name revised document ve... |
| [reply-existing-comment-ensure-reply-appears-nested-unde...](./reply-existing-comment-ensure-reply-appears-nested-under-original.cs) | `Document`, `DocumentBuilder`, `Comment` | Reply existing comment ensure reply appears nested under original |
| [search-comments-containing-specific-keyword-highlight-c...](./search-comments-containing-specific-keyword-highlight-corresponding-text-range-document.cs) | `NodeType`, `Document`, `Font` | Search comments containing specific keyword highlight corresponding text rang... |
| [set-custom-author-name-initials-programmatically-added-...](./set-custom-author-name-initials-programmatically-added-comments-document.cs) | `Document`, `DocumentBuilder`, `Comment` | Set custom author name initials programmatically added comments document |
| [synchronize-comment-positions-after-document-sections-a...](./synchronize-comment-positions-after-document-sections-are-reordered-maintain-accurate.cs) | `Sections`, `NodeType`, `Document` | Synchronize comment positions after document sections are reordered maintain... |
| [validate-that-comment-reference-ids-update-correctly-af...](./validate-that-comment-reference-ids-update-correctly-after-inserting-new-paragraphs.cs) | `PreviousSibling`, `Document`, `DocumentBuilder` | Validate that comment reference ids update correctly after inserting new para... |

## Category Statistics
- Total examples: 22

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for comments patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
