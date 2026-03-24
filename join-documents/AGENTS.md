---
name: join-documents
description: C# examples for join-documents using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - join-documents

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **join-documents** category.
This folder contains standalone C# examples for join-documents operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **join-documents**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using System;` (34/34 files) ← category-specific
- `using Aspose.Words;` (34/34 files)
- `using Aspose.Words.Saving;` (21/34 files)
- `using System.IO;` (13/34 files)
- `using System.Text.RegularExpressions;` (4/34 files)
- `using Aspose.Words.Replacing;` (4/34 files)
- `using System.Data;` (2/34 files)
- `using System.Collections.Generic;` (2/34 files)
- `using Aspose.Words.Tables;` (2/34 files)
- `using System.Linq;` (1/34 files)
- `using Aspose.Words.MailMerging;` (1/34 files)
- `using Aspose.Words.Loading;` (1/34 files)

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
| [append-docx-containing-bibliography-research-paper-docx...](./append-docx-containing-bibliography-research-paper-docx-then-update-fields-export-pdf.cs) | `Document`, `DocumentBuilder`, `PDF` | Append docx containing bibliography research paper docx then update fields ex... |
| [append-docx-containing-custom-styles-destination-docx-i...](./append-docx-containing-custom-styles-destination-docx-importformatmode.cs) | `Document`, `DocumentBuilder`, `Styles` | Append docx containing custom styles destination docx importformatmode |
| [append-docx-containing-footnotes-destination-docx-ensur...](./append-docx-containing-footnotes-destination-docx-ensure-footnote-numbering-continues.cs) | `Document`, `DocumentBuilder`, `ImportFormatMode` | Append docx containing footnotes destination docx ensure footnote numbering c... |
| [append-docx-mail-merge-operation-existing-pdf-converted...](./append-docx-mail-merge-operation-existing-pdf-converted-docx-preserving-destination.cs) | `Document`, `Columns`, `DocumentBuilder` | Append docx mail merge operation existing pdf converted docx preserving desti... |
| [append-docx-rest-api-existing-docx-then-encrypt-final-p...](./append-docx-rest-api-existing-docx-then-encrypt-final-pdf-password.cs) | `Document`, `DocumentBuilder`, `PdfEncryptionDetails` | Append docx rest api existing docx then encrypt final pdf password |
| [append-docx-video-files-destination-docx-ensuring-video...](./append-docx-video-files-destination-docx-ensuring-videos-remain-functional-after-as-pdf.cs) | `Document`, `DocumentBuilder`, `PdfSaveOptions` | Append docx video files destination docx ensuring videos remain functional af... |
| [append-docx-web-service-existing-docx-then-encrypt-fina...](./append-docx-web-service-existing-docx-then-encrypt-final-pdf-password.cs) | `Document`, `DocumentBuilder`, `PdfEncryptionDetails` | Append docx web service existing docx then encrypt final pdf password |
| [append-multiple-docx-files-loop-each-different-importfo...](./append-multiple-docx-files-loop-each-different-importformatmode-settings-output.cs) | `Document`, `ImportFormatMode`, `DocumentBuilder` | Append multiple docx files loop each different importformatmode settings output |
| [append-rtf-document-existing-docx-usedestinationstyles-...](./append-rtf-document-existing-docx-usedestinationstyles-then-export-combined-file-docx.cs) | `Document`, `DocumentBuilder`, `SaveFormat` | Append rtf document existing docx usedestinationstyles then export combined f... |
| [batch-append-docx-files-multiple-subfolders-single-mast...](./batch-append-docx-files-multiple-subfolders-single-master-document-applying.cs) | `Document`, `DocumentBuilder`, `SearchOption` | Batch append docx files multiple subfolders single master document applying |
| [batch-process-directory-docx-files-appending-each-maste...](./batch-process-directory-docx-files-appending-each-master-document-result-as-pdf.cs) | `Document`, `DocumentBuilder`, `Merged` | Batch process directory docx files appending each master document result as pdf |
| [combine-docx-odt-file-single-pdf-importformatmode-keeps...](./combine-docx-odt-file-single-pdf-importformatmode-keepsourceformatting-odt-content.cs) | `Document`, `DocumentBuilder`, `ImportFormatMode` | Combine docx odt file single pdf importformatmode keepsourceformatting odt co... |
| [combine-multiple-docx-files-single-pdf-ensuring-each-so...](./combine-multiple-docx-files-single-pdf-ensuring-each-source-document-retains-its.cs) | `Document`, `ArgumentNullException`, `ArgumentException` | Combine multiple docx files single pdf ensuring each source document retains its |
| [combine-three-odt-files-one-docx-applying-importformatm...](./combine-three-odt-files-one-docx-applying-importformatmode-keepsourceformatting-then.cs) | `Document`, `SaveFormat`, `DocumentBuilder` | Combine three odt files one docx applying importformatmode keepsourceformatti... |
| [documentbuilder-insert-docx-at-end-sections-destination...](./documentbuilder-insert-docx-at-end-sections-destination-docx-then-export-docx.cs) | `Document`, `DocumentBuilder`, `Result` | Documentbuilder insert docx at end sections destination docx then export docx |
| [documentbuilder-insertdocument-importformatmode-usedest...](./documentbuilder-insertdocument-importformatmode-usedestinationstyles-insert-docx.cs) | `DocumentBuilder`, `Document`, `ParagraphFormat` | Documentbuilder insertdocument importformatmode usedestinationstyles insert docx |
| [documentbuilder-insertdocument-insert-docx-at-current-c...](./documentbuilder-insertdocument-insert-docx-at-current-cursor-position.cs) | `Document`, `DocumentBuilder`, `ImportFormatMode` | Documentbuilder insertdocument insert docx at current cursor position |
| [documentbuilder-movetobookmark-navigate-bookmark-named-...](./documentbuilder-movetobookmark-navigate-bookmark-named-content-then-insert-source.cs) | `Document`, `DocumentBuilder`, `InvalidOperationException` | Documentbuilder movetobookmark navigate bookmark named content then insert so... |
| [docx-insert-pdf-converted-docx-at-bookmark-merged-docum...](./docx-insert-pdf-converted-docx-at-bookmark-merged-document-as-odt.cs) | `Document`, `Pdf`, `DocumentBuilder` | Docx insert pdf converted docx at bookmark merged document as odt |
| [docx-template-insert-source-docx-at-bookmark-result-as-pdf](./docx-template-insert-source-docx-at-bookmark-result-as-pdf.cs) | `DocumentBuilder`, `Document`, `ImportFormatMode` | Docx template insert source docx at bookmark result as pdf |
| [findreplaceoptions-replace-phrase-placeholder-docx-docu...](./findreplaceoptions-replace-phrase-placeholder-docx-document-then-as-odt.cs) | `Document`, `DocumentBuilder`, `OdtSaveOptions` | Findreplaceoptions replace phrase placeholder docx document then as odt |
| [findreplaceoptions-replacingcallback-insert-docx-whenev...](./findreplaceoptions-replacingcallback-insert-docx-whenever-phrase-insert-here-is-found.cs) | `Document`, `DocumentBuilder`, `NodeType` | Findreplaceoptions replacingcallback insert docx whenever phrase insert here... |
| [insert-docx-at-bookmark-inside-header-then-document-as-...](./insert-docx-at-bookmark-inside-header-then-document-as-docx-preserving-header.cs) | `Document`, `DocumentBuilder`, `InvalidOperationException` | Insert docx at bookmark inside header then document as docx preserving header |
| [insert-docx-at-bookmark-inside-table-cell-then-export-d...](./insert-docx-at-bookmark-inside-table-cell-then-export-document-odt-preserving-table.cs) | `Document`, `DocumentBuilder`, `OdtSaveOptions` | Insert docx at bookmark inside table cell then export document odt preserving... |
| [insert-docx-at-multiple-bookmarks-named-header-footer-t...](./insert-docx-at-multiple-bookmarks-named-header-footer-then-export-final-document-html.cs) | `DocumentBuilder`, `Document`, `ImportFormatMode` | Insert docx at multiple bookmarks named header footer then export final docum... |
| [insert-docx-during-find-replace-operation-that-matches-...](./insert-docx-during-find-replace-operation-that-matches-word-insertme-then-as-docx.cs) | `Document`, `DocumentBuilder`, `NodeType` | Insert docx during find replace operation that matches word insertme then as... |
| [insert-docx-during-find-replace-operation-that-replaces...](./insert-docx-during-find-replace-operation-that-replaces-xml-tag-then-document-as-docx.cs) | `Document`, `NodeType`, `DocumentBuilder` | Insert docx during find replace operation that replaces xml tag then document... |
| [manually-import-specific-tables-source-docx-nodeimporte...](./manually-import-specific-tables-source-docx-nodeimporter-insert-them-destination-docx.cs) | `Document`, `Paragraph`, `Table` | Manually import specific tables source docx nodeimporter insert them destinat... |
| [nodeimporter-import-only-paragraph-nodes-source-docx-in...](./nodeimporter-import-only-paragraph-nodes-source-docx-insert-them-after-specific.cs) | `Document`, `Section`, `Body` | Nodeimporter import only paragraph nodes source docx insert them after specific |
| [password-protected-docx-insert-source-docx-at-bookmark-...](./password-protected-docx-insert-source-docx-at-bookmark-then-remove-protection-as-docx.cs) | `Document`, `LoadOptions`, `DocumentBuilder` | Password protected docx insert source docx at bookmark then remove protection... |
| ... | | *and 4 more files* |

## Category Statistics
- Total examples: 34

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for join-documents patterns.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
