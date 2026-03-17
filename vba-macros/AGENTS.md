---
name: vba-macros
description: C# examples for vba-macros using Aspose.Words for .NET
language: csharp
framework: net8.0
parent: ../AGENTS.md
---

# AGENTS - vba-macros

## Persona

You are a C# developer specializing in Word processing using Aspose.Words for .NET,
working within the **vba-macros** category.
This folder contains standalone C# examples for vba-macros operations.
See the root [AGENTS.md](../AGENTS.md) for repository-wide conventions and boundaries.

## Scope
- This folder contains examples for **vba-macros**.
- Files are standalone `.cs` examples stored directly in this folder.

## Required Namespaces

- `using Aspose.Words;` (33/33 files) ← category-specific
- `using Aspose.Words.Vba;` (33/33 files)
- `using System;` (28/33 files)
- `using System.IO;` (10/33 files)
- `using System.Collections.Generic;` (6/33 files)
- `using System.Text.Json;` (3/33 files)
- `using System.Text.RegularExpressions;` (2/33 files)
- `using System.Linq;` (1/33 files)
- `using System.IO.Compression;` (1/33 files)

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
| [add-error-logging-code-each-vba-module-inserting-standa...](./add-error-logging-code-each-vba-module-inserting-standardized-logging-routine-at.cs) | `Document`, `Aspose`, `Words` | Add error logging code each vba module inserting standardized logging routine at |
| [add-reference-microsoft-excel-object-library-vbaproject...](./add-reference-microsoft-excel-object-library-vbaproject-verify-its-presence.cs) | `VbaReferenceType`, `Aspose`, `Document` | Add reference microsoft excel object library vbaproject verify its presence |
| [add-reference-microsoft-scripting-runtime-library-its-d...](./add-reference-microsoft-scripting-runtime-library-its-dictionary-object-within-macro.cs) | `VbaProject`, `Document`, `VbaModule` | Add reference microsoft scripting runtime library its dictionary object withi... |
| [batch-process-folder-docm-files-extracting-macro-names-...](./batch-process-folder-docm-files-extracting-macro-names-writing-them-csv-summary-file.cs) | `Document`, `System`, `Aspose` | Batch process folder docm files extracting macro names writing them csv summa... |
| [clone-entire-vba-project-one-word-document-another-targ...](./clone-entire-vba-project-one-word-document-another-target-document.cs) | `VbaProject`, `Document`, `Aspose` | Clone entire vba project one word document another target document |
| [clone-vba-project-template-document-newly-reports-ensur...](./clone-vba-project-template-document-newly-reports-ensure-consistent-macros.cs) | `VbaProject`, `Document`, `Aspose` | Clone vba project template document newly reports ensure consistent macros |
| [clone-vba-project-then-remove-all-references-external-l...](./clone-vba-project-then-remove-all-references-external-libraries-self-contained-macro.cs) | `Document`, `Aspose`, `Words` | Clone vba project then remove all references external libraries self containe... |
| [clonevbaproject-method-duplicate-vba-project-while-pres...](./clonevbaproject-method-duplicate-vba-project-while-preserving-module-order-references.cs) | `VbaProject`, `Document`, `Modules` | Clonevbaproject method duplicate vba project while preserving module order re... |
| [compare-source-code-two-vbamodules-different-documents-...](./compare-source-code-two-vbamodules-different-documents-diff-report.cs) | `Document`, `StreamWriter`, `System` | Compare source code two vbamodules different documents diff report |
| [configuration-file-specify-which-vba-modules-copy-betwe...](./configuration-file-specify-which-vba-modules-copy-between-documents-during-batch.cs) | `Document`, `System`, `Aspose` | Configuration file specify which vba modules copy between documents during batch |
| [copy-single-vbamodule-source-document-s-vba-project-des...](./copy-single-vbamodule-source-document-s-vba-project-destination-document-s-project.cs) | `VbaProject`, `Document`, `Aspose` | Copy single vbamodule source document s vba project destination document s pr... |
| [document-vbaproject-property-replace-existing-vba-proje...](./document-vbaproject-property-replace-existing-vba-project-pre-configured-project.cs) | `Document`, `Aspose`, `Words` | Document vbaproject property replace existing vba project pre configured project |
| [documentation-report-listing-each-vba-module-its-type-n...](./documentation-report-listing-each-vba-module-its-type-number-lines-code.cs) | `Document`, `DocumentBuilder`, `Aspose` | Documentation report listing each vba module its type number lines code |
| [docx-file-add-new-vba-module-that-automates-table-forma...](./docx-file-add-new-vba-module-that-automates-table-formatting-updated-document.cs) | `VbaProject`, `Document`, `VbaModule` | Docx file add new vba module that automates table formatting updated document |
| [docx-file-containing-vba-macros-enumerate-all-modules-v...](./docx-file-containing-vba-macros-enumerate-all-modules-vbaproject.cs) | `Document`, `Aspose`, `Words` | Docx file containing vba macros enumerate all modules vbaproject |
| [docx-file-new-vba-project-copy-selected-modules-another...](./docx-file-new-vba-project-copy-selected-modules-another-document-it.cs) | `Document`, `VbaProject`, `Aspose` | Docx file new vba project copy selected modules another document it |
| [enumerate-vbaproject-references-filter-out-com-referenc...](./enumerate-vbaproject-references-filter-out-com-references-log-remaining-references.cs) | `Document`, `Aspose`, `Words` | Enumerate vbaproject references filter out com references log remaining refer... |
| [export-all-vba-modules-document-zip-archive-maintaining...](./export-all-vba-modules-document-zip-archive-maintaining-original-module-filenames.cs) | `VbaProject`, `Document`, `ZipArchive` | Export all vba modules document zip archive maintaining original module filen... |
| [export-macro-source-code-json-format-including-module-n...](./export-macro-source-code-json-format-including-module-names-code-strings-external.cs) | `System`, `Aspose`, `Document` | Export macro source code json format including module names code strings exte... |
| [import-macro-definitions-json-file-corresponding-vbamod...](./import-macro-definitions-json-file-corresponding-vbamodules-assign-their-source-code.cs) | `System`, `Document`, `VbaProject` | Import macro definitions json file corresponding vbamodules assign their sour... |
| [import-vba-module-external-bas-file-document-s-vbaproje...](./import-vba-module-external-bas-file-document-s-vbaproject-set-its-name.cs) | `Document`, `VbaProject`, `Aspose` | Import vba module external bas file document s vbaproject set its name |
| [iterate-through-all-vbamodules-document-replace-depreca...](./iterate-through-all-vbamodules-document-replace-deprecated-api-calls-updated.cs) | `Document`, `Aspose`, `SourceCode` | Iterate through all vbamodules document replace deprecated api calls updated |
| [iterate-through-all-vbamodules-prepend-comment-header-c...](./iterate-through-all-vbamodules-prepend-comment-header-containing-author-date.cs) | `Aspose`, `Document`, `Words` | Iterate through all vbamodules prepend comment header containing author date |
| [multiple-docm-files-batch-extract-each-macro-s-source-c...](./multiple-docm-files-batch-extract-each-macro-s-source-code-store-them-separate-files.cs) | `Document`, `DocumentBuilder`, `Aspose` | Multiple docm files batch extract each macro s source code store them separat... |
| [new-vbaproject-add-standard-module-assign-custom-macro-...](./new-vbaproject-add-standard-module-assign-custom-macro-code-its-sourcecode-property.cs) | `Document`, `Aspose`, `VbaProject` | New vbaproject add standard module assign custom macro code its sourcecode pr... |
| [perform-case-insensitive-search-across-all-vba-modules-...](./perform-case-insensitive-search-across-all-vba-modules-deprecated-function-names.cs) | `Document`, `System`, `Aspose` | Perform case insensitive search across all vba modules deprecated function names |
| [remove-all-standard-modules-vbaproject-leaving-only-cla...](./remove-all-standard-modules-vbaproject-leaving-only-class-modules-document.cs) | `Document`, `Aspose`, `Words` | Remove all standard modules vbaproject leaving only class modules document |
| [remove-existing-reference-vbaproject-s-references-colle...](./remove-existing-reference-vbaproject-s-references-collection-confirm-reference-count.cs) | `Document`, `Aspose`, `Words` | Remove existing reference vbaproject s references collection confirm referenc... |
| [replace-hard-coded-file-paths-macro-source-code-relativ...](./replace-hard-coded-file-paths-macro-source-code-relative-paths-string-manipulation.cs) | `Document`, `System`, `Aspose` | Replace hard coded file paths macro source code relative paths string manipul... |
| [retrieve-source-code-specific-vbamodule-write-it-text-f...](./retrieve-source-code-specific-vbamodule-write-it-text-file-analysis.cs) | `Document`, `Aspose`, `ModuleSource` | Retrieve source code specific vbamodule write it text file analysis |
| ... | | *and 3 more files* |

## Category Statistics
- Total examples: 33

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for:
  - **Boundaries** — Always / Ask First / Never rules for all examples
  - **Common Mistakes** — verified anti-patterns that cause build failures
  - **Domain Knowledge** — cross-cutting API-specific gotchas
  - **Testing Guide** — build and run verification steps
- Review code examples in this folder for vba-macros patterns

<!-- AUTOGENERATED:START -->
Updated: 2026-03-16 | Run: `20260316_082635`
<!-- AUTOGENERATED:END -->
