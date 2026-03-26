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

- `using Aspose.Words;` (30/30 files) ← category-specific
- `using Aspose.Words.Vba;` (30/30 files)
- `using System;` (28/30 files)
- `using System.IO;` (21/30 files)
- `using System.Collections.Generic;` (5/30 files)
- `using System.Text.Json;` (3/30 files)
- `using System.Text.RegularExpressions;` (2/30 files)
- `using System.IO.Compression;` (1/30 files)

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
| [add-error-logging-code-each-vba-module-inserting-standa...](./add-error-logging-code-each-vba-module-inserting-standardized-logging-routine-at.cs) | `Document`, `Input`, `Output_WithLogging` | Add error logging code each vba module inserting standardized logging routine at |
| [add-reference-microsoft-excel-object-library-vbaproject...](./add-reference-microsoft-excel-object-library-vbaproject-verify-its-presence.cs) | `VbaReferenceType`, `Document`, `ArgumentOutOfRangeException` | Add reference microsoft excel object library vbaproject verify its presence |
| [add-reference-microsoft-scripting-runtime-library-its-d...](./add-reference-microsoft-scripting-runtime-library-its-dictionary-object-within-macro.cs) | `VbaProject`, `Document`, `VbaModule` | Add reference microsoft scripting runtime library its dictionary object withi... |
| [batch-process-folder-docm-files-extracting-macro-names-...](./batch-process-folder-docm-files-extracting-macro-names-writing-them-csv-summary-file.cs) | `Document`, `StreamWriter`, `MacroSummary` | Batch process folder docm files extracting macro names writing them csv summa... |
| [clone-entire-vba-project-one-word-document-another-targ...](./clone-entire-vba-project-one-word-document-another-target-document.cs) | `Document`, `VbaProject`, `Modules` | Clone entire vba project one word document another target document |
| [clone-vba-project-template-document-newly-reports-ensur...](./clone-vba-project-template-document-newly-reports-ensure-consistent-macros.cs) | `Document`, `VbaProject`, `Modules` | Clone vba project template document newly reports ensure consistent macros |
| [clone-vba-project-then-remove-all-references-external-l...](./clone-vba-project-then-remove-all-references-external-libraries-self-contained-macro.cs) | `Document`, `Source`, `VbaReferenceType` | Clone vba project then remove all references external libraries self containe... |
| [clonevbaproject-method-duplicate-vba-project-while-pres...](./clonevbaproject-method-duplicate-vba-project-while-preserving-module-order-references.cs) | `VbaProject`, `Document`, `Modules` | Clonevbaproject method duplicate vba project while preserving module order re... |
| [compare-source-code-two-vbamodules-different-documents-...](./compare-source-code-two-vbamodules-different-documents-diff-report.cs) | `Document`, `StreamWriter`, `StringSplitOptions` | Compare source code two vbamodules different documents diff report |
| [configuration-file-specify-which-vba-modules-copy-betwe...](./configuration-file-specify-which-vba-modules-copy-between-documents-during-batch.cs) | `Document`, `VbaProject`, `Collections` | Configuration file specify which vba modules copy between documents during batch |
| [copy-single-vbamodule-source-document-s-vba-project-des...](./copy-single-vbamodule-source-document-s-vba-project-destination-document-s-project.cs) | `VbaProject`, `Document`, `Modules` | Copy single vbamodule source document s vba project destination document s pr... |
| [documentation-report-listing-each-vba-module-its-type-n...](./documentation-report-listing-each-vba-module-its-type-number-lines-code.cs) | `Document`, `DocumentBuilder`, `ParagraphFormat` | Documentation report listing each vba module its type number lines code |
| [docx-file-add-new-vba-module-that-automates-table-forma...](./docx-file-add-new-vba-module-that-automates-table-formatting-updated-document.cs) | `VbaProject`, `Document`, `VbaModule` | Docx file add new vba module that automates table formatting updated document |
| [docx-file-containing-vba-macros-enumerate-all-modules-v...](./docx-file-containing-vba-macros-enumerate-all-modules-vbaproject.cs) | `Document` | Docx file containing vba macros enumerate all modules vbaproject |
| [enumerate-vbaproject-references-filter-out-com-referenc...](./enumerate-vbaproject-references-filter-out-com-references-log-remaining-references.cs) | `Document`, `Input`, `SaveFormat` | Enumerate vbaproject references filter out com references log remaining refer... |
| [export-all-vba-modules-document-zip-archive-maintaining...](./export-all-vba-modules-document-zip-archive-maintaining-original-module-filenames.cs) | `VbaProject`, `Document`, `ZipArchive` | Export all vba modules document zip archive maintaining original module filen... |
| [export-macro-source-code-json-format-including-module-n...](./export-macro-source-code-json-format-including-module-names-code-strings-external.cs) | `Document`, `Collections`, `Text` | Export macro source code json format including module names code strings exte... |
| [import-macro-definitions-json-file-corresponding-vbamod...](./import-macro-definitions-json-file-corresponding-vbamodules-assign-their-source-code.cs) | `Document`, `VbaProject`, `VbaModule` | Import macro definitions json file corresponding vbamodules assign their sour... |
| [import-vba-module-external-bas-file-document-s-vbaproje...](./import-vba-module-external-bas-file-document-s-vbaproject-set-its-name.cs) | `Document`, `VbaModule`, `VbaModuleType` | Import vba module external bas file document s vbaproject set its name |
| [iterate-through-all-vbamodules-document-replace-depreca...](./iterate-through-all-vbamodules-document-replace-deprecated-api-calls-updated.cs) | `Document`, `SourceCode`, `Input` | Iterate through all vbamodules document replace deprecated api calls updated |
| [multiple-docm-files-batch-extract-each-macro-s-source-c...](./multiple-docm-files-batch-extract-each-macro-s-source-code-store-them-separate-files.cs) | `Document`, `DocumentBuilder` | Multiple docm files batch extract each macro s source code store them separat... |
| [new-vbaproject-add-standard-module-assign-custom-macro-...](./new-vbaproject-add-standard-module-assign-custom-macro-code-its-sourcecode-property.cs) | `Document`, `VbaProject`, `VbaModule` | New vbaproject add standard module assign custom macro code its sourcecode pr... |
| [perform-case-insensitive-search-across-all-vba-modules-...](./perform-case-insensitive-search-across-all-vba-modules-deprecated-function-names.cs) | `Document`, `Collections`, `Text` | Perform case insensitive search across all vba modules deprecated function names |
| [remove-all-standard-modules-vbaproject-leaving-only-cla...](./remove-all-standard-modules-vbaproject-leaving-only-class-modules-document.cs) | `Document`, `Input`, `VbaProject` | Remove all standard modules vbaproject leaving only class modules document |
| [remove-existing-reference-vbaproject-s-references-colle...](./remove-existing-reference-vbaproject-s-references-collection-confirm-reference-count.cs) | `Document`, `AppDomain`, `CurrentDomain` | Remove existing reference vbaproject s references collection confirm referenc... |
| [replace-hard-coded-file-paths-macro-source-code-relativ...](./replace-hard-coded-file-paths-macro-source-code-relative-paths-string-manipulation.cs) | `Document`, `Text`, `AppContext` | Replace hard coded file paths macro source code relative paths string manipul... |
| [retrieve-source-code-specific-vbamodule-write-it-text-f...](./retrieve-source-code-specific-vbamodule-write-it-text-file-analysis.cs) | `Document`, `ModuleSource`, `Input` | Retrieve source code specific vbamodule write it text file analysis |
| [update-sourcecode-vbamodule-include-error-handling-erro...](./update-sourcecode-vbamodule-include-error-handling-error-resume-next-statements.cs) | `Document`, `InputDocument`, `OutputDocument` | Update sourcecode vbamodule include error handling error resume next statements |
| [validate-that-loaded-document-s-vbaproject-contains-at-...](./validate-that-loaded-document-s-vbaproject-contains-at-least-one-module-before.cs) | `Document`, `InvalidOperationException`, `VbaProject` | Validate that loaded document s vbaproject contains at least one module before |
| [vba-project-class-module-define-public-method-invoke-it...](./vba-project-class-module-define-public-method-invoke-it-another-module.cs) | `VbaProject`, `VbaModule`, `Document` | Vba project class module define public method invoke it another module |

## Category Statistics
- Total examples: 30

## General Tips
- See parent [AGENTS.md](../AGENTS.md) for repository-wide boundaries and testing guidance.
- Review code examples in this folder for vba-macros patterns.


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
Copy-Item ..\vba-macros\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

### Notes for Agents and Developers

- Treat every `.cs` file in `vba-macros/` as a full console program, not a snippet.
- Run one file at a time by copying it to `Program.cs`.
- If a sample needs input documents, images, fonts, or data files, place them in the temporary project directory before running.
- See the root [AGENTS.md](../AGENTS.md) for repository-wide prerequisites, project file template, and testing guidance.

<!-- AUTOGENERATED:START -->
Updated: 2026-03-24 | Run: `20260315_203408_65134f`
<!-- AUTOGENERATED:END -->
