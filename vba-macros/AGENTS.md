---
name: vba-macros
description: Verified C# examples for VBA project and macro workflows in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - VBA Macros

## Purpose

This folder is a **live, curated example set** for VBA and macro scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free use of documented Aspose.Words APIs for loading, creating, cloning, exporting, and updating VBA projects and modules.

## Non-negotiable conventions

- Always use documented Aspose.Words APIs directly.
- Always create local macro-enabled sample documents when a task refers to an existing file, folder, stream, template, or input asset.
- Save macro content in macro-enabled formats such as `DOCM` or `DOTM`.
- Keep validation narrow and task-specific.
- Do not invent VBA or reference helper APIs.

## Recommended workflow selection

- **VBA workflow**: 34 examples

This category performed best with light primary rules plus a narrow safety patch for source editing, references, and whole-project replacement.

## Validation priorities

1. The code must compile and run without manual input.
2. Required sample inputs must be bootstrapped locally inside the example.
3. Requested VBA project or output results must be produced successfully.
4. Validation should focus only on the exact requested module, source, reference, clone, export, or replacement result.

## File-to-task reference

- `load-a-docx-file-containing-vba-macros-and-enumerate-all-modules-in-the-vbaproject.cs`
  - Task: Load a DOCX file containing VBA macros and enumerate all modules in the VbaProject.
  - Workflow: vba-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `validate-that-a-loaded-document-s-vbaproject-contains-at-least-one-module-before-attemptin.cs`
  - Task: Validate that a loaded document's VbaProject contains at least one module before attempting modifications.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: mcp
- `enumerate-vbaproject-references-filter-out-com-references-and-log-remaining-references-to.cs`
  - Task: Enumerate VbaProject.References, filter out COM references, and log remaining references to console.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `retrieve-the-source-code-of-a-specific-vbamodule-and-write-it-to-a-text-file-for-analysis.cs`
  - Task: Retrieve the source code of a specific VbaModule and write it to a text file for analysis.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `iterate-through-all-vbamodules-in-a-document-and-replace-deprecated-api-calls-with-updated.cs`
  - Task: Iterate through all VbaModules in a document and replace deprecated API calls with updated equivalents.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: mcp
- `perform-a-case-insensitive-search-across-all-vba-modules-for-deprecated-function-names-and.cs`
  - Task: Perform a case‑insensitive search across all VBA modules for deprecated function names and replace them.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `replace-hard-coded-file-paths-in-macro-source-code-with-relative-paths-using-string-manipu.cs`
  - Task: Replace hard‑coded file paths in macro source code with relative paths using string manipulation functions.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `update-the-sourcecode-of-a-vbamodule-to-include-error-handling-using-on-error-resume-next.cs`
  - Task: Update the SourceCode of a VbaModule to include error handling using On Error Resume Next statements.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `iterate-through-all-vbamodules-prepend-a-comment-header-containing-author-and-date-informa.cs`
  - Task: Iterate through all VbaModules, prepend a comment header containing author and date information to each source code.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `add-error-logging-code-to-each-vba-module-by-inserting-a-standardized-logging-routine-at-t.cs`
  - Task: Add error logging code to each VBA module by inserting a standardized logging routine at the beginning.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `create-a-new-vbaproject-add-a-standard-module-and-assign-custom-macro-code-to-its-sourceco.cs`
  - Task: Create a new VbaProject, add a standard module, and assign custom macro code to its SourceCode property.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `create-a-vba-project-with-a-class-module-define-a-public-method-and-invoke-it-from-another.cs`
  - Task: Create a VBA project with a class module, define a public method, and invoke it from another module.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `add-a-reference-to-the-microsoft-excel-object-library-in-a-vbaproject-and-verify-its-prese.cs`
  - Task: Add a reference to the Microsoft Excel Object Library in a VbaProject and verify its presence.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `add-a-reference-to-the-microsoft-scripting-runtime-library-and-use-its-dictionary-object-w.cs`
  - Task: Add a reference to the Microsoft Scripting Runtime library and use its Dictionary object within a macro.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `add-a-reference-to-a-custom-com-library-then-call-its-methods-from-a-newly-added-vba-modul.cs`
  - Task: Add a reference to a custom COM library, then call its methods from a newly added VBA module.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `remove-an-existing-reference-from-a-vbaproject-s-references-collection-and-confirm-the-ref.cs`
  - Task: Remove an existing reference from a VbaProject's References collection and confirm the reference count decreases.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `remove-all-standard-modules-from-a-vbaproject-leaving-only-class-modules-and-save-the-docu.cs`
  - Task: Remove all standard modules from a VbaProject, leaving only class modules, and save the document.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: mcp
- `clone-the-entire-vba-project-from-one-word-document-to-another-and-save-the-target-documen.cs`
  - Task: Clone the entire VBA project from one Word document to another and save the target document.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: mcp
- `use-the-clonevbaproject-method-to-duplicate-a-vba-project-while-preserving-module-order-an.cs`
  - Task: Use the CloneVbaProject method to duplicate a VBA project while preserving module order and references.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `clone-a-vba-project-from-a-template-document-into-newly-generated-reports-to-ensure-consis.cs`
  - Task: Clone a VBA project from a template document into newly generated reports to ensure consistent macros.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: mcp
- `clone-a-vba-project-then-remove-all-references-to-external-libraries-to-create-a-self-cont.cs`
  - Task: Clone a VBA project, then remove all references to external libraries to create a self‑contained macro set.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `copy-a-single-vbamodule-from-a-source-document-s-vba-project-into-a-destination-document-s.cs`
  - Task: Copy a single VbaModule from a source document's VBA project into a destination document's project.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: mcp
- `load-a-docx-file-create-a-new-vba-project-and-copy-selected-modules-from-another-document.cs`
  - Task: Load a DOCX file, create a new VBA project, and copy selected modules from another document into it.
  - Workflow: vba-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `use-a-configuration-file-to-specify-which-vba-modules-to-copy-between-documents-during-bat.cs`
  - Task: Use a configuration file to specify which VBA modules to copy between documents during batch processing.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: mcp
- `import-a-vba-module-from-an-external-bas-file-into-a-document-s-vbaproject-and-set-its-nam.cs`
  - Task: Import a VBA module from an external .bas file into a document's VbaProject and set its name.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: mcp
- `import-macro-definitions-from-a-json-file-create-corresponding-vbamodules-and-assign-their.cs`
  - Task: Import macro definitions from a JSON file, create corresponding VbaModules, and assign their source code.
  - Workflow: vba-workflow
  - Outputs: json
  - Selected engine: mcp
- `export-all-vba-modules-from-a-document-to-a-zip-archive-maintaining-original-module-filena.cs`
  - Task: Export all VBA modules from a document to a ZIP archive, maintaining original module filenames.
  - Workflow: vba-workflow
  - Outputs: doc, zip
  - Selected engine: mcp
- `export-macro-source-code-to-json-format-including-module-names-and-code-strings-for-extern.cs`
  - Task: Export macro source code to JSON format, including module names and code strings, for external processing.
  - Workflow: vba-workflow
  - Outputs: json
  - Selected engine: mcp
- `load-multiple-docm-files-in-a-batch-extract-each-macro-s-source-code-and-store-them-in-sep.cs`
  - Task: Load multiple DOCM files in a batch, extract each macro's source code, and store them in separate files.
  - Workflow: vba-workflow
  - Outputs: docm, doc
  - Selected engine: mcp
- `batch-process-a-folder-of-docm-files-extracting-macro-names-and-writing-them-to-a-csv-summ.cs`
  - Task: Batch process a folder of DOCM files, extracting macro names and writing them to a CSV summary file.
  - Workflow: vba-workflow
  - Outputs: docm, doc, csv
  - Selected engine: mcp
- `compare-the-source-code-of-two-vbamodules-from-different-documents-and-generate-a-diff-rep.cs`
  - Task: Compare the source code of two VbaModules from different documents and generate a diff report.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: mcp
- `generate-a-documentation-report-listing-each-vba-module-its-type-and-the-number-of-lines-o.cs`
  - Task: Generate a documentation report listing each VBA module, its type, and the number of lines of code.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: mcp
- `load-a-docx-file-add-a-new-vba-module-that-automates-table-formatting-and-save-the-updated.cs`
  - Task: Load a DOCX file, add a new VBA module that automates table formatting, and save the updated document.
  - Workflow: vba-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `use-the-document-vbaproject-property-to-replace-an-existing-vba-project-with-a-pre-configu.cs`
  - Task: Use the Document.VbaProject property to replace an existing VBA project with a pre‑configured project template.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Null VBA source text**
  - Symptom: String operations fail because a module has null source text.
  - Preferred fix: Treat null source as an empty string before replacement or analysis.
- **Invented reference APIs**
  - Symptom: Build failures caused by direct `VbaReference` construction or unsupported collection mutation.
  - Preferred fix: Use only compile-safe, documented APIs and rely on template or cloned projects when reference creation is not directly supported.
- **Brittle project replacement validation**
  - Symptom: Project replacement works, but validation fails because of a guessed exact module count.
  - Preferred fix: Validate required modules or expected transferred content instead of a guessed total.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.3.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required package

```bash
dotnet add package Aspose.Words --version 26.3.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\vba-macros\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve exact file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the verified winner from the latest batch report rather than a merely compiling draft.
- Bootstrap file-based inputs locally instead of depending on machine-specific paths.
