---
name: vba-macros
description: Verified C# examples for vba-macros scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - VBA Macros

## Purpose

An AI coding agent should generate C# code that accesses Document.VbaProject, works with VbaModuleCollection and VbaReferenceCollection, and saves the result as a DOCM file. The agent must rely exclusively on documented Aspose.Words APIs, guard against null objects, and keep validation scoped to the specific macro task.

## Non-negotiable conventions

- Do not call any member that is not present in Aspose.Words 26.10.0.
- Do not instantiate VbaReference or VbaModule with arguments; only use the parameterless constructors.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Vba Workflow workflow**: 34 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. Validate that the loaded document contains a non-null VbaProject and that the Modules collection has the expected number of items before any modification.

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
  - Selected engine: llm
- `enumerate-vbaproject-references-filter-out-com-references-and-log-remaining-references-to.cs`
  - Task: Enumerate VbaProject.References, filter out COM references, and log remaining references to console.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: llm
- `retrieve-the-source-code-of-a-specific-vbamodule-and-write-it-to-a-text-file-for-analysis.cs`
  - Task: Retrieve the source code of a specific VbaModule and write it to a text file for analysis.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: existing_repo
- `iterate-through-all-vbamodules-in-a-document-and-replace-deprecated-api-calls-with-updated.cs`
  - Task: Iterate through all VbaModules in a document and replace deprecated API calls with updated equivalents.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `perform-a-case-insensitive-search-across-all-vba-modules-for-deprecated-function-names-and.cs`
  - Task: Perform a case-insensitive search across all VBA modules for deprecated function names and replace them.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `replace-hard-coded-file-paths-in-macro-source-code-with-relative-paths-using-string-manipu.cs`
  - Task: Replace hard-coded file paths in macro source code with relative paths using string manipulation functions.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: existing_repo
- `update-the-sourcecode-of-a-vbamodule-to-include-error-handling-using-on-error-resume-next.cs`
  - Task: Update the SourceCode of a VbaModule to include error handling using On Error Resume Next statements.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: existing_repo
- `iterate-through-all-vbamodules-prepend-a-comment-header-containing-author-and-date-informa.cs`
  - Task: Iterate through all VbaModules, prepend a comment header containing author and date information to each source code.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: existing_repo
- `add-error-logging-code-to-each-vba-module-by-inserting-a-standardized-logging-routine-at-t.cs`
  - Task: Add error logging code to each VBA module by inserting a standardized logging routine at the beginning.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: existing_repo
- `create-a-new-vbaproject-add-a-standard-module-and-assign-custom-macro-code-to-its-sourceco.cs`
  - Task: Create a new VbaProject, add a standard module, and assign custom macro code to its SourceCode property.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: existing_repo
- `create-a-vba-project-with-a-class-module-define-a-public-method-and-invoke-it-from-another.cs`
  - Task: Create a VBA project with a class module, define a public method, and invoke it from another module.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: llm
- `add-a-reference-to-the-microsoft-excel-object-library-in-a-vbaproject-and-verify-its-prese.cs`
  - Task: Add a reference to the Microsoft Excel Object Library in a VbaProject and verify its presence.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: existing_repo
- `add-a-reference-to-the-microsoft-scripting-runtime-library-and-use-its-dictionary-object-w.cs`
  - Task: Add a reference to the Microsoft Scripting Runtime library and use its Dictionary object within a macro.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `add-a-reference-to-a-custom-com-library-then-call-its-methods-from-a-newly-added-vba-modul.cs`
  - Task: Add a reference to a custom COM library, then call its methods from a newly added VBA module.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: existing_repo
- `remove-an-existing-reference-from-a-vbaproject-s-references-collection-and-confirm-the-ref.cs`
  - Task: Remove an existing reference from a VbaProject's References collection and confirm the reference count decreases.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: mcp
- `remove-all-standard-modules-from-a-vbaproject-leaving-only-class-modules-and-save-the-docu.cs`
  - Task: Remove all standard modules from a VbaProject, leaving only class modules, and save the document.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `clone-the-entire-vba-project-from-one-word-document-to-another-and-save-the-target-documen.cs`
  - Task: Clone the entire VBA project from one Word document to another and save the target document.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `use-the-clonevbaproject-method-to-duplicate-a-vba-project-while-preserving-module-order-an.cs`
  - Task: Use the CloneVbaProject method to duplicate a VBA project while preserving module order and references.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: existing_repo
- `clone-a-vba-project-from-a-template-document-into-newly-generated-reports-to-ensure-consis.cs`
  - Task: Clone a VBA project from a template document into newly generated reports to ensure consistent macros.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `clone-a-vba-project-then-remove-all-references-to-external-libraries-to-create-a-self-cont.cs`
  - Task: Clone a VBA project, then remove all references to external libraries to create a self-contained macro set.
  - Workflow: vba-workflow
  - Outputs: docm
  - Selected engine: existing_repo
- `copy-a-single-vbamodule-from-a-source-document-s-vba-project-into-a-destination-document-s.cs`
  - Task: Copy a single VbaModule from a source document's VBA project into a destination document's project.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `load-a-docx-file-create-a-new-vba-project-and-copy-selected-modules-from-another-document.cs`
  - Task: Load a DOCX file, create a new VBA project, and copy selected modules from another document into it.
  - Workflow: vba-workflow
  - Outputs: docx, doc
  - Selected engine: llm
- `use-a-configuration-file-to-specify-which-vba-modules-to-copy-between-documents-during-bat.cs`
  - Task: Use a configuration file to specify which VBA modules to copy between documents during batch processing.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: existing_repo
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
  - Selected engine: llm
- `load-multiple-docm-files-in-a-batch-extract-each-macro-s-source-code-and-store-them-in-sep.cs`
  - Task: Load multiple DOCM files in a batch, extract each macro's source code, and store them in separate files.
  - Workflow: vba-workflow
  - Outputs: docm, doc
  - Selected engine: mcp
- `batch-process-a-folder-of-docm-files-extracting-macro-names-and-writing-them-to-a-csv-summ.cs`
  - Task: Batch process a folder of DOCM files, extracting macro names and writing them to a CSV summary file.
  - Workflow: vba-workflow
  - Outputs: docm, doc, csv
  - Selected engine: existing_repo
- `compare-the-source-code-of-two-vbamodules-from-different-documents-and-generate-a-diff-rep.cs`
  - Task: Compare the source code of two VbaModules from different documents and generate a diff report.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: llm
- `generate-a-documentation-report-listing-each-vba-module-its-type-and-the-number-of-lines-o.cs`
  - Task: Generate a documentation report listing each VBA module, its type, and the number of lines of code.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: existing_repo
- `load-a-docx-file-add-a-new-vba-module-that-automates-table-formatting-and-save-the-updated.cs`
  - Task: Load a DOCX file, add a new VBA module that automates table formatting, and save the updated document.
  - Workflow: vba-workflow
  - Outputs: docx, doc
  - Selected engine: mcp
- `use-the-document-vbaproject-property-to-replace-an-existing-vba-project-with-a-pre-configu.cs`
  - Task: Use the Document.VbaProject property to replace an existing VBA project with a pre-configured project template.
  - Workflow: vba-workflow
  - Outputs: doc
  - Selected engine: existing_repo

## Common failure patterns and preferred agent fixes

- **Invented API members**
  - Symptom: Compilation fails with CS1501 or CS1061 because the code references constructors, properties, or enum values that do not exist in the selected Aspose.Words version.
  - Root cause: The generator created calls to non-existent members such as VbaModule(string, string) or VbaReferenceType.COM.
  - Preferred fix: Replace invented members with documented APIs: use the parameterless VbaModule constructor, set properties explicitly, and reference only enum members defined in the library.
  - Prevention: Consult the official Aspose.Words VBA API reference for the exact member list before writing code.

- **Incorrect VbaModule construction**
  - Symptom: CS1729 errors indicate that VbaModule does not have a constructor accepting arguments.
  - Root cause: VbaModule provides only a parameterless constructor in version 26.10.0.
  - Preferred fix: Instantiate VbaModule with new VbaModule() and then assign Name, Type, and SourceCode properties.
  - Prevention: Always use the default constructor and configure the object via its properties.

- **Invalid Add overload usage**
  - Symptom: CS1501 errors occur when collection.Add is called with two arguments instead of a single VbaModule instance.
  - Root cause: VbaModuleCollection.Add expects exactly one VbaModule parameter.
  - Preferred fix: Pass only the VbaModule object to Modules.Add.
  - Prevention: Review method signatures and supply the correct number of arguments.

- **Incorrect enum/value usage**
  - Symptom: CS0117 errors appear when referencing VbaReferenceType.COM or VbaReferenceType.Library, which are not defined.
  - Root cause: The enum does not contain those members in the current library version.
  - Preferred fix: Use only the enum values that exist, such as VbaReferenceType.Standard, or filter references by their Type property without relying on missing members.
  - Prevention: Verify enum definitions in the documentation before using them.

- **Wrong member access on VbaReference**
  - Symptom: CS1061 errors are raised when code attempts to read VbaReference.Name, a property that does not exist.
  - Root cause: VbaReference exposes LibId and Type, not Name.
  - Preferred fix: Replace reference.Name with reference.LibId for identification.
  - Prevention: Use property names exactly as defined in the API spec.

- **Missing namespace imports**
  - Symptom: CS0246 errors indicate that types like VbaProject or VbaModule cannot be found.
  - Root cause: The required Aspose.Words.Vba namespace is not imported or the assembly reference is missing.
  - Preferred fix: Add using Aspose.Words.Vba; and ensure the Aspose.Words DLL is referenced by the project.
  - Prevention: Include all namespaces listed in the category metadata at the top of the source file.

- **Over-broad validation**
  - Symptom: Runtime failures occur because validation logic inspects unrelated document parts instead of the VBA project.
  - Root cause: Validation code is not scoped to the specific macro task.
  - Preferred fix: Validate only the presence of a VBA project and the required modules or references.
  - Prevention: Keep validation narrow and task-specific as described in the workflow rules.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.10.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0
```

```powershell
Copy-Item ..\vba-macros\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Review the Aspose.Words release notes periodically and update examples when new constructors, properties, or enum members are added.
