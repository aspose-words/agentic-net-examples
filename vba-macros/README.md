# VBA Macros Examples for Aspose.Words for .NET

The VBA Macros category illustrates how to work with macro-enabled Word documents using Aspose.Words 26.11.0. Verified snippets demonstrate loading, creating, enumerating, and editing VBA projects, modules, and references while preserving macro integrity.

## Snapshot

- Category: VBA Macros
- Slug: vba-macros
- Total examples: 34
- Publish-ready successful examples: 34 / 34
- Source run: 20261124_190740_1c00f5
- Vba Workflow examples: 34

## Category rules that shaped these examples

- Do not reference members that are not present in Aspose.Words 26.11.0.
- Do not instantiate VbaReference or VbaModule with unsupported constructors.
- Do not create VbaModuleType values that are not defined in the library.
- Access the VBA project via Document.VbaProject and use its Modules and References collections.
- Create VbaModule objects with the parameter-less constructor, then set Name, Type, and SourceCode properties.
- Add modules or references using the collection Add method and save the document with SaveFormat.Docm.
- Check that Document.VbaProject is not null before accessing its collections.
- Verify that a VbaModule's SourceCode property is not null before performing string operations.
- Confirm that the Modules collection contains items before attempting modifications.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.11.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\vba-macros\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `vba-macros/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0

# PowerShell example
Copy-Item ..\vba-macros\load-a-docx-file-containing-vba-macros-and-enumerate-all-modules-in-the-vbaproject.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-containing-vba-macros-and-enumerate-all-modules-in-the-vbaproject.cs` | Load a DOCX file containing VBA macros and enumerate all modules in the VbaProject. | vba-workflow | docx, doc | llm |
| 2 | `validate-that-a-loaded-document-s-vbaproject-contains-at-least-one-module-before-attemptin.cs` | Validate that a loaded document's VbaProject contains at least one module before attempting modifications. | vba-workflow | doc | llm |
| 3 | `enumerate-vbaproject-references-filter-out-com-references-and-log-remaining-references-to.cs` | Enumerate VbaProject.References, filter out COM references, and log remaining references to console. | vba-workflow | docm | mcp |
| 4 | `retrieve-the-source-code-of-a-specific-vbamodule-and-write-it-to-a-text-file-for-analysis.cs` | Retrieve the source code of a specific VbaModule and write it to a text file for analysis. | vba-workflow | docm | mcp |
| 5 | `iterate-through-all-vbamodules-in-a-document-and-replace-deprecated-api-calls-with-updated.cs` | Iterate through all VbaModules in a document and replace deprecated API calls with updated equivalents. | vba-workflow | doc | existing_repo |
| 6 | `perform-a-case-insensitive-search-across-all-vba-modules-for-deprecated-function-names-and.cs` | Perform a case-insensitive search across all VBA modules for deprecated function names and replace them. | vba-workflow | docm | existing_repo |
| 7 | `replace-hard-coded-file-paths-in-macro-source-code-with-relative-paths-using-string-manipu.cs` | Replace hard-coded file paths in macro source code with relative paths using string manipulation functions. | vba-workflow | docm | llm |
| 8 | `update-the-sourcecode-of-a-vbamodule-to-include-error-handling-using-on-error-resume-next.cs` | Update the SourceCode of a VbaModule to include error handling using On Error Resume Next statements. | vba-workflow | docm | existing_repo |
| 9 | `iterate-through-all-vbamodules-prepend-a-comment-header-containing-author-and-date-informa.cs` | Iterate through all VbaModules, prepend a comment header containing author and date information to each source code. | vba-workflow | docm | existing_repo |
| 10 | `add-error-logging-code-to-each-vba-module-by-inserting-a-standardized-logging-routine-at-t.cs` | Add error logging code to each VBA module by inserting a standardized logging routine at the beginning. | vba-workflow | docm | existing_repo |
| 11 | `create-a-new-vbaproject-add-a-standard-module-and-assign-custom-macro-code-to-its-sourceco.cs` | Create a new VbaProject, add a standard module, and assign custom macro code to its SourceCode property. | vba-workflow | docm | mcp |
| 12 | `create-a-vba-project-with-a-class-module-define-a-public-method-and-invoke-it-from-another.cs` | Create a VBA project with a class module, define a public method, and invoke it from another module. | vba-workflow | docm | existing_repo |
| 13 | `add-a-reference-to-the-microsoft-excel-object-library-in-a-vbaproject-and-verify-its-prese.cs` | Add a reference to the Microsoft Excel Object Library in a VbaProject and verify its presence. | vba-workflow | docm | existing_repo |
| 14 | `add-a-reference-to-the-microsoft-scripting-runtime-library-and-use-its-dictionary-object-w.cs` | Add a reference to the Microsoft Scripting Runtime library and use its Dictionary object within a macro. | vba-workflow | docm | mcp |
| 15 | `add-a-reference-to-a-custom-com-library-then-call-its-methods-from-a-newly-added-vba-modul.cs` | Add a reference to a custom COM library, then call its methods from a newly added VBA module. | vba-workflow | docm | mcp |
| 16 | `remove-an-existing-reference-from-a-vbaproject-s-references-collection-and-confirm-the-ref.cs` | Remove an existing reference from a VbaProject's References collection and confirm the reference count decreases. | vba-workflow | docm | mcp |
| 17 | `remove-all-standard-modules-from-a-vbaproject-leaving-only-class-modules-and-save-the-docu.cs` | Remove all standard modules from a VbaProject, leaving only class modules, and save the document. | vba-workflow | doc | existing_repo |
| 18 | `clone-the-entire-vba-project-from-one-word-document-to-another-and-save-the-target-documen.cs` | Clone the entire VBA project from one Word document to another and save the target document. | vba-workflow | doc | existing_repo |
| 19 | `use-the-clonevbaproject-method-to-duplicate-a-vba-project-while-preserving-module-order-an.cs` | Use the CloneVbaProject method to duplicate a VBA project while preserving module order and references. | vba-workflow | docm | existing_repo |
| 20 | `clone-a-vba-project-from-a-template-document-into-newly-generated-reports-to-ensure-consis.cs` | Clone a VBA project from a template document into newly generated reports to ensure consistent macros. | vba-workflow | doc | existing_repo |
| 21 | `clone-a-vba-project-then-remove-all-references-to-external-libraries-to-create-a-self-cont.cs` | Clone a VBA project, then remove all references to external libraries to create a self-contained macro set. | vba-workflow | docm | existing_repo |
| 22 | `copy-a-single-vbamodule-from-a-source-document-s-vba-project-into-a-destination-document-s.cs` | Copy a single VbaModule from a source document's VBA project into a destination document's project. | vba-workflow | doc | mcp |
| 23 | `load-a-docx-file-create-a-new-vba-project-and-copy-selected-modules-from-another-document.cs` | Load a DOCX file, create a new VBA project, and copy selected modules from another document into it. | vba-workflow | docx, doc | existing_repo |
| 24 | `use-a-configuration-file-to-specify-which-vba-modules-to-copy-between-documents-during-bat.cs` | Use a configuration file to specify which VBA modules to copy between documents during batch processing. | vba-workflow | doc | existing_repo |
| 25 | `import-a-vba-module-from-an-external-bas-file-into-a-document-s-vbaproject-and-set-its-nam.cs` | Import a VBA module from an external .bas file into a document's VbaProject and set its name. | vba-workflow | doc | existing_repo |
| 26 | `import-macro-definitions-from-a-json-file-create-corresponding-vbamodules-and-assign-their.cs` | Import macro definitions from a JSON file, create corresponding VbaModules, and assign their source code. | vba-workflow | json | existing_repo |
| 27 | `export-all-vba-modules-from-a-document-to-a-zip-archive-maintaining-original-module-filena.cs` | Export all VBA modules from a document to a ZIP archive, maintaining original module filenames. | vba-workflow | doc, zip | existing_repo |
| 28 | `export-macro-source-code-to-json-format-including-module-names-and-code-strings-for-extern.cs` | Export macro source code to JSON format, including module names and code strings, for external processing. | vba-workflow | json | existing_repo |
| 29 | `load-multiple-docm-files-in-a-batch-extract-each-macro-s-source-code-and-store-them-in-sep.cs` | Load multiple DOCM files in a batch, extract each macro's source code, and store them in separate files. | vba-workflow | docm, doc | existing_repo |
| 30 | `batch-process-a-folder-of-docm-files-extracting-macro-names-and-writing-them-to-a-csv-summ.cs` | Batch process a folder of DOCM files, extracting macro names and writing them to a CSV summary file. | vba-workflow | docm, doc, csv | llm |
| 31 | `compare-the-source-code-of-two-vbamodules-from-different-documents-and-generate-a-diff-rep.cs` | Compare the source code of two VbaModules from different documents and generate a diff report. | vba-workflow | doc | existing_repo |
| 32 | `generate-a-documentation-report-listing-each-vba-module-its-type-and-the-number-of-lines-o.cs` | Generate a documentation report listing each VBA module, its type, and the number of lines of code. | vba-workflow | doc | existing_repo |
| 33 | `load-a-docx-file-add-a-new-vba-module-that-automates-table-formatting-and-save-the-updated.cs` | Load a DOCX file, add a new VBA module that automates table formatting, and save the updated document. | vba-workflow | docx, doc | existing_repo |
| 34 | `use-the-document-vbaproject-property-to-replace-an-existing-vba-project-with-a-pre-configu.cs` | Use the Document.VbaProject property to replace an existing VBA project with a pre-configured project template. | vba-workflow | doc | existing_repo |

## Common failure patterns seen during generation and how they were corrected

### Unsupported API invention

- Symptom: Compilation fails because the code references members that do not exist in the current Aspose.Words version.
- Root cause: The generator invented API members such as nonexistent constructors or enum values.
- Preferred fix: Replace invented members with documented Aspose.Words.Vba APIs that are available in version 26.11.0.
- Prevention: Consult the official API reference and limit code to members that appear in the generated documentation.

### Incorrect VbaModule construction

- Symptom: CS1729 errors indicate that VbaModule does not have a constructor accepting arguments.
- Root cause: The code attempted to use overloaded constructors that are not provided by the library.
- Preferred fix: Instantiate VbaModule with the default constructor and assign Name, Type, and SourceCode properties explicitly.
- Prevention: Always create VbaModule objects using the parameter-less constructor and set properties after creation.

### Invalid VbaModuleType enumeration

- Symptom: CS0117 errors show that VbaModuleType lacks members such as Standard or StandardModule.
- Root cause: The code referenced enum values that were renamed or removed in the current library version.
- Preferred fix: Use the correct enum members, e.g., VbaModuleType.ProceduralModule or VbaModuleType.ClassModule.
- Prevention: Reference the enum definition in the Aspose.Words.Vba namespace to select valid values.

### Incorrect VbaReference usage

- Symptom: CS0117 and CS1061 errors arise when accessing non-existent VbaReference.Type values or Name property.
- Root cause: The code attempted to filter references using undefined enum members like VbaReferenceType.COM.
- Preferred fix: Filter references using the supported VbaReferenceType values (e.g., Project) and access available properties such as LibId.
- Prevention: Review the VbaReference class documentation to use only the properties and enum members that exist.

### Missing local bootstrap inputs

- Symptom: Runtime failures occur because the example assumes external files or folders already exist.
- Root cause: The snippet does not create sample documents, modules, or output files before operating on them.
- Preferred fix: Create deterministic local inputs (e.g., a new Document, sample VbaModule, or temporary output file) within the example.
- Prevention: Never rely on pre-existing files; always generate required inputs programmatically.

### Over-broad validation

- Symptom: Runtime errors appear when the code validates unrelated document internals instead of the requested condition.
- Root cause: The validation logic checks properties that are irrelevant to the task, causing false negatives.
- Preferred fix: Validate only the specific condition required by the task, such as the presence of a VBA project or module count.
- Prevention: Scope validation to the exact prerequisite of the operation and avoid generic document checks.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Update examples whenever a new Aspose.Words version changes the VBA API surface to keep the documentation accurate.
