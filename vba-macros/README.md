# VBA Macros Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **VBA Macros** category. Each file is a standalone example selected from the latest verified generation run and aligned with the active category rules.

## Snapshot

- Category: **VBA Macros**
- Slug: **vba-macros**
- Total examples: **34**
- Verified winners: **both=0**, **mcp-only=34**, **llm-only=0**, **none=0**
- Workflow examples: **34 / 34** use the standard VBA workflow

## Category rules that shaped these examples

- Use native Aspose.Words APIs directly.
- Create local macro-enabled documents when a task refers to an existing document or template.
- Do not assume external files already exist.
- Save VBA content in macro-enabled formats such as `DOCM` or `DOTM`.
- Keep validation narrow and task-specific.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.3.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
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

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# PowerShell example
Copy-Item ..\vba-macros\load-a-docx-file-containing-vba-macros-and-enumerate-all-modules-in-the-vbaproject.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-file-containing-vba-macros-and-enumerate-all-modules-in-the-vbaproject.cs` | Load a DOCX file containing VBA macros and enumerate all modules in the VbaProject. | vba-workflow | docx, doc | mcp |
| 2 | `validate-that-a-loaded-document-s-vbaproject-contains-at-least-one-module-before-attemptin.cs` | Validate that a loaded document's VbaProject contains at least one module before attempting modifications. | vba-workflow | doc | mcp |
| 3 | `enumerate-vbaproject-references-filter-out-com-references-and-log-remaining-references-to.cs` | Enumerate VbaProject.References, filter out COM references, and log remaining references to console. | vba-workflow | docm | mcp |
| 4 | `retrieve-the-source-code-of-a-specific-vbamodule-and-write-it-to-a-text-file-for-analysis.cs` | Retrieve the source code of a specific VbaModule and write it to a text file for analysis. | vba-workflow | docm | mcp |
| 5 | `iterate-through-all-vbamodules-in-a-document-and-replace-deprecated-api-calls-with-updated.cs` | Iterate through all VbaModules in a document and replace deprecated API calls with updated equivalents. | vba-workflow | doc | mcp |
| 6 | `perform-a-case-insensitive-search-across-all-vba-modules-for-deprecated-function-names-and.cs` | Perform a case‑insensitive search across all VBA modules for deprecated function names and replace them. | vba-workflow | docm | mcp |
| 7 | `replace-hard-coded-file-paths-in-macro-source-code-with-relative-paths-using-string-manipu.cs` | Replace hard‑coded file paths in macro source code with relative paths using string manipulation functions. | vba-workflow | docm | mcp |
| 8 | `update-the-sourcecode-of-a-vbamodule-to-include-error-handling-using-on-error-resume-next.cs` | Update the SourceCode of a VbaModule to include error handling using On Error Resume Next statements. | vba-workflow | docm | mcp |
| 9 | `iterate-through-all-vbamodules-prepend-a-comment-header-containing-author-and-date-informa.cs` | Iterate through all VbaModules, prepend a comment header containing author and date information to each source code. | vba-workflow | docm | mcp |
| 10 | `add-error-logging-code-to-each-vba-module-by-inserting-a-standardized-logging-routine-at-t.cs` | Add error logging code to each VBA module by inserting a standardized logging routine at the beginning. | vba-workflow | docm | mcp |
| 11 | `create-a-new-vbaproject-add-a-standard-module-and-assign-custom-macro-code-to-its-sourceco.cs` | Create a new VbaProject, add a standard module, and assign custom macro code to its SourceCode property. | vba-workflow | docm | mcp |
| 12 | `create-a-vba-project-with-a-class-module-define-a-public-method-and-invoke-it-from-another.cs` | Create a VBA project with a class module, define a public method, and invoke it from another module. | vba-workflow | docm | mcp |
| 13 | `add-a-reference-to-the-microsoft-excel-object-library-in-a-vbaproject-and-verify-its-prese.cs` | Add a reference to the Microsoft Excel Object Library in a VbaProject and verify its presence. | vba-workflow | docm | mcp |
| 14 | `add-a-reference-to-the-microsoft-scripting-runtime-library-and-use-its-dictionary-object-w.cs` | Add a reference to the Microsoft Scripting Runtime library and use its Dictionary object within a macro. | vba-workflow | docm | mcp |
| 15 | `add-a-reference-to-a-custom-com-library-then-call-its-methods-from-a-newly-added-vba-modul.cs` | Add a reference to a custom COM library, then call its methods from a newly added VBA module. | vba-workflow | docm | mcp |
| 16 | `remove-an-existing-reference-from-a-vbaproject-s-references-collection-and-confirm-the-ref.cs` | Remove an existing reference from a VbaProject's References collection and confirm the reference count decreases. | vba-workflow | docm | mcp |
| 17 | `remove-all-standard-modules-from-a-vbaproject-leaving-only-class-modules-and-save-the-docu.cs` | Remove all standard modules from a VbaProject, leaving only class modules, and save the document. | vba-workflow | doc | mcp |
| 18 | `clone-the-entire-vba-project-from-one-word-document-to-another-and-save-the-target-documen.cs` | Clone the entire VBA project from one Word document to another and save the target document. | vba-workflow | doc | mcp |
| 19 | `use-the-clonevbaproject-method-to-duplicate-a-vba-project-while-preserving-module-order-an.cs` | Use the CloneVbaProject method to duplicate a VBA project while preserving module order and references. | vba-workflow | docm | mcp |
| 20 | `clone-a-vba-project-from-a-template-document-into-newly-generated-reports-to-ensure-consis.cs` | Clone a VBA project from a template document into newly generated reports to ensure consistent macros. | vba-workflow | doc | mcp |
| 21 | `clone-a-vba-project-then-remove-all-references-to-external-libraries-to-create-a-self-cont.cs` | Clone a VBA project, then remove all references to external libraries to create a self‑contained macro set. | vba-workflow | docm | mcp |
| 22 | `copy-a-single-vbamodule-from-a-source-document-s-vba-project-into-a-destination-document-s.cs` | Copy a single VbaModule from a source document's VBA project into a destination document's project. | vba-workflow | doc | mcp |
| 23 | `load-a-docx-file-create-a-new-vba-project-and-copy-selected-modules-from-another-document.cs` | Load a DOCX file, create a new VBA project, and copy selected modules from another document into it. | vba-workflow | docx, doc | mcp |
| 24 | `use-a-configuration-file-to-specify-which-vba-modules-to-copy-between-documents-during-bat.cs` | Use a configuration file to specify which VBA modules to copy between documents during batch processing. | vba-workflow | doc | mcp |
| 25 | `import-a-vba-module-from-an-external-bas-file-into-a-document-s-vbaproject-and-set-its-nam.cs` | Import a VBA module from an external .bas file into a document's VbaProject and set its name. | vba-workflow | doc | mcp |
| 26 | `import-macro-definitions-from-a-json-file-create-corresponding-vbamodules-and-assign-their.cs` | Import macro definitions from a JSON file, create corresponding VbaModules, and assign their source code. | vba-workflow | json | mcp |
| 27 | `export-all-vba-modules-from-a-document-to-a-zip-archive-maintaining-original-module-filena.cs` | Export all VBA modules from a document to a ZIP archive, maintaining original module filenames. | vba-workflow | doc, zip | mcp |
| 28 | `export-macro-source-code-to-json-format-including-module-names-and-code-strings-for-extern.cs` | Export macro source code to JSON format, including module names and code strings, for external processing. | vba-workflow | json | mcp |
| 29 | `load-multiple-docm-files-in-a-batch-extract-each-macro-s-source-code-and-store-them-in-sep.cs` | Load multiple DOCM files in a batch, extract each macro's source code, and store them in separate files. | vba-workflow | docm, doc | mcp |
| 30 | `batch-process-a-folder-of-docm-files-extracting-macro-names-and-writing-them-to-a-csv-summ.cs` | Batch process a folder of DOCM files, extracting macro names and writing them to a CSV summary file. | vba-workflow | docm, doc, csv | mcp |
| 31 | `compare-the-source-code-of-two-vbamodules-from-different-documents-and-generate-a-diff-rep.cs` | Compare the source code of two VbaModules from different documents and generate a diff report. | vba-workflow | doc | mcp |
| 32 | `generate-a-documentation-report-listing-each-vba-module-its-type-and-the-number-of-lines-o.cs` | Generate a documentation report listing each VBA module, its type, and the number of lines of code. | vba-workflow | doc | mcp |
| 33 | `load-a-docx-file-add-a-new-vba-module-that-automates-table-formatting-and-save-the-updated.cs` | Load a DOCX file, add a new VBA module that automates table formatting, and save the updated document. | vba-workflow | docx, doc | mcp |
| 34 | `use-the-document-vbaproject-property-to-replace-an-existing-vba-project-with-a-pre-configu.cs` | Use the Document.VbaProject property to replace an existing VBA project with a pre‑configured project template. | vba-workflow | doc | mcp |

## Common failure patterns seen during generation and how they were corrected

### Editing null VBA source text

- Symptom: Runtime failures when string or Regex operations are performed on a module with null source code.
- Fix: Treat null source as an empty string before doing replacements or inserts.

### Inventing unsupported reference APIs

- Symptom: Build failures from direct `VbaReference` construction, fake reference types, or unsupported mutation APIs on reference collections.
- Fix: Use compile-safe workflows such as loading or cloning a preconfigured project when direct reference construction is not supported.

### Over-validating whole-project replacement

- Symptom: Project replacement succeeds, but validation fails because it assumes an exact guessed module count.
- Fix: Validate the presence of required modules or transferred content rather than a brittle guessed total.

## Notes for maintainers

- The selected file for each task is the verified winner recorded in the batch run.
- This category performed best with light primary rules plus a narrow safety patch for source editing, references, and project replacement.
- Preserve exact file-to-task traceability when updating the category.
- Bootstrap all sample input files locally inside the example when the task refers to an existing asset.
