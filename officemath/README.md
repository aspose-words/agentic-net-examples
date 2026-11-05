# OfficeMath Examples for Aspose.Words for .NET

The OfficeMath category addresses the handling of mathematical equations represented by Aspose.Words.OfficeMath nodes within Word documents. Verified examples illustrate loading, enumerating, inserting, justifying, cloning, and replacing OfficeMath objects using the native EQ-field bootstrap workflow.

## Snapshot

- Category: OfficeMath
- Slug: officemath
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261102_031946_08fb5b
- Inspection Reporting examples: 7
- Formatting Display examples: 11
- Creation Modification examples: 7
- Native Officemath examples: 2
- Export Validation examples: 3

## Category rules that shaped these examples

- Do not call any API that is not part of Aspose.Words.Math, such as InsertMathObject, MathParser, or LoadMathML.
- Use only the documented properties DisplayType, Justification, and MathObjectType on top-level OfficeMath nodes.
- Check for null after AsOfficeMath and before inserting or replacing nodes to avoid NullReferenceException.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.10.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0

# Copy one example from this folder into the project as Program.cs
# PowerShell:
Copy-Item ..\officemath\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `officemath/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.10.0

# PowerShell example
Copy-Item ..\officemath\load-a-docx-document-containing-officemath-equations-and-enumerate-each-equation-node.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-document-containing-officemath-equations-and-enumerate-each-equation-node.cs` | Load a DOCX document containing OfficeMath equations and enumerate each equation node. | inspection-reporting | docx, txt | mcp |
| 2 | `load-multiple-docx-files-iterate-officemath-equations-and-standardize-justification-across.cs` | Load multiple DOCX files, iterate OfficeMath equations, and standardize justification across all documents. | formatting-display | docx | existing_repo |
| 3 | `create-a-new-officemath-object-with-documentbuilder-and-insert-it-as-an-inline-equation.cs` | Create a new OfficeMath object with DocumentBuilder and insert it as an inline equation. | formatting-display | docx | mcp |
| 4 | `insert-an-officemath-equation-from-a-latex-string-representation-into-a-specific-document.cs` | Insert an OfficeMath equation from a LaTeX string representation into a specific document location. | creation-modification | docx | mcp |
| 5 | `insert-an-officemath-equation-from-a-mathml-string-into-a-paragraph-using-documentbuilder.cs` | Insert an OfficeMath equation from a MathML string into a paragraph using DocumentBuilder. | creation-modification | docx | mcp |
| 6 | `create-a-batch-process-that-inserts-a-predefined-officemath-equation-into-each-paragraph-o.cs` | Create a batch process that inserts a predefined OfficeMath equation into each paragraph of a document. | creation-modification | docx | mcp |
| 7 | `replace-the-content-of-an-existing-officemath-object-with-a-new-equation-defined-by-a-stri.cs` | Replace the content of an existing OfficeMath object with a new equation defined by a string. | creation-modification | docx | llm |
| 8 | `load-a-docx-file-replace-specific-officemath-equations-based-on-identifier-and-save-as-doc.cs` | Load a DOCX file, replace specific OfficeMath equations based on identifier, and save as DOCX. | creation-modification | docx | llm |
| 9 | `delete-an-unwanted-officemath-node-from-the-document-and-adjust-surrounding-paragraph-spac.cs` | Delete an unwanted OfficeMath node from the document and adjust surrounding paragraph spacing. | creation-modification | docx | mcp |
| 10 | `clone-an-existing-officemath-object-and-insert-the-clone-at-a-new-location.cs` | Clone an existing OfficeMath object and insert the clone at a new location. | creation-modification | docx | mcp |
| 11 | `set-the-justification-of-an-officemath-equation-to-center-alignment-using-the-justificatio.cs` | Set the justification of an OfficeMath equation to center alignment using the Justification property. | formatting-display | docx | mcp |
| 12 | `update-the-justification-of-all-officemath-equations-to-right-alignment-in-a-template-docu.cs` | Update the justification of all OfficeMath equations to right alignment in a template document. | formatting-display | docx | llm |
| 13 | `apply-a-uniform-justification-setting-to-all-officemath-equations-across-multiple-sections.cs` | Apply a uniform justification setting to all OfficeMath equations across multiple sections of a large document. | formatting-display | docx | existing_repo |
| 14 | `validate-that-after-bulk-justification-changes-no-officemath-equation-exceeds-the-page-mar.cs` | Validate that after bulk justification changes, no OfficeMath equation exceeds the page margin limits. | formatting-display | docx | existing_repo |
| 15 | `configure-an-officemath-object-to-display-on-a-separate-line-rather-than-inline-within-the.cs` | Configure an OfficeMath object to display on a separate line rather than inline within the paragraph. | formatting-display | docx | llm_secondary |
| 16 | `programmatically-change-all-officemath-equations-from-inline-to-separate-line-display-in-a.cs` | Programmatically change all OfficeMath equations from inline to separate line display in a large report. | formatting-display | docx | llm |
| 17 | `replace-all-inline-officemath-equations-with-separate-line-display-to-enhance-visual-clari.cs` | Replace all inline OfficeMath equations with separate line display to enhance visual clarity in presentations. | formatting-display | docx | llm_secondary |
| 18 | `programmatically-set-the-display-mode-of-officemath-equations-to-inline-for-compact-docume.cs` | Programmatically set the display mode of OfficeMath equations to inline for compact document layout. | formatting-display | docx | mcp |
| 19 | `retrieve-the-mathobjecttype-of-each-officemath-node-to-determine-whether-it-is-a-fraction.cs` | Retrieve the MathObjectType of each OfficeMath node to determine whether it is a fraction or radical. | inspection-reporting | docx | mcp |
| 20 | `read-mathobjecttype-of-officemath-nodes-and-log-any-unsupported-equation-types-for-review.cs` | Read MathObjectType of OfficeMath nodes and log any unsupported equation types for review. | inspection-reporting | docx | existing_repo |
| 21 | `create-a-function-that-returns-true-if-an-officemath-node-matches-a-specified-mathobjectty.cs` | Create a function that returns true if an OfficeMath node matches a specified MathObjectType criteria. | inspection-reporting | docx | llm |
| 22 | `generate-a-report-listing-each-officemath-equation-s-mathobjecttype-and-its-position-withi.cs` | Generate a report listing each OfficeMath equation's MathObjectType and its position within the document. | inspection-reporting | docx, txt | mcp |
| 23 | `extract-all-officemath-equations-from-a-document-and-write-them-to-a-text-file.cs` | Extract all OfficeMath equations from a document and write them to a text file. | inspection-reporting | txt | llm_secondary |
| 24 | `save-the-modified-document-as-docx-while-preserving-all-officemath-equations-and-their-for.cs` | Save the modified document as DOCX while preserving all OfficeMath equations and their formatting. | native-officemath | docx | mcp |
| 25 | `export-the-document-containing-officemath-equations-to-pdf-and-verify-equation-layout-rema.cs` | Export the document containing OfficeMath equations to PDF and verify equation layout remains intact. | export-validation | docx, pdf | llm |
| 26 | `validate-that-exported-pdf-retains-the-exact-positioning-of-officemath-equations-as-in-the.cs` | Validate that exported PDF retains the exact positioning of OfficeMath equations as in the source DOCX. | export-validation | docx, pdf | llm |
| 27 | `batch-convert-a-collection-of-docx-files-containing-officemath-to-pdf-while-preserving-equ.cs` | Batch convert a collection of DOCX files containing OfficeMath to PDF while preserving equation fidelity. | export-validation | docx, pdf | existing_repo |
| 28 | `validate-that-each-officemath-object-conforms-to-expected-mathobjecttype-after-performing.cs` | Validate that each OfficeMath object conforms to expected MathObjectType after performing bulk updates. | native-officemath | docx | existing_repo |
| 29 | `create-a-macro-that-toggles-the-display-mode-of-selected-officemath-equations-between-inli.cs` | Create a macro that toggles the display mode of selected OfficeMath equations between inline and separate line. | formatting-display | docx | mcp |
| 30 | `iterate-over-all-officemath-nodes-in-a-document-to-count-the-total-number-of-equations.cs` | Iterate over all OfficeMath nodes in a document to count the total number of equations. | inspection-reporting | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Invented OfficeMath creation APIs

- Symptom: Compiler errors for unknown types such as InsertMathObject or LoadMathML.
- Root cause: Code attempts to use non-existent helper methods that are not part of the Aspose.Words API.
- Preferred fix: Replace the invented calls with the documented EQ-field bootstrap sequence.
- Prevention: Reference only the APIs listed in the official Aspose.Words.Math documentation.

### Invalid OfficeMath member usage

- Symptom: Compiler errors for members like OfficeMath.Start or IsInline.
- Root cause: Incorrect property or method names that do not exist on the OfficeMath class.
- Preferred fix: Use only the valid members DisplayType, Justification, and MathObjectType.
- Prevention: Consult the Aspose.Words.Math reference to verify member names before coding.

### Improper nested formatting

- Symptom: Runtime exceptions when setting DisplayType or Justification on nested OfficeMath nodes.
- Root cause: Attempting to modify display mode on equations that are not top-level OMathPara objects.
- Preferred fix: Restrict DisplayType and Justification changes to nodes where MathObjectType equals OMathPara.
- Prevention: Filter OfficeMath nodes by MathObjectType before applying formatting properties.

### Null OfficeMath insertion

- Symptom: Runtime failures such as "Cannot insert a null node" during replacement or insertion.
- Root cause: AsOfficeMath returned null because the EQ field failed to convert.
- Preferred fix: Verify that AsOfficeMath returns a non-null OfficeMath instance before insertion or replacement.
- Prevention: Add explicit null checks and throw informative errors if conversion fails.

### Missing sample assets

- Symptom: Runtime errors caused by attempts to load non-existent files or folders.
- Root cause: Examples assumed the presence of external DOCX files, streams, or directories.
- Preferred fix: Create all required sample documents, folders, and streams programmatically within the example.
- Prevention: Never rely on external assets; generate them locally at runtime.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Preserve the EQ-field bootstrap sequence unchanged and update sample EQ arguments only when new equation types are explicitly required.
