# OfficeMath Examples for Aspose.Words for .NET

This folder contains the live, publish-ready C# examples for the OfficeMath category. Each file is a standalone console example selected from the verified 26.5.0 run.

## Snapshot

- Category: OfficeMath
- Slug: officemath
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20260619_131835_59df5f
- Creation Modification examples: 7
- Export Validation examples: 3
- Formatting Display examples: 11
- Inspection Reporting examples: 7
- Native Officemath examples: 2

## Category rules that shaped these examples

- Always use native `Aspose.Words` OfficeMath APIs and real OfficeMath nodes.
- Always bootstrap local sample inputs when a scenario references an existing DOC, DOCX, stream, folder, or source document.
- Use documented OfficeMath properties such as `DisplayType`, `Justification`, and `MathObjectType` only.
- Do not invent unsupported equation creation, import, or conversion helpers.
- Validate OfficeMath-sensitive results through actual node traversal or saved output artifacts.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words 26.5.0

## Running Examples

Each file in this folder is a single, standalone `.cs` console example. To run one example:

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.5.0

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
dotnet add package Aspose.Words --version 26.5.0

# PowerShell example
Copy-Item ..\officemath\load-a-docx-document-containing-officemath-equations-and-enumerate-each-equation-node.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `load-a-docx-document-containing-officemath-equations-and-enumerate-each-equation-node.cs` | Load a DOCX document containing OfficeMath equations and enumerate each equation node. | Inspection Reporting | docx, txt | mcp |
| 2 | `load-multiple-docx-files-iterate-officemath-equations-and-standardize-justification-across.cs` | Load multiple DOCX files, iterate OfficeMath equations, and standardize justification across all documents. | Formatting Display | docx | mcp |
| 3 | `create-a-new-officemath-object-with-documentbuilder-and-insert-it-as-an-inline-equation.cs` | Create a new OfficeMath object with DocumentBuilder and insert it as an inline equation. | Formatting Display | docx | mcp |
| 4 | `insert-an-officemath-equation-from-a-latex-string-representation-into-a-specific-document.cs` | Insert an OfficeMath equation from a LaTeX string representation into a specific document location. | Creation Modification | docx | llm |
| 5 | `insert-an-officemath-equation-from-a-mathml-string-into-a-paragraph-using-documentbuilder.cs` | Insert an OfficeMath equation from a MathML string into a paragraph using DocumentBuilder. | Creation Modification | docx | mcp_secondary |
| 6 | `create-a-batch-process-that-inserts-a-predefined-officemath-equation-into-each-paragraph-o.cs` | Create a batch process that inserts a predefined OfficeMath equation into each paragraph of a document. | Creation Modification | docx | mcp |
| 7 | `replace-the-content-of-an-existing-officemath-object-with-a-new-equation-defined-by-a-stri.cs` | Replace the content of an existing OfficeMath object with a new equation defined by a string. | Creation Modification | docx | existing_repo |
| 8 | `load-a-docx-file-replace-specific-officemath-equations-based-on-identifier-and-save-as-doc.cs` | Load a DOCX file, replace specific OfficeMath equations based on identifier, and save as DOCX. | Creation Modification | docx | existing_repo |
| 9 | `delete-an-unwanted-officemath-node-from-the-document-and-adjust-surrounding-paragraph-spac.cs` | Delete an unwanted OfficeMath node from the document and adjust surrounding paragraph spacing. | Creation Modification | docx | existing_repo |
| 10 | `clone-an-existing-officemath-object-and-insert-the-clone-at-a-new-location.cs` | Clone an existing OfficeMath object and insert the clone at a new location. | Creation Modification | docx | llm |
| 11 | `set-the-justification-of-an-officemath-equation-to-center-alignment-using-the-justificatio.cs` | Set the justification of an OfficeMath equation to center alignment using the Justification property. | Formatting Display | docx | llm |
| 12 | `update-the-justification-of-all-officemath-equations-to-right-alignment-in-a-template-docu.cs` | Update the justification of all OfficeMath equations to right alignment in a template document. | Formatting Display | docx | mcp_secondary |
| 13 | `apply-a-uniform-justification-setting-to-all-officemath-equations-across-multiple-sections.cs` | Apply a uniform justification setting to all OfficeMath equations across multiple sections of a large document. | Formatting Display | docx | mcp_secondary |
| 14 | `validate-that-after-bulk-justification-changes-no-officemath-equation-exceeds-the-page-mar.cs` | Validate that after bulk justification changes, no OfficeMath equation exceeds the page margin limits. | Formatting Display | docx | mcp |
| 15 | `configure-an-officemath-object-to-display-on-a-separate-line-rather-than-inline-within-the.cs` | Configure an OfficeMath object to display on a separate line rather than inline within the paragraph. | Formatting Display | docx | mcp |
| 16 | `programmatically-change-all-officemath-equations-from-inline-to-separate-line-display-in-a.cs` | Programmatically change all OfficeMath equations from inline to separate line display in a large report. | Formatting Display | docx | mcp |
| 17 | `replace-all-inline-officemath-equations-with-separate-line-display-to-enhance-visual-clari.cs` | Replace all inline OfficeMath equations with separate line display to enhance visual clarity in presentations. | Formatting Display | docx | mcp |
| 18 | `programmatically-set-the-display-mode-of-officemath-equations-to-inline-for-compact-docume.cs` | Programmatically set the display mode of OfficeMath equations to inline for compact document layout. | Formatting Display | docx | mcp |
| 19 | `retrieve-the-mathobjecttype-of-each-officemath-node-to-determine-whether-it-is-a-fraction.cs` | Retrieve the MathObjectType of each OfficeMath node to determine whether it is a fraction or radical. | Inspection Reporting | docx | mcp |
| 20 | `read-mathobjecttype-of-officemath-nodes-and-log-any-unsupported-equation-types-for-review.cs` | Read MathObjectType of OfficeMath nodes and log any unsupported equation types for review. | Inspection Reporting | docx | mcp |
| 21 | `create-a-function-that-returns-true-if-an-officemath-node-matches-a-specified-mathobjectty.cs` | Create a function that returns true if an OfficeMath node matches a specified MathObjectType criteria. | Inspection Reporting | docx | mcp |
| 22 | `generate-a-report-listing-each-officemath-equation-s-mathobjecttype-and-its-position-withi.cs` | Generate a report listing each OfficeMath equation's MathObjectType and its position within the document. | Inspection Reporting | docx, txt | mcp |
| 23 | `extract-all-officemath-equations-from-a-document-and-write-them-to-a-text-file.cs` | Extract all OfficeMath equations from a document and write them to a text file. | Inspection Reporting | txt | mcp |
| 24 | `save-the-modified-document-as-docx-while-preserving-all-officemath-equations-and-their-for.cs` | Save the modified document as DOCX while preserving all OfficeMath equations and their formatting. | Native Officemath | docx | mcp |
| 25 | `export-the-document-containing-officemath-equations-to-pdf-and-verify-equation-layout-rema.cs` | Export the document containing OfficeMath equations to PDF and verify equation layout remains intact. | Export Validation | docx, pdf | llm |
| 26 | `validate-that-exported-pdf-retains-the-exact-positioning-of-officemath-equations-as-in-the.cs` | Validate that exported PDF retains the exact positioning of OfficeMath equations as in the source DOCX. | Export Validation | docx, pdf | existing_repo |
| 27 | `batch-convert-a-collection-of-docx-files-containing-officemath-to-pdf-while-preserving-equ.cs` | Batch convert a collection of DOCX files containing OfficeMath to PDF while preserving equation fidelity. | Export Validation | docx, pdf | mcp_secondary |
| 28 | `validate-that-each-officemath-object-conforms-to-expected-mathobjecttype-after-performing.cs` | Validate that each OfficeMath object conforms to expected MathObjectType after performing bulk updates. | Native Officemath | docx | existing_repo |
| 29 | `create-a-macro-that-toggles-the-display-mode-of-selected-officemath-equations-between-inli.cs` | Create a macro that toggles the display mode of selected OfficeMath equations between inline and separate line. | Formatting Display | docx | mcp |
| 30 | `iterate-over-all-officemath-nodes-in-a-document-to-count-the-total-number-of-equations.cs` | Iterate over all OfficeMath nodes in a document to count the total number of equations. | Inspection Reporting | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Invented OfficeMath creation APIs

- Symptom: Compiler errors around InsertMathObject, InsertEquation, StartMath, EndMath, MathParser, MathMLImporter, LoadMathML, LoadOoxml, or new OfficeMath(document).
- Fix: Use only documented Aspose.Words OfficeMath APIs and follow the verified bootstrap pattern already used in this category.

### Invalid OfficeMath type or property names

- Symptom: Compiler errors around unsupported members such as OfficeMath.Start, IsInline, OfficeMathObjectType, or OMathObjectType.
- Fix: Use real OfficeMath nodes and valid members such as DisplayType, Justification, and MathObjectType.

### Nested OfficeMath formatting mistakes

- Symptom: Runtime exceptions when changing DisplayType or Justification on nested OfficeMath nodes.
- Fix: Restrict display-mode and justification changes to top-level OfficeMath nodes whose MathObjectType is OMathPara.

### Null OfficeMath insertion or replacement

- Symptom: Runtime failures such as Cannot insert a null node or failed replacement after an invalid bootstrap attempt.
- Fix: Null-check created or imported nodes before insertion and only replace valid top-level OfficeMath targets.

### Verifier missing input asset or path

- Symptom: Runtime failures caused by missing sample files, folders, or streams during sandbox verification.
- Fix: Create temporary local assets inside the example before loading or processing them.

### Brittle logical equation counts

- Symptom: Validation failures caused by counting all nested OfficeMath descendants instead of logical top-level equations.
- Fix: Validate against the exact intended top-level nodes for formatting, replacement, cloning, and display-mode workflows.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.5.0 run.
- Preserve file-to-task traceability when updating this folder.
- Keep examples standalone and bootstrap local inputs inside the example whenever external sources are mentioned.
