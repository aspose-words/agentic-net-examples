# OfficeMath Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **OfficeMath** category. Each file is a standalone example selected from the verified publish-ready set and aligned with the active category rules.

## Snapshot

- Category: **OfficeMath**
- Slug: **officemath**
- Total examples: **30**
- Formatting and display examples: **11 / 30**
- Export and validation examples: **3 / 30**
- Creation and modification examples: **7 / 30**
- Inspection and reporting examples: **7 / 30**
- Native OfficeMath utility examples: **2 / 30**

## Category rules that shaped these examples

- Always use native `Aspose.Words` OfficeMath APIs and real OfficeMath nodes.
- Always bootstrap local sample inputs when a scenario references an existing DOC, DOCX, stream, folder, or source document.
- Use documented OfficeMath properties such as `DisplayType`, `Justification`, and `MathObjectType` only.
- Do not invent unsupported equation creation, import, or conversion helpers.
- Validate OfficeMath-sensitive results through actual node traversal or saved output artifacts.

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
Copy-Item ..\officemath\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `officemath/`.

## Running a single example with a real file name

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# PowerShell example
Copy-Item ..\officemath\apply-a-uniform-justification-setting-to-all-officemath-equations-across-multiple-sections.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `apply-a-uniform-justification-setting-to-all-officemath-equations-across-multiple-sections.cs` | Apply a uniform justification setting to all OfficeMath equations across multiple sections of a large document. | formatting-display | docx | verified |
| 2 | `batch-convert-a-collection-of-docx-files-containing-officemath-to-pdf-while-preserving-equ.cs` | Batch convert a collection of DOCX files containing OfficeMath to PDF while preserving equation fidelity. | export-validation | docx, pdf | verified |
| 3 | `clone-an-existing-officemath-object-and-insert-the-clone-at-a-new-location.cs` | Clone an existing OfficeMath object and insert the clone at a new location. | creation-modification | docx | verified |
| 4 | `configure-an-officemath-object-to-display-on-a-separate-line-rather-than-inline-within-the.cs` | Configure an OfficeMath object to display on a separate line rather than inline within the paragraph. | formatting-display | docx | verified |
| 5 | `create-a-batch-process-that-inserts-a-predefined-officemath-equation-into-each-paragraph-o.cs` | Create a batch process that inserts a predefined OfficeMath equation into each paragraph of a document. | creation-modification | docx | verified |
| 6 | `create-a-function-that-returns-true-if-an-officemath-node-matches-a-specified-mathobjectty.cs` | Create a function that returns true if an OfficeMath node matches a specified MathObjectType criteria. | inspection-reporting | docx | verified |
| 7 | `create-a-macro-that-toggles-the-display-mode-of-selected-officemath-equations-between-inli.cs` | Create a macro that toggles the display mode of selected OfficeMath equations between inline and separate line. | formatting-display | docx | verified |
| 8 | `create-a-new-officemath-object-with-documentbuilder-and-insert-it-as-an-inline-equation.cs` | Create a new OfficeMath object with DocumentBuilder and insert it as an inline equation. | formatting-display | docx | verified |
| 9 | `delete-an-unwanted-officemath-node-from-the-document-and-adjust-surrounding-paragraph-spac.cs` | Delete an unwanted OfficeMath node from the document and adjust surrounding paragraph spacing. | creation-modification | docx | verified |
| 10 | `export-the-document-containing-officemath-equations-to-pdf-and-verify-equation-layout-rema.cs` | Export the document containing OfficeMath equations to PDF and verify equation layout remains intact. | export-validation | docx, pdf | verified |
| 11 | `extract-all-officemath-equations-from-a-document-and-write-them-to-a-text-file.cs` | Extract all OfficeMath equations from a document and write them to a text file. | inspection-reporting | txt | verified |
| 12 | `generate-a-report-listing-each-officemath-equation-s-mathobjecttype-and-its-position-withi.cs` | Generate a report listing each OfficeMath equation's MathObjectType and its position within the document. | inspection-reporting | docx, txt | verified |
| 13 | `insert-an-officemath-equation-from-a-latex-string-representation-into-a-specific-document.cs` | Insert an OfficeMath equation from a LaTeX string representation into a specific document location. | creation-modification | docx | verified |
| 14 | `insert-an-officemath-equation-from-a-mathml-string-into-a-paragraph-using-documentbuilder.cs` | Insert an OfficeMath equation from a MathML string into a paragraph using DocumentBuilder. | creation-modification | docx | verified |
| 15 | `iterate-over-all-officemath-nodes-in-a-document-to-count-the-total-number-of-equations.cs` | Iterate over all OfficeMath nodes in a document to count the total number of equations. | inspection-reporting | docx | verified |
| 16 | `load-a-docx-document-containing-officemath-equations-and-enumerate-each-equation-node.cs` | Load a DOCX document containing OfficeMath equations and enumerate each equation node. | inspection-reporting | docx, txt | verified |
| 17 | `load-a-docx-file-replace-specific-officemath-equations-based-on-identifier-and-save-as-doc.cs` | Load a DOCX file, replace specific OfficeMath equations based on identifier, and save as DOCX. | creation-modification | docx | verified |
| 18 | `load-multiple-docx-files-iterate-officemath-equations-and-standardize-justification-across.cs` | Load multiple DOCX files, iterate OfficeMath equations, and standardize justification across all documents. | formatting-display | docx | verified |
| 19 | `programmatically-change-all-officemath-equations-from-inline-to-separate-line-display-in-a.cs` | Programmatically change all OfficeMath equations from inline to separate line display in a large report. | formatting-display | docx | verified |
| 20 | `programmatically-set-the-display-mode-of-officemath-equations-to-inline-for-compact-docume.cs` | Programmatically set the display mode of OfficeMath equations to inline for compact document layout. | formatting-display | docx | verified |
| 21 | `read-mathobjecttype-of-officemath-nodes-and-log-any-unsupported-equation-types-for-review.cs` | Read MathObjectType of OfficeMath nodes and log any unsupported equation types for review. | inspection-reporting | docx | verified |
| 22 | `replace-all-inline-officemath-equations-with-separate-line-display-to-enhance-visual-clari.cs` | Replace all inline OfficeMath equations with separate line display to enhance visual clarity in presentations. | formatting-display | docx | verified |
| 23 | `replace-the-content-of-an-existing-officemath-object-with-a-new-equation-defined-by-a-stri.cs` | Replace the content of an existing OfficeMath object with a new equation defined by a string. | creation-modification | docx | verified |
| 24 | `retrieve-the-mathobjecttype-of-each-officemath-node-to-determine-whether-it-is-a-fraction.cs` | Retrieve the MathObjectType of each OfficeMath node to determine whether it is a fraction or radical. | inspection-reporting | docx | verified |
| 25 | `save-the-modified-document-as-docx-while-preserving-all-officemath-equations-and-their-for.cs` | Save the modified document as DOCX while preserving all OfficeMath equations and their formatting. | native-officemath | docx | verified |
| 26 | `set-the-justification-of-an-officemath-equation-to-center-alignment-using-the-justificatio.cs` | Set the justification of an OfficeMath equation to center alignment using the Justification property. | formatting-display | docx | verified |
| 27 | `update-the-justification-of-all-officemath-equations-to-right-alignment-in-a-template-docu.cs` | Update the justification of all OfficeMath equations to right alignment in a template document. | formatting-display | docx | verified |
| 28 | `validate-that-after-bulk-justification-changes-no-officemath-equation-exceeds-the-page-mar.cs` | Validate that after bulk justification changes, no OfficeMath equation exceeds the page margin limits. | formatting-display | docx | verified |
| 29 | `validate-that-each-officemath-object-conforms-to-expected-mathobjecttype-after-performing.cs` | Validate that each OfficeMath object conforms to expected MathObjectType after performing bulk updates. | native-officemath | docx | verified |
| 30 | `validate-that-exported-pdf-retains-the-exact-positioning-of-officemath-equations-as-in-the.cs` | Validate that exported PDF retains the exact positioning of OfficeMath equations as in the source DOCX. | export-validation | docx, pdf | verified |

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

## Notes for maintainers

- The selected file for each task is the verified winner recorded in the batch report.
- When updating this category, preserve the current OfficeMath conventions around real node traversal, local input bootstrap, and documented property usage.
- File-based, folder-based, and stream-based samples should create temporary local assets for verification instead of assuming machine-specific paths.
