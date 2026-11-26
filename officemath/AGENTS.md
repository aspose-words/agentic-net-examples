---
name: officemath
description: Verified C# examples for officemath scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - OfficeMath

## Purpose

An AI coding agent must generate code that creates, modifies, or removes OfficeMath objects by bootstrapping EQ fields, converting them with AsOfficeMath(), and applying only documented properties such as DisplayType, Justification, and MathObjectType.

## Non-negotiable conventions

- Do not call any API that is not part of Aspose.Words.Math, such as InsertMathObject, LoadMathML, or MathParser.
- Do not use undocumented helpers for equation creation, import, or conversion.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Inspection Reporting workflow**: 7 examples
- **Formatting Display workflow**: 11 examples
- **Creation Modification workflow**: 7 examples
- **Native Officemath workflow**: 2 examples
- **Export Validation workflow**: 3 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. After each modification, reload the document and confirm that the expected OfficeMath nodes exist and that their properties match the intended values.

## File-to-task reference

- `load-a-docx-document-containing-officemath-equations-and-enumerate-each-equation-node.cs`
  - Task: Load a DOCX document containing OfficeMath equations and enumerate each equation node.
  - Workflow: inspection-reporting
  - Outputs: docx, txt
  - Selected engine: mcp
- `load-multiple-docx-files-iterate-officemath-equations-and-standardize-justification-across.cs`
  - Task: Load multiple DOCX files, iterate OfficeMath equations, and standardize justification across all documents.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: existing_repo
- `create-a-new-officemath-object-with-documentbuilder-and-insert-it-as-an-inline-equation.cs`
  - Task: Create a new OfficeMath object with DocumentBuilder and insert it as an inline equation.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: llm
- `insert-an-officemath-equation-from-a-latex-string-representation-into-a-specific-document.cs`
  - Task: Insert an OfficeMath equation from a LaTeX string representation into a specific document location.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: mcp
- `insert-an-officemath-equation-from-a-mathml-string-into-a-paragraph-using-documentbuilder.cs`
  - Task: Insert an OfficeMath equation from a MathML string into a paragraph using DocumentBuilder.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: llm
- `create-a-batch-process-that-inserts-a-predefined-officemath-equation-into-each-paragraph-o.cs`
  - Task: Create a batch process that inserts a predefined OfficeMath equation into each paragraph of a document.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: mcp_secondary
- `replace-the-content-of-an-existing-officemath-object-with-a-new-equation-defined-by-a-stri.cs`
  - Task: Replace the content of an existing OfficeMath object with a new equation defined by a string.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: existing_repo
- `load-a-docx-file-replace-specific-officemath-equations-based-on-identifier-and-save-as-doc.cs`
  - Task: Load a DOCX file, replace specific OfficeMath equations based on identifier, and save as DOCX.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: mcp_secondary
- `delete-an-unwanted-officemath-node-from-the-document-and-adjust-surrounding-paragraph-spac.cs`
  - Task: Delete an unwanted OfficeMath node from the document and adjust surrounding paragraph spacing.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: mcp
- `clone-an-existing-officemath-object-and-insert-the-clone-at-a-new-location.cs`
  - Task: Clone an existing OfficeMath object and insert the clone at a new location.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: existing_repo
- `set-the-justification-of-an-officemath-equation-to-center-alignment-using-the-justificatio.cs`
  - Task: Set the justification of an OfficeMath equation to center alignment using the Justification property.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: mcp
- `update-the-justification-of-all-officemath-equations-to-right-alignment-in-a-template-docu.cs`
  - Task: Update the justification of all OfficeMath equations to right alignment in a template document.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: llm
- `apply-a-uniform-justification-setting-to-all-officemath-equations-across-multiple-sections.cs`
  - Task: Apply a uniform justification setting to all OfficeMath equations across multiple sections of a large document.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: existing_repo
- `validate-that-after-bulk-justification-changes-no-officemath-equation-exceeds-the-page-mar.cs`
  - Task: Validate that after bulk justification changes, no OfficeMath equation exceeds the page margin limits.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: existing_repo
- `configure-an-officemath-object-to-display-on-a-separate-line-rather-than-inline-within-the.cs`
  - Task: Configure an OfficeMath object to display on a separate line rather than inline within the paragraph.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: existing_repo
- `programmatically-change-all-officemath-equations-from-inline-to-separate-line-display-in-a.cs`
  - Task: Programmatically change all OfficeMath equations from inline to separate line display in a large report.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: existing_repo
- `replace-all-inline-officemath-equations-with-separate-line-display-to-enhance-visual-clari.cs`
  - Task: Replace all inline OfficeMath equations with separate line display to enhance visual clarity in presentations.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: existing_repo
- `programmatically-set-the-display-mode-of-officemath-equations-to-inline-for-compact-docume.cs`
  - Task: Programmatically set the display mode of OfficeMath equations to inline for compact document layout.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-the-mathobjecttype-of-each-officemath-node-to-determine-whether-it-is-a-fraction.cs`
  - Task: Retrieve the MathObjectType of each OfficeMath node to determine whether it is a fraction or radical.
  - Workflow: inspection-reporting
  - Outputs: docx
  - Selected engine: mcp_secondary
- `read-mathobjecttype-of-officemath-nodes-and-log-any-unsupported-equation-types-for-review.cs`
  - Task: Read MathObjectType of OfficeMath nodes and log any unsupported equation types for review.
  - Workflow: inspection-reporting
  - Outputs: docx
  - Selected engine: mcp
- `create-a-function-that-returns-true-if-an-officemath-node-matches-a-specified-mathobjectty.cs`
  - Task: Create a function that returns true if an OfficeMath node matches a specified MathObjectType criteria.
  - Workflow: inspection-reporting
  - Outputs: docx
  - Selected engine: llm
- `generate-a-report-listing-each-officemath-equation-s-mathobjecttype-and-its-position-withi.cs`
  - Task: Generate a report listing each OfficeMath equation's MathObjectType and its position within the document.
  - Workflow: inspection-reporting
  - Outputs: docx, txt
  - Selected engine: mcp
- `extract-all-officemath-equations-from-a-document-and-write-them-to-a-text-file.cs`
  - Task: Extract all OfficeMath equations from a document and write them to a text file.
  - Workflow: inspection-reporting
  - Outputs: txt
  - Selected engine: llm
- `save-the-modified-document-as-docx-while-preserving-all-officemath-equations-and-their-for.cs`
  - Task: Save the modified document as DOCX while preserving all OfficeMath equations and their formatting.
  - Workflow: native-officemath
  - Outputs: docx
  - Selected engine: mcp
- `export-the-document-containing-officemath-equations-to-pdf-and-verify-equation-layout-rema.cs`
  - Task: Export the document containing OfficeMath equations to PDF and verify equation layout remains intact.
  - Workflow: export-validation
  - Outputs: docx, pdf
  - Selected engine: mcp
- `validate-that-exported-pdf-retains-the-exact-positioning-of-officemath-equations-as-in-the.cs`
  - Task: Validate that exported PDF retains the exact positioning of OfficeMath equations as in the source DOCX.
  - Workflow: export-validation
  - Outputs: docx, pdf
  - Selected engine: mcp_secondary
- `batch-convert-a-collection-of-docx-files-containing-officemath-to-pdf-while-preserving-equ.cs`
  - Task: Batch convert a collection of DOCX files containing OfficeMath to PDF while preserving equation fidelity.
  - Workflow: export-validation
  - Outputs: docx, pdf
  - Selected engine: mcp
- `validate-that-each-officemath-object-conforms-to-expected-mathobjecttype-after-performing.cs`
  - Task: Validate that each OfficeMath object conforms to expected MathObjectType after performing bulk updates.
  - Workflow: native-officemath
  - Outputs: docx
  - Selected engine: llm
- `create-a-macro-that-toggles-the-display-mode-of-selected-officemath-equations-between-inli.cs`
  - Task: Create a macro that toggles the display mode of selected OfficeMath equations between inline and separate line.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: existing_repo
- `iterate-over-all-officemath-nodes-in-a-document-to-count-the-total-number-of-equations.cs`
  - Task: Iterate over all OfficeMath nodes in a document to count the total number of equations.
  - Workflow: inspection-reporting
  - Outputs: docx
  - Selected engine: llm

## Common failure patterns and preferred agent fixes

- **Invented OfficeMath APIs**
  - Symptom: Compiler errors such as CS1061 or CS0246 indicating missing members like InsertMathObject or LoadMathML.
  - Root cause: The code attempts to use APIs that are not defined in the Aspose.Words.Math namespace.
  - Preferred fix: Replace the nonexistent calls with the documented EQ-field bootstrap sequence and AsOfficeMath() conversion.
  - Prevention: Consult the official Aspose.Words.Math documentation and restrict code to the listed key APIs.

- **Invalid Member References**
  - Symptom: Compiler errors reporting undefined types or properties such as Justification, OfficeMathDisplayType, or MathObjectType.
  - Root cause: The code references members that either belong to a different namespace or do not exist in the current version of the library.
  - Preferred fix: Use the correct enum names OfficeMathJustification, OfficeMathDisplayType, and MathObjectType from Aspose.Words.Math.
  - Prevention: Reference the Aspose.Words.Math namespace explicitly and rely on IntelliSense to select valid members.

- **Improper Node Insertion**
  - Symptom: Runtime exceptions like InvalidOperationException or InvalidCastException when calling InsertBefore, InsertAfter, or Remove on unsupported node types.
  - Root cause: The code attempts to insert or replace OfficeMath nodes using generic Node methods that do not support the operation.
  - Preferred fix: Insert the OfficeMath node via the CompositeNode.InsertBefore method on the field's parent node and remove the original field with Field.Remove().
  - Prevention: Always obtain the CompositeNode parent of the field start and use its insertion methods; avoid calling InsertBefore on the abstract Node class.

- **Missing Sample Assets**
  - Symptom: FileNotFoundException or ArgumentNullException during sandbox verification because expected sample files or folders are absent.
  - Root cause: The example assumes the existence of input documents, folders, or streams that were not created at runtime.
  - Preferred fix: Create all required sample DOCX files, input directories, and temporary assets programmatically before loading or processing them.
  - Prevention: Embed sample document generation in the example code and verify file existence with explicit checks.

- **Incorrect Equation Counting**
  - Symptom: Validation failures reporting mismatched equation counts, such as expecting 7 equations but finding 0 or 20.
  - Root cause: The code counts all OfficeMath descendants instead of only top-level OMathPara nodes, leading to inflated or zero counts.
  - Preferred fix: Filter the NodeCollection with .Where(om => om.MathObjectType == MathObjectType.OMathPara) before counting.
  - Prevention: Always count equations based on MathObjectType.OMathPara to reflect logical top-level equations.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
dotnet add package Newtonsoft.Json
```

```powershell
Copy-Item ..\officemath\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- When Aspose.Words releases a new major version, review the EQ-to-OfficeMath conversion flow to ensure AsOfficeMath() still behaves as expected and update sample EQ arguments if necessary.
