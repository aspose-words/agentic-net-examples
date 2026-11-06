---
name: officemath
description: Verified C# examples for officemath scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - OfficeMath

## Purpose

Supply concise, reproducible code snippets that generate realistic DOCX inputs, convert EQ fields to genuine OfficeMath nodes, and perform supported operations such as enumeration, justification, insertion, cloning, and replacement while conforming to Aspose.Words 26.10.0 APIs.

## Non-negotiable conventions

- Do not call any API that is not part of Aspose.Words.Math, such as InsertMathObject, MathParser, or LoadMathML.
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

1. Validate that each generated document contains at least one top-level OfficeMath node and that its Justification matches the expected value.

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
  - Selected engine: mcp
- `insert-an-officemath-equation-from-a-latex-string-representation-into-a-specific-document.cs`
  - Task: Insert an OfficeMath equation from a LaTeX string representation into a specific document location.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: mcp
- `insert-an-officemath-equation-from-a-mathml-string-into-a-paragraph-using-documentbuilder.cs`
  - Task: Insert an OfficeMath equation from a MathML string into a paragraph using DocumentBuilder.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: mcp
- `create-a-batch-process-that-inserts-a-predefined-officemath-equation-into-each-paragraph-o.cs`
  - Task: Create a batch process that inserts a predefined OfficeMath equation into each paragraph of a document.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: mcp
- `replace-the-content-of-an-existing-officemath-object-with-a-new-equation-defined-by-a-stri.cs`
  - Task: Replace the content of an existing OfficeMath object with a new equation defined by a string.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: llm
- `load-a-docx-file-replace-specific-officemath-equations-based-on-identifier-and-save-as-doc.cs`
  - Task: Load a DOCX file, replace specific OfficeMath equations based on identifier, and save as DOCX.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: llm
- `delete-an-unwanted-officemath-node-from-the-document-and-adjust-surrounding-paragraph-spac.cs`
  - Task: Delete an unwanted OfficeMath node from the document and adjust surrounding paragraph spacing.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: mcp
- `clone-an-existing-officemath-object-and-insert-the-clone-at-a-new-location.cs`
  - Task: Clone an existing OfficeMath object and insert the clone at a new location.
  - Workflow: creation-modification
  - Outputs: docx
  - Selected engine: mcp
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
  - Selected engine: llm_secondary
- `programmatically-change-all-officemath-equations-from-inline-to-separate-line-display-in-a.cs`
  - Task: Programmatically change all OfficeMath equations from inline to separate line display in a large report.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: llm
- `replace-all-inline-officemath-equations-with-separate-line-display-to-enhance-visual-clari.cs`
  - Task: Replace all inline OfficeMath equations with separate line display to enhance visual clarity in presentations.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: llm_secondary
- `programmatically-set-the-display-mode-of-officemath-equations-to-inline-for-compact-docume.cs`
  - Task: Programmatically set the display mode of OfficeMath equations to inline for compact document layout.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-the-mathobjecttype-of-each-officemath-node-to-determine-whether-it-is-a-fraction.cs`
  - Task: Retrieve the MathObjectType of each OfficeMath node to determine whether it is a fraction or radical.
  - Workflow: inspection-reporting
  - Outputs: docx
  - Selected engine: mcp
- `read-mathobjecttype-of-officemath-nodes-and-log-any-unsupported-equation-types-for-review.cs`
  - Task: Read MathObjectType of OfficeMath nodes and log any unsupported equation types for review.
  - Workflow: inspection-reporting
  - Outputs: docx
  - Selected engine: existing_repo
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
  - Selected engine: llm_secondary
- `save-the-modified-document-as-docx-while-preserving-all-officemath-equations-and-their-for.cs`
  - Task: Save the modified document as DOCX while preserving all OfficeMath equations and their formatting.
  - Workflow: native-officemath
  - Outputs: docx
  - Selected engine: mcp
- `export-the-document-containing-officemath-equations-to-pdf-and-verify-equation-layout-rema.cs`
  - Task: Export the document containing OfficeMath equations to PDF and verify equation layout remains intact.
  - Workflow: export-validation
  - Outputs: docx, pdf
  - Selected engine: llm
- `validate-that-exported-pdf-retains-the-exact-positioning-of-officemath-equations-as-in-the.cs`
  - Task: Validate that exported PDF retains the exact positioning of OfficeMath equations as in the source DOCX.
  - Workflow: export-validation
  - Outputs: docx, pdf
  - Selected engine: llm
- `batch-convert-a-collection-of-docx-files-containing-officemath-to-pdf-while-preserving-equ.cs`
  - Task: Batch convert a collection of DOCX files containing OfficeMath to PDF while preserving equation fidelity.
  - Workflow: export-validation
  - Outputs: docx, pdf
  - Selected engine: existing_repo
- `validate-that-each-officemath-object-conforms-to-expected-mathobjecttype-after-performing.cs`
  - Task: Validate that each OfficeMath object conforms to expected MathObjectType after performing bulk updates.
  - Workflow: native-officemath
  - Outputs: docx
  - Selected engine: existing_repo
- `create-a-macro-that-toggles-the-display-mode-of-selected-officemath-equations-between-inli.cs`
  - Task: Create a macro that toggles the display mode of selected OfficeMath equations between inline and separate line.
  - Workflow: formatting-display
  - Outputs: docx
  - Selected engine: mcp
- `iterate-over-all-officemath-nodes-in-a-document-to-count-the-total-number-of-equations.cs`
  - Task: Iterate over all OfficeMath nodes in a document to count the total number of equations.
  - Workflow: inspection-reporting
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Invented OfficeMath creation APIs**
  - Symptom: Compiler errors for unknown types such as InsertMathObject or LoadMathML.
  - Root cause: Code attempts to use non-existent helper methods that are not part of the Aspose.Words API.
  - Preferred fix: Replace the invented calls with the documented EQ-field bootstrap sequence.
  - Prevention: Reference only the APIs listed in the official Aspose.Words.Math documentation.

- **Invalid OfficeMath member usage**
  - Symptom: Compiler errors for members like OfficeMath.Start or IsInline.
  - Root cause: Incorrect property or method names that do not exist on the OfficeMath class.
  - Preferred fix: Use only the valid members DisplayType, Justification, and MathObjectType.
  - Prevention: Consult the Aspose.Words.Math reference to verify member names before coding.

- **Improper nested formatting**
  - Symptom: Runtime exceptions when setting DisplayType or Justification on nested OfficeMath nodes.
  - Root cause: Attempting to modify display mode on equations that are not top-level OMathPara objects.
  - Preferred fix: Restrict DisplayType and Justification changes to nodes where MathObjectType equals OMathPara.
  - Prevention: Filter OfficeMath nodes by MathObjectType before applying formatting properties.

- **Null OfficeMath insertion**
  - Symptom: Runtime failures such as "Cannot insert a null node" during replacement or insertion.
  - Root cause: AsOfficeMath returned null because the EQ field failed to convert.
  - Preferred fix: Verify that AsOfficeMath returns a non-null OfficeMath instance before insertion or replacement.
  - Prevention: Add explicit null checks and throw informative errors if conversion fails.

- **Missing sample assets**
  - Symptom: Runtime errors caused by attempts to load non-existent files or folders.
  - Root cause: Examples assumed the presence of external DOCX files, streams, or directories.
  - Preferred fix: Create all required sample documents, folders, and streams programmatically within the example.
  - Prevention: Never rely on external assets; generate them locally at runtime.

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
Copy-Item ..\officemath\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve the EQ-field bootstrap sequence unchanged and update sample EQ arguments only when new equation types are explicitly required.
