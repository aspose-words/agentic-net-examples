---
name: officemath
description: Verified C# examples for OfficeMath scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - OfficeMath

## Purpose

This folder is a live, curated example set for OfficeMath scenarios. Treat every `.cs` file as a standalone console application. The goal is correct, warning-free examples that use documented Aspose APIs and match the original task intent.

## Non-negotiable conventions

- Always use native `Aspose.Words` OfficeMath APIs directly.
- Always work with real OfficeMath nodes discovered through `NodeType.OfficeMath`.
- Always use documented OfficeMath properties such as `DisplayType`, `Justification`, and `MathObjectType`.
- Always create local sample DOC or DOCX inputs inside the example when a task references an existing file, stream, folder, or source document.
- Never invent unsupported helpers such as `InsertMathObject`, `InsertEquation`, `StartMath`, `EndMath`, `MathParser`, `MathMLImporter`, `OfficeMath.LoadMathML`, `OfficeMath.LoadOoxml`, or `new OfficeMath(document)`.
- If a task requires real OfficeMath creation, use the category-safe bootstrap and verification patterns required by the active rules for this category.

## Recommended workflow selection

- Creation Modification workflow: 7 examples
- Export Validation workflow: 3 examples
- Formatting Display workflow: 11 examples
- Inspection Reporting workflow: 7 examples
- Native Officemath workflow: 2 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. The code must compile and run without manual input.
2. API usage must be supported by the configured package versions.
3. Exported outputs must actually be written by the example.
4. Validation scenarios must inspect only the behavior requested by the task.
5. Examples that depend on files, folders, streams, images, or data should bootstrap those inputs locally during the example run.

## File-to-task reference

- `load-a-docx-document-containing-officemath-equations-and-enumerate-each-equation-node.cs`
  - Task: Load a DOCX document containing OfficeMath equations and enumerate each equation node.
  - Workflow: Inspection Reporting
  - Outputs: docx, txt
  - Selected engine: mcp
- `load-multiple-docx-files-iterate-officemath-equations-and-standardize-justification-across.cs`
  - Task: Load multiple DOCX files, iterate OfficeMath equations, and standardize justification across all documents.
  - Workflow: Formatting Display
  - Outputs: docx
  - Selected engine: mcp
- `create-a-new-officemath-object-with-documentbuilder-and-insert-it-as-an-inline-equation.cs`
  - Task: Create a new OfficeMath object with DocumentBuilder and insert it as an inline equation.
  - Workflow: Formatting Display
  - Outputs: docx
  - Selected engine: mcp
- `insert-an-officemath-equation-from-a-latex-string-representation-into-a-specific-document.cs`
  - Task: Insert an OfficeMath equation from a LaTeX string representation into a specific document location.
  - Workflow: Creation Modification
  - Outputs: docx
  - Selected engine: llm
- `insert-an-officemath-equation-from-a-mathml-string-into-a-paragraph-using-documentbuilder.cs`
  - Task: Insert an OfficeMath equation from a MathML string into a paragraph using DocumentBuilder.
  - Workflow: Creation Modification
  - Outputs: docx
  - Selected engine: mcp_secondary
- `create-a-batch-process-that-inserts-a-predefined-officemath-equation-into-each-paragraph-o.cs`
  - Task: Create a batch process that inserts a predefined OfficeMath equation into each paragraph of a document.
  - Workflow: Creation Modification
  - Outputs: docx
  - Selected engine: mcp
- `replace-the-content-of-an-existing-officemath-object-with-a-new-equation-defined-by-a-stri.cs`
  - Task: Replace the content of an existing OfficeMath object with a new equation defined by a string.
  - Workflow: Creation Modification
  - Outputs: docx
  - Selected engine: existing_repo
- `load-a-docx-file-replace-specific-officemath-equations-based-on-identifier-and-save-as-doc.cs`
  - Task: Load a DOCX file, replace specific OfficeMath equations based on identifier, and save as DOCX.
  - Workflow: Creation Modification
  - Outputs: docx
  - Selected engine: existing_repo
- `delete-an-unwanted-officemath-node-from-the-document-and-adjust-surrounding-paragraph-spac.cs`
  - Task: Delete an unwanted OfficeMath node from the document and adjust surrounding paragraph spacing.
  - Workflow: Creation Modification
  - Outputs: docx
  - Selected engine: existing_repo
- `clone-an-existing-officemath-object-and-insert-the-clone-at-a-new-location.cs`
  - Task: Clone an existing OfficeMath object and insert the clone at a new location.
  - Workflow: Creation Modification
  - Outputs: docx
  - Selected engine: llm
- `set-the-justification-of-an-officemath-equation-to-center-alignment-using-the-justificatio.cs`
  - Task: Set the justification of an OfficeMath equation to center alignment using the Justification property.
  - Workflow: Formatting Display
  - Outputs: docx
  - Selected engine: llm
- `update-the-justification-of-all-officemath-equations-to-right-alignment-in-a-template-docu.cs`
  - Task: Update the justification of all OfficeMath equations to right alignment in a template document.
  - Workflow: Formatting Display
  - Outputs: docx
  - Selected engine: mcp_secondary
- `apply-a-uniform-justification-setting-to-all-officemath-equations-across-multiple-sections.cs`
  - Task: Apply a uniform justification setting to all OfficeMath equations across multiple sections of a large document.
  - Workflow: Formatting Display
  - Outputs: docx
  - Selected engine: mcp_secondary
- `validate-that-after-bulk-justification-changes-no-officemath-equation-exceeds-the-page-mar.cs`
  - Task: Validate that after bulk justification changes, no OfficeMath equation exceeds the page margin limits.
  - Workflow: Formatting Display
  - Outputs: docx
  - Selected engine: mcp
- `configure-an-officemath-object-to-display-on-a-separate-line-rather-than-inline-within-the.cs`
  - Task: Configure an OfficeMath object to display on a separate line rather than inline within the paragraph.
  - Workflow: Formatting Display
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-change-all-officemath-equations-from-inline-to-separate-line-display-in-a.cs`
  - Task: Programmatically change all OfficeMath equations from inline to separate line display in a large report.
  - Workflow: Formatting Display
  - Outputs: docx
  - Selected engine: mcp
- `replace-all-inline-officemath-equations-with-separate-line-display-to-enhance-visual-clari.cs`
  - Task: Replace all inline OfficeMath equations with separate line display to enhance visual clarity in presentations.
  - Workflow: Formatting Display
  - Outputs: docx
  - Selected engine: mcp
- `programmatically-set-the-display-mode-of-officemath-equations-to-inline-for-compact-docume.cs`
  - Task: Programmatically set the display mode of OfficeMath equations to inline for compact document layout.
  - Workflow: Formatting Display
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-the-mathobjecttype-of-each-officemath-node-to-determine-whether-it-is-a-fraction.cs`
  - Task: Retrieve the MathObjectType of each OfficeMath node to determine whether it is a fraction or radical.
  - Workflow: Inspection Reporting
  - Outputs: docx
  - Selected engine: mcp
- `read-mathobjecttype-of-officemath-nodes-and-log-any-unsupported-equation-types-for-review.cs`
  - Task: Read MathObjectType of OfficeMath nodes and log any unsupported equation types for review.
  - Workflow: Inspection Reporting
  - Outputs: docx
  - Selected engine: mcp
- `create-a-function-that-returns-true-if-an-officemath-node-matches-a-specified-mathobjectty.cs`
  - Task: Create a function that returns true if an OfficeMath node matches a specified MathObjectType criteria.
  - Workflow: Inspection Reporting
  - Outputs: docx
  - Selected engine: mcp
- `generate-a-report-listing-each-officemath-equation-s-mathobjecttype-and-its-position-withi.cs`
  - Task: Generate a report listing each OfficeMath equation's MathObjectType and its position within the document.
  - Workflow: Inspection Reporting
  - Outputs: docx, txt
  - Selected engine: mcp
- `extract-all-officemath-equations-from-a-document-and-write-them-to-a-text-file.cs`
  - Task: Extract all OfficeMath equations from a document and write them to a text file.
  - Workflow: Inspection Reporting
  - Outputs: txt
  - Selected engine: mcp
- `save-the-modified-document-as-docx-while-preserving-all-officemath-equations-and-their-for.cs`
  - Task: Save the modified document as DOCX while preserving all OfficeMath equations and their formatting.
  - Workflow: Native Officemath
  - Outputs: docx
  - Selected engine: mcp
- `export-the-document-containing-officemath-equations-to-pdf-and-verify-equation-layout-rema.cs`
  - Task: Export the document containing OfficeMath equations to PDF and verify equation layout remains intact.
  - Workflow: Export Validation
  - Outputs: docx, pdf
  - Selected engine: llm
- `validate-that-exported-pdf-retains-the-exact-positioning-of-officemath-equations-as-in-the.cs`
  - Task: Validate that exported PDF retains the exact positioning of OfficeMath equations as in the source DOCX.
  - Workflow: Export Validation
  - Outputs: docx, pdf
  - Selected engine: existing_repo
- `batch-convert-a-collection-of-docx-files-containing-officemath-to-pdf-while-preserving-equ.cs`
  - Task: Batch convert a collection of DOCX files containing OfficeMath to PDF while preserving equation fidelity.
  - Workflow: Export Validation
  - Outputs: docx, pdf
  - Selected engine: mcp_secondary
- `validate-that-each-officemath-object-conforms-to-expected-mathobjecttype-after-performing.cs`
  - Task: Validate that each OfficeMath object conforms to expected MathObjectType after performing bulk updates.
  - Workflow: Native Officemath
  - Outputs: docx
  - Selected engine: existing_repo
- `create-a-macro-that-toggles-the-display-mode-of-selected-officemath-equations-between-inli.cs`
  - Task: Create a macro that toggles the display mode of selected OfficeMath equations between inline and separate line.
  - Workflow: Formatting Display
  - Outputs: docx
  - Selected engine: mcp
- `iterate-over-all-officemath-nodes-in-a-document-to-count-the-total-number-of-equations.cs`
  - Task: Iterate over all OfficeMath nodes in a document to count the total number of equations.
  - Workflow: Inspection Reporting
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- Invented OfficeMath creation APIs
  - Symptom: Compiler errors around InsertMathObject, InsertEquation, StartMath, EndMath, MathParser, MathMLImporter, LoadMathML, LoadOoxml, or new OfficeMath(document).
  - Preferred fix: Use only documented Aspose.Words OfficeMath APIs and follow the verified bootstrap pattern already used in this category.

- Invalid OfficeMath type or property names
  - Symptom: Compiler errors around unsupported members such as OfficeMath.Start, IsInline, OfficeMathObjectType, or OMathObjectType.
  - Preferred fix: Use real OfficeMath nodes and valid members such as DisplayType, Justification, and MathObjectType.

- Nested OfficeMath formatting mistakes
  - Symptom: Runtime exceptions when changing DisplayType or Justification on nested OfficeMath nodes.
  - Preferred fix: Restrict display-mode and justification changes to top-level OfficeMath nodes whose MathObjectType is OMathPara.

- Null OfficeMath insertion or replacement
  - Symptom: Runtime failures such as Cannot insert a null node or failed replacement after an invalid bootstrap attempt.
  - Preferred fix: Null-check created or imported nodes before insertion and only replace valid top-level OfficeMath targets.

- Verifier missing input asset or path
  - Symptom: Runtime failures caused by missing sample files, folders, or streams during sandbox verification.
  - Preferred fix: Create temporary local assets inside the example before loading or processing them.

- Brittle logical equation counts
  - Symptom: Validation failures caused by counting all nested OfficeMath descendants instead of logical top-level equations.
  - Preferred fix: Validate against the exact intended top-level nodes for formatting, replacement, cloning, and display-mode workflows.

## Build and run contract

- Target framework: `net8.0`
- Package: `Aspose.Words` `26.5.0`

## Command reference

### Create a temporary console project

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
```

### Add required packages

```bash
dotnet add package Aspose.Words --version 26.5.0
```

### Copy a category example into the temp project

```powershell
Copy-Item ..\officemath\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original task text associated with the file in metadata.
- When replacing a file, prefer the latest verified publish-ready version rather than a merely compiling draft.
- Keep examples standalone, bootstrap missing inputs locally, and prefer documented Aspose APIs over speculative shortcuts.
