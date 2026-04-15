---
name: ole-objects
description: Verified C# examples for OLE object scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - OLE Objects

## Purpose

This folder is a **live, curated example set** for OLE object scenarios. Treat every file as a standalone console example. The main goal is correct, warning-free insertion, inspection, extraction, replacement, and validation of embedded or linked OLE content in Word-centric workflows.

## Non-negotiable conventions

- Always use `Shape` nodes with `OleFormat` for OLE object discovery and inspection.
- Always bootstrap local sample input files before embedding them.
- Never assume external files, folders, handlers, or machine-specific assets already exist.
- Use `DocumentBuilder.InsertOleObject(...)` or documented online-video helpers when the task is about insertion.
- For extraction workflows, always validate `shape.OleFormat != null` before reading metadata or raw data.
- For replacement workflows, prefer inserting a new valid OLE object and removing the old shape rather than mutating undocumented internals.

## Recommended workflow selection

- **inspection workflow**: 8 examples
- **extraction workflow**: 7 examples
- **online-video workflow**: 3 examples
- **ole-package workflow**: 3 examples
- **validation workflow**: 3 examples
- **insertion workflow**: 3 examples
- **batch-processing workflow**: 2 examples
- **replacement workflow**: 1 example

Choose the simplest workflow that satisfies the task. Default to insertion for new embeddings, inspection for metadata/reporting scenarios, extraction for raw-data export, replacement for swap/clone scenarios, and batch-processing for folder-scale work.

## Validation priorities

1. The code must compile and run without manual input.
2. Any expected output document must be saved locally.
3. File-based examples must create their own local input assets before use.
4. Extraction/reporting examples must verify the output file exists.
5. OLE inspection examples must guard against shapes that do not contain `OleFormat`.

## File-to-task reference

- `adjust-ole-object-icon-dimensions-to-maintain-aspect-ratio-after-resizing-for-document-lay.cs`
  - Task: Adjust OLE object icon dimensions to maintain aspect ratio after resizing for document layout.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: publish-ready
- `batch-insert-identical-excel-ole-object-into-multiple-word-files-using-a-foreach-loop.cs`
  - Task: Batch insert identical Excel OLE object into multiple Word files using a foreach loop.
  - Workflow: batch-processing
  - Outputs: docx
  - Selected engine: publish-ready
- `batch-insert-the-same-excel-ole-object-into-multiple-word-documents-within-a-folder-using.cs`
  - Task: Batch insert the same Excel OLE object into multiple Word documents within a folder using a loop.
  - Workflow: batch-processing
  - Outputs: docx
  - Selected engine: publish-ready
- `batch-process-a-collection-of-word-files-to-embed-a-common-online-video-into-each-document.cs`
  - Task: Batch process a collection of Word files to embed a common online video into each document.
  - Workflow: online-video
  - Outputs: docx
  - Selected engine: publish-ready
- `clone-an-ole-object-from-one-document-and-insert-it-into-another-using-insertoleobject-wit.cs`
  - Task: Clone an OLE object from one document and insert it into another using InsertOleObject with extracted data.
  - Workflow: extraction
  - Outputs: docx
  - Selected engine: publish-ready
- `create-a-utility-that-extracts-ole-objects-from-documents-and-stores-them-in-a-database-as.cs`
  - Task: Create a utility that extracts OLE objects from documents and stores them in a database as BLOBs.
  - Workflow: extraction
  - Outputs: docx, bin
  - Selected engine: publish-ready
- `embed-a-legacy-ole-package-when-the-handler-is-unknown-using-the-packager-approach.cs`
  - Task: Embed a legacy OLE package when the handler is unknown using the Packager approach.
  - Workflow: ole-package
  - Outputs: docx
  - Selected engine: publish-ready
- `export-all-ole-objects-from-a-document-into-separate-files-preserving-original-extensions.cs`
  - Task: Export all OLE objects from a document into separate files preserving original extensions.
  - Workflow: extraction
  - Outputs: docx, bin
  - Selected engine: publish-ready
- `extract-ole-object-metadata-such-as-source-file-name-and-size-and-write-to-a-csv-report.cs`
  - Task: Extract OLE object metadata, such as source file name and size, and write to a CSV report.
  - Workflow: extraction
  - Outputs: docx, csv
  - Selected engine: publish-ready
- `extract-the-ole-object-stream-from-a-document-and-save-it-to-a-binary-file.cs`
  - Task: Extract the OLE object stream from a document and save it to a binary file.
  - Workflow: extraction
  - Outputs: docx, bin
  - Selected engine: publish-ready
- `implement-error-handling-for-insertoleobject-when-the-specified-progid-is-not-registered-o.cs`
  - Task: Implement error handling for InsertOleObject when the specified ProgId is not registered on the system.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: publish-ready
- `insert-a-pdf-file-as-an-ole-icon-specifying-custom-icon-image-and-display-size.cs`
  - Task: Insert a PDF file as an OLE icon, specifying custom icon image and display size.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: publish-ready
- `insert-a-spreadsheet-ole-object-into-a-docx-using-documentbuilder-insertoleobject-with-its.cs`
  - Task: Insert a spreadsheet OLE object into a DOCX using DocumentBuilder.InsertOleObject with its ProgId.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: publish-ready
- `insert-an-ole-object-as-an-icon-using-default-system-icon-without-specifying-custom-image.cs`
  - Task: Insert an OLE object as an icon using default system icon without specifying custom image.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: publish-ready
- `insert-an-online-video-with-specified-width-and-height-parameters-for-precise-placement.cs`
  - Task: Insert an online video with specified width and height parameters for precise placement.
  - Workflow: online-video
  - Outputs: docx
  - Selected engine: publish-ready
- `insert-an-online-youtube-video-into-a-word-document-using-documentbuilder-insertonlinevide.cs`
  - Task: Insert an online YouTube video into a Word document using DocumentBuilder.InsertOnlineVideo.
  - Workflow: online-video
  - Outputs: docx
  - Selected engine: publish-ready
- `iterate-through-all-ole-objects-in-a-docx-file-and-log-their-progids-and-display-sizes.cs`
  - Task: Iterate through all OLE objects in a DOCX file and log their ProgIds and display sizes.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: publish-ready
- `read-the-file-name-property-of-an-ole-package-and-compare-it-with-the-original-source-file.cs`
  - Task: Read the file name property of an OLE package and compare it with the original source file.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: publish-ready
- `replace-an-existing-ole-object-in-a-doc-file-with-a-new-image-ole-object-using-insertoleob.cs`
  - Task: Replace an existing OLE object in a DOC file with a new image OLE object using InsertOleObject.
  - Workflow: replacement
  - Outputs: docx
  - Selected engine: publish-ready
- `retrieve-display-width-and-height-of-an-ole-object-and-adjust-its-size-after-insertion.cs`
  - Task: Retrieve display width and height of an OLE object and adjust its size after insertion.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: publish-ready
- `retrieve-ole-object-display-width-and-height-after-insertion-and-store-dimensions-for-layo.cs`
  - Task: Retrieve OLE object display width and height after insertion and store dimensions for layout calculations.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: publish-ready
- `retrieve-raw-binary-data-from-an-ole-object-via-its-data-property-for-custom-processing.cs`
  - Task: Retrieve raw binary data from an OLE object via its Data property for custom processing.
  - Workflow: extraction
  - Outputs: docx
  - Selected engine: publish-ready
- `retrieve-the-progid-of-an-inserted-ole-object-and-log-it-for-diagnostic-purposes.cs`
  - Task: Retrieve the ProgId of an inserted OLE object and log it for diagnostic purposes.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: publish-ready
- `save-raw-binary-data-of-an-ole-object-to-a-temporary-file-for-external-analysis.cs`
  - Task: Save raw binary data of an OLE object to a temporary file for external analysis.
  - Workflow: extraction
  - Outputs: docx, bin
  - Selected engine: publish-ready
- `set-the-file-name-and-extension-for-an-inserted-ole-object-to-preserve-original-metadata.cs`
  - Task: Set the file name and extension for an inserted OLE object to preserve original metadata.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: publish-ready
- `use-documentbuilder-to-insert-an-ole-object-with-custom-width-and-height-parameters-for-pr.cs`
  - Task: Use DocumentBuilder to insert an OLE object with custom width and height parameters for precise layout control.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: publish-ready
- `use-olepackage-class-to-enumerate-all-parts-of-a-legacy-ole-package-for-inspection.cs`
  - Task: Use OlePackage class to enumerate all parts of a legacy OLE package for inspection.
  - Workflow: ole-package
  - Outputs: docx
  - Selected engine: publish-ready
- `use-olepackage-class-to-read-and-modify-properties-of-a-legacy-ole-package-in-net.cs`
  - Task: Use OlePackage class to read and modify properties of a legacy OLE package in .NET.
  - Workflow: ole-package
  - Outputs: docx
  - Selected engine: publish-ready
- `validate-the-progid-of-an-ole-object-before-insertion-to-avoid-runtime-errors.cs`
  - Task: Validate the ProgId of an OLE object before insertion to avoid runtime errors.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: publish-ready
- `verify-successful-ole-object-insertion-by-checking-the-returned-object-reference-is-not-nu.cs`
  - Task: Verify successful OLE object insertion by checking the returned object reference is not null.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: publish-ready

## Common failure patterns and preferred agent fixes

- **Missing local input asset**
  - Symptom: Runtime failure because a sample spreadsheet, PDF, DOCX, icon, or binary file does not exist.
  - Preferred fix: Create the input locally inside the example before loading or embedding it.
- **Accessing OleFormat on a non-OLE shape**
  - Symptom: Null-reference failures when iterating shapes.
  - Preferred fix: Always check `shape.OleFormat != null` before reading metadata or data.
- **Wrong replacement strategy**
  - Symptom: Invalid node manipulation or broken document structure during OLE replacement.
  - Preferred fix: Insert the replacement object at the correct parent level, then remove the old shape.
- **Unstable external dependency assumption**
  - Symptom: Examples fail because they rely on software handlers, shared folders, or user-specific files.
  - Preferred fix: Keep examples self-contained and local to the run.

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
Copy-Item ..\ole-objects\<example-file>.cs .\Program.cs
```

### Build and run

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Preserve file-to-task traceability. Any future update should keep the original CSV task text associated with the file in metadata.
- When replacing a file, prefer the verified publish-ready version rather than a merely compiling draft.
- Keep examples self-contained by generating all required local assets at runtime.
