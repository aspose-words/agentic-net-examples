---
name: ole-objects
description: Verified C# examples for ole-objects scenarios in Aspose.Words for .NET
language: csharp
framework: net8.0
product: Aspose.Words
---

# AGENTS.md - OLE Objects

## Purpose

Provide an AI coding assistant with concise, authoritative guidance for implementing reliable OLE object workflows in Aspose.Words, ensuring correct API usage, robust error handling, and reproducible results across environments.

## Non-negotiable conventions

- Never call undocumented constructors or helpers for OLE insertion; only the documented InsertOleObject overloads may be used.
- Bootstrap local sample inputs for existing-file, folder, stream, image, and data workflows.
- Prefer simple, verifiable workflows over speculative formatting shortcuts.

## Recommended workflow selection

- **Inspection workflow**: 8 examples
- **Insertion workflow**: 3 examples
- **Online Video workflow**: 3 examples
- **Ole Package workflow**: 3 examples
- **Extraction workflow**: 7 examples
- **Replacement workflow**: 1 examples
- **Batch Processing workflow**: 2 examples
- **Validation workflow**: 3 examples

Use the simplest workflow that satisfies the task.

## Validation priorities

1. After each InsertOleObject call, confirm that the returned Shape is non-null and that its OleFormat.ProgId matches the expected value.

## File-to-task reference

- `insert-a-spreadsheet-ole-object-into-a-docx-using-documentbuilder-insertoleobject-with-its.cs`
  - Task: Insert a spreadsheet OLE object into a DOCX using DocumentBuilder.InsertOleObject with its ProgId.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: llm
- `insert-a-pdf-file-as-an-ole-icon-specifying-custom-icon-image-and-display-size.cs`
  - Task: Insert a PDF file as an OLE icon, specifying custom icon image and display size.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: llm
- `insert-an-ole-object-as-an-icon-using-default-system-icon-without-specifying-custom-image.cs`
  - Task: Insert an OLE object as an icon using default system icon without specifying custom image.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: mcp
- `set-the-file-name-and-extension-for-an-inserted-ole-object-to-preserve-original-metadata.cs`
  - Task: Set the file name and extension for an inserted OLE object to preserve original metadata.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: llm
- `use-documentbuilder-to-insert-an-ole-object-with-custom-width-and-height-parameters-for-pr.cs`
  - Task: Use DocumentBuilder to insert an OLE object with custom width and height parameters for precise layout control.
  - Workflow: insertion
  - Outputs: docx
  - Selected engine: existing_repo
- `insert-an-online-youtube-video-into-a-word-document-using-documentbuilder-insertonlinevide.cs`
  - Task: Insert an online YouTube video into a Word document using DocumentBuilder.InsertOnlineVideo.
  - Workflow: online-video
  - Outputs: docx
  - Selected engine: mcp
- `insert-an-online-video-with-specified-width-and-height-parameters-for-precise-placement.cs`
  - Task: Insert an online video with specified width and height parameters for precise placement.
  - Workflow: online-video
  - Outputs: docx
  - Selected engine: mcp
- `embed-a-legacy-ole-package-when-the-handler-is-unknown-using-the-packager-approach.cs`
  - Task: Embed a legacy OLE package when the handler is unknown using the Packager approach.
  - Workflow: ole-package
  - Outputs: docx
  - Selected engine: existing_repo
- `retrieve-the-progid-of-an-inserted-ole-object-and-log-it-for-diagnostic-purposes.cs`
  - Task: Retrieve the ProgId of an inserted OLE object and log it for diagnostic purposes.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-ole-object-display-width-and-height-after-insertion-and-store-dimensions-for-layo.cs`
  - Task: Retrieve OLE object display width and height after insertion and store dimensions for layout calculations.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: llm
- `adjust-ole-object-icon-dimensions-to-maintain-aspect-ratio-after-resizing-for-document-lay.cs`
  - Task: Adjust OLE object icon dimensions to maintain aspect ratio after resizing for document layout.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: llm
- `retrieve-display-width-and-height-of-an-ole-object-and-adjust-its-size-after-insertion.cs`
  - Task: Retrieve display width and height of an OLE object and adjust its size after insertion.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: mcp
- `retrieve-raw-binary-data-from-an-ole-object-via-its-data-property-for-custom-processing.cs`
  - Task: Retrieve raw binary data from an OLE object via its Data property for custom processing.
  - Workflow: extraction
  - Outputs: docx
  - Selected engine: llm
- `save-raw-binary-data-of-an-ole-object-to-a-temporary-file-for-external-analysis.cs`
  - Task: Save raw binary data of an OLE object to a temporary file for external analysis.
  - Workflow: extraction
  - Outputs: docx, bin
  - Selected engine: llm
- `use-olepackage-class-to-read-and-modify-properties-of-a-legacy-ole-package-in-net.cs`
  - Task: Use OlePackage class to read and modify properties of a legacy OLE package in .NET.
  - Workflow: ole-package
  - Outputs: docx
  - Selected engine: llm
- `use-olepackage-class-to-enumerate-all-parts-of-a-legacy-ole-package-for-inspection.cs`
  - Task: Use OlePackage class to enumerate all parts of a legacy OLE package for inspection.
  - Workflow: ole-package
  - Outputs: docx
  - Selected engine: llm
- `read-the-file-name-property-of-an-ole-package-and-compare-it-with-the-original-source-file.cs`
  - Task: Read the file name property of an OLE package and compare it with the original source file.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: mcp
- `extract-the-ole-object-stream-from-a-document-and-save-it-to-a-binary-file.cs`
  - Task: Extract the OLE object stream from a document and save it to a binary file.
  - Workflow: extraction
  - Outputs: docx, bin
  - Selected engine: llm
- `clone-an-ole-object-from-one-document-and-insert-it-into-another-using-insertoleobject-wit.cs`
  - Task: Clone an OLE object from one document and insert it into another using InsertOleObject with extracted data.
  - Workflow: extraction
  - Outputs: docx
  - Selected engine: existing_repo
- `replace-an-existing-ole-object-in-a-doc-file-with-a-new-image-ole-object-using-insertoleob.cs`
  - Task: Replace an existing OLE object in a DOC file with a new image OLE object using InsertOleObject.
  - Workflow: replacement
  - Outputs: docx
  - Selected engine: existing_repo
- `batch-insert-the-same-excel-ole-object-into-multiple-word-documents-within-a-folder-using.cs`
  - Task: Batch insert the same Excel OLE object into multiple Word documents within a folder using a loop.
  - Workflow: batch-processing
  - Outputs: docx
  - Selected engine: llm
- `batch-insert-identical-excel-ole-object-into-multiple-word-files-using-a-foreach-loop.cs`
  - Task: Batch insert identical Excel OLE object into multiple Word files using a foreach loop.
  - Workflow: batch-processing
  - Outputs: docx
  - Selected engine: llm
- `batch-process-a-collection-of-word-files-to-embed-a-common-online-video-into-each-document.cs`
  - Task: Batch process a collection of Word files to embed a common online video into each document.
  - Workflow: online-video
  - Outputs: docx
  - Selected engine: llm
- `validate-the-progid-of-an-ole-object-before-insertion-to-avoid-runtime-errors.cs`
  - Task: Validate the ProgId of an OLE object before insertion to avoid runtime errors.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: mcp
- `implement-error-handling-for-insertoleobject-when-the-specified-progid-is-not-registered-o.cs`
  - Task: Implement error handling for InsertOleObject when the specified ProgId is not registered on the system.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: llm
- `verify-successful-ole-object-insertion-by-checking-the-returned-object-reference-is-not-nu.cs`
  - Task: Verify successful OLE object insertion by checking the returned object reference is not null.
  - Workflow: validation
  - Outputs: docx
  - Selected engine: llm
- `extract-ole-object-metadata-such-as-source-file-name-and-size-and-write-to-a-csv-report.cs`
  - Task: Extract OLE object metadata, such as source file name and size, and write to a CSV report.
  - Workflow: extraction
  - Outputs: docx, csv
  - Selected engine: llm
- `create-a-utility-that-extracts-ole-objects-from-documents-and-stores-them-in-a-database-as.cs`
  - Task: Create a utility that extracts OLE objects from documents and stores them in a database as BLOBs.
  - Workflow: extraction
  - Outputs: docx, bin
  - Selected engine: llm
- `export-all-ole-objects-from-a-document-into-separate-files-preserving-original-extensions.cs`
  - Task: Export all OLE objects from a document into separate files preserving original extensions.
  - Workflow: extraction
  - Outputs: docx, bin
  - Selected engine: llm
- `iterate-through-all-ole-objects-in-a-docx-file-and-log-their-progids-and-display-sizes.cs`
  - Task: Iterate through all OLE objects in a DOCX file and log their ProgIds and display sizes.
  - Workflow: inspection
  - Outputs: docx
  - Selected engine: mcp

## Common failure patterns and preferred agent fixes

- **Incorrect InsertOleObject overload**
  - Symptom: Compilation fails with CS1501 or CS1503 because the argument list does not match any overload.
  - Root cause: The code supplies parameters in the wrong order or uses unsupported types such as string for a Stream argument.
  - Preferred fix: Use the overload signature InsertOleObject(Stream data, string progId, bool asIcon, Stream presentation) and pass null for the presentation stream when no custom icon is required.
  - Prevention: Consult the Aspose.Words API reference for the exact parameter types and order before writing insertion code.

- **Missing required type or namespace**
  - Symptom: Compilation fails with CS0246 indicating that types like OlePackage, IStorage, or System.Windows.Forms cannot be found.
  - Root cause: The example references legacy OLE APIs that are not part of the Aspose.Words assembly or the target .NET framework.
  - Preferred fix: Remove references to unsupported types and rely solely on Aspose.Words.Drawing.OlePackage and related classes that are available in the current package version.
  - Prevention: Limit code to the namespaces listed in the product documentation and verify package compatibility with the target framework.

- **Accessing non-existent OleFormat members**
  - Symptom: Compilation fails with CS1061 because properties such as FileName, Data, or OleData are not defined on OleFormat.
  - Root cause: The code assumes OleFormat exposes file-metadata or raw-data properties that are actually provided by OlePackage.
  - Preferred fix: Retrieve file name and binary data through shape.OleFormat.OlePackage.FileName and shape.OleFormat.OlePackage.Data respectively.
  - Prevention: Always navigate through OleFormat.OlePackage for package-level information and avoid direct member access on OleFormat.

- **Null icon stream when displaying as icon**
  - Symptom: Runtime throws ArgumentNullException for the parameter 'key' during InsertOleObject execution.
  - Root cause: The asIcon flag is true while the icon stream argument is null, causing the API to attempt to load a missing system resource.
  - Preferred fix: Pass a valid Stream containing an image when asIcon is true, or set asIcon to false to embed the object without an icon.
  - Prevention: Guard the InsertOleObject call with a conditional that supplies a non-null icon stream whenever asIcon is true.

- **Missing local input asset**
  - Symptom: Runtime fails because the sample file (e.g., spreadsheet, PDF, or image) does not exist at the expected path.
  - Root cause: Examples rely on external files that are not created or bundled with the code.
  - Preferred fix: Generate the required sample content programmatically (e.g., using MemoryStream) or write a temporary file before insertion.
  - Prevention: Include asset creation logic at the start of each example to guarantee that all inputs are present.

## Build and run contract

- Target framework: `net8.0`
- Primary package: `Aspose.Words` `26.11.0`

## Command reference

```bash
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0
```

```powershell
Copy-Item ..\ole-objects\<example-file>.cs .\Program.cs
```

```bash
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## Category update guidance

- Periodically verify that the InsertOleObject overload signatures have not changed in newer Aspose.Words releases and update the examples accordingly.
