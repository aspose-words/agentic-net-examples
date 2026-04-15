# OLE Objects Examples for Aspose.Words for .NET

This folder contains the **live, publish-ready** C# examples for the **OLE Objects** category. Each file is a standalone example selected from the latest verified publish-ready set and aligned with the active category rules.

## Snapshot

- Category: **OLE Objects**
- Slug: **ole-objects**
- Total examples: **30**
- Verified publish-ready examples: **30 / 30**
- Inspection workflow examples: **8**
- Extraction workflow examples: **7**
- Online Video workflow examples: **3**
- Ole Package workflow examples: **3**
- Validation workflow examples: **3**
- Insertion workflow examples: **3**
- Batch Processing workflow examples: **2**
- Replacement workflow examples: **1**


## Category rules that shaped these examples

- Always bootstrap local sample input files before embedding or inspecting OLE objects.
- Use `DocumentBuilder.InsertOleObject(...)` or documented APIs for insertion rather than speculative constructors or undocumented helpers.
- Inspect OLE objects through `Shape` nodes and `shape.OleFormat` only.
- Keep extraction and reporting workflows deterministic by writing outputs to local files and validating their existence.
- Use replacement by insert-then-remove at the correct parent level instead of mutating undocumented internals.

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
Copy-Item ..\ole-objects\<example-file>.cs .\Program.cs

# Build and run
dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

Replace `<example-file>` with a real file from `ole-objects/`.

## Running a single example with a real file name

Example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.3.0

# PowerShell example
Copy-Item ..\ole-objects\insert-a-spreadsheet-ole-object-into-a-docx-using-documentbuilder-insertoleobject-with-its.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `adjust-ole-object-icon-dimensions-to-maintain-aspect-ratio-after-resizing-for-document-lay.cs` | Adjust OLE object icon dimensions to maintain aspect ratio after resizing for document layout. | inspection | docx | publish-ready |
| 2 | `batch-insert-identical-excel-ole-object-into-multiple-word-files-using-a-foreach-loop.cs` | Batch insert identical Excel OLE object into multiple Word files using a foreach loop. | batch-processing | docx | publish-ready |
| 3 | `batch-insert-the-same-excel-ole-object-into-multiple-word-documents-within-a-folder-using.cs` | Batch insert the same Excel OLE object into multiple Word documents within a folder using a loop. | batch-processing | docx | publish-ready |
| 4 | `batch-process-a-collection-of-word-files-to-embed-a-common-online-video-into-each-document.cs` | Batch process a collection of Word files to embed a common online video into each document. | online-video | docx | publish-ready |
| 5 | `clone-an-ole-object-from-one-document-and-insert-it-into-another-using-insertoleobject-wit.cs` | Clone an OLE object from one document and insert it into another using InsertOleObject with extracted data. | extraction | docx | publish-ready |
| 6 | `create-a-utility-that-extracts-ole-objects-from-documents-and-stores-them-in-a-database-as.cs` | Create a utility that extracts OLE objects from documents and stores them in a database as BLOBs. | extraction | docx, bin | publish-ready |
| 7 | `embed-a-legacy-ole-package-when-the-handler-is-unknown-using-the-packager-approach.cs` | Embed a legacy OLE package when the handler is unknown using the Packager approach. | ole-package | docx | publish-ready |
| 8 | `export-all-ole-objects-from-a-document-into-separate-files-preserving-original-extensions.cs` | Export all OLE objects from a document into separate files preserving original extensions. | extraction | docx, bin | publish-ready |
| 9 | `extract-ole-object-metadata-such-as-source-file-name-and-size-and-write-to-a-csv-report.cs` | Extract OLE object metadata, such as source file name and size, and write to a CSV report. | extraction | docx, csv | publish-ready |
| 10 | `extract-the-ole-object-stream-from-a-document-and-save-it-to-a-binary-file.cs` | Extract the OLE object stream from a document and save it to a binary file. | extraction | docx, bin | publish-ready |
| 11 | `implement-error-handling-for-insertoleobject-when-the-specified-progid-is-not-registered-o.cs` | Implement error handling for InsertOleObject when the specified ProgId is not registered on the system. | validation | docx | publish-ready |
| 12 | `insert-a-pdf-file-as-an-ole-icon-specifying-custom-icon-image-and-display-size.cs` | Insert a PDF file as an OLE icon, specifying custom icon image and display size. | insertion | docx | publish-ready |
| 13 | `insert-a-spreadsheet-ole-object-into-a-docx-using-documentbuilder-insertoleobject-with-its.cs` | Insert a spreadsheet OLE object into a DOCX using DocumentBuilder.InsertOleObject with its ProgId. | inspection | docx | publish-ready |
| 14 | `insert-an-ole-object-as-an-icon-using-default-system-icon-without-specifying-custom-image.cs` | Insert an OLE object as an icon using default system icon without specifying custom image. | insertion | docx | publish-ready |
| 15 | `insert-an-online-video-with-specified-width-and-height-parameters-for-precise-placement.cs` | Insert an online video with specified width and height parameters for precise placement. | online-video | docx | publish-ready |
| 16 | `insert-an-online-youtube-video-into-a-word-document-using-documentbuilder-insertonlinevide.cs` | Insert an online YouTube video into a Word document using DocumentBuilder.InsertOnlineVideo. | online-video | docx | publish-ready |
| 17 | `iterate-through-all-ole-objects-in-a-docx-file-and-log-their-progids-and-display-sizes.cs` | Iterate through all OLE objects in a DOCX file and log their ProgIds and display sizes. | inspection | docx | publish-ready |
| 18 | `read-the-file-name-property-of-an-ole-package-and-compare-it-with-the-original-source-file.cs` | Read the file name property of an OLE package and compare it with the original source file. | inspection | docx | publish-ready |
| 19 | `replace-an-existing-ole-object-in-a-doc-file-with-a-new-image-ole-object-using-insertoleob.cs` | Replace an existing OLE object in a DOC file with a new image OLE object using InsertOleObject. | replacement | docx | publish-ready |
| 20 | `retrieve-display-width-and-height-of-an-ole-object-and-adjust-its-size-after-insertion.cs` | Retrieve display width and height of an OLE object and adjust its size after insertion. | inspection | docx | publish-ready |
| 21 | `retrieve-ole-object-display-width-and-height-after-insertion-and-store-dimensions-for-layo.cs` | Retrieve OLE object display width and height after insertion and store dimensions for layout calculations. | inspection | docx | publish-ready |
| 22 | `retrieve-raw-binary-data-from-an-ole-object-via-its-data-property-for-custom-processing.cs` | Retrieve raw binary data from an OLE object via its Data property for custom processing. | extraction | docx | publish-ready |
| 23 | `retrieve-the-progid-of-an-inserted-ole-object-and-log-it-for-diagnostic-purposes.cs` | Retrieve the ProgId of an inserted OLE object and log it for diagnostic purposes. | inspection | docx | publish-ready |
| 24 | `save-raw-binary-data-of-an-ole-object-to-a-temporary-file-for-external-analysis.cs` | Save raw binary data of an OLE object to a temporary file for external analysis. | extraction | docx, bin | publish-ready |
| 25 | `set-the-file-name-and-extension-for-an-inserted-ole-object-to-preserve-original-metadata.cs` | Set the file name and extension for an inserted OLE object to preserve original metadata. | inspection | docx | publish-ready |
| 26 | `use-documentbuilder-to-insert-an-ole-object-with-custom-width-and-height-parameters-for-pr.cs` | Use DocumentBuilder to insert an OLE object with custom width and height parameters for precise layout control. | insertion | docx | publish-ready |
| 27 | `use-olepackage-class-to-enumerate-all-parts-of-a-legacy-ole-package-for-inspection.cs` | Use OlePackage class to enumerate all parts of a legacy OLE package for inspection. | ole-package | docx | publish-ready |
| 28 | `use-olepackage-class-to-read-and-modify-properties-of-a-legacy-ole-package-in-net.cs` | Use OlePackage class to read and modify properties of a legacy OLE package in .NET. | ole-package | docx | publish-ready |
| 29 | `validate-the-progid-of-an-ole-object-before-insertion-to-avoid-runtime-errors.cs` | Validate the ProgId of an OLE object before insertion to avoid runtime errors. | validation | docx | publish-ready |
| 30 | `verify-successful-ole-object-insertion-by-checking-the-returned-object-reference-is-not-nu.cs` | Verify successful OLE object insertion by checking the returned object reference is not null. | validation | docx | publish-ready |


## Common failure patterns seen during generation and how they were corrected

### Missing local input asset

- Symptom: Runtime failure because a sample spreadsheet, PDF, DOCX, icon, or binary file does not exist.
- Fix: Create the input locally inside the example before loading or embedding it.

### Accessing OleFormat on a non-OLE shape

- Symptom: Null-reference failures when iterating shapes.
- Fix: Always check `shape.OleFormat != null` before reading metadata or raw data.

### Wrong replacement strategy

- Symptom: Invalid node manipulation or broken document structure during OLE replacement.
- Fix: Insert the replacement object at the correct parent level, then remove the old shape.

### Unstable external dependency assumption

- Symptom: Examples fail because they rely on software handlers, shared folders, or user-specific files.
- Fix: Keep examples self-contained and local to the run.

## Notes for maintainers

- The selected file for each task is the verified publish-ready file in the category folder.
- When updating this category, preserve local asset bootstrapping and `Shape`/`OleFormat`-based inspection conventions.
- Template-, icon-, and file-based samples should create temporary local assets for verification instead of assuming machine-specific paths.
