# OLE Objects Examples for Aspose.Words for .NET

The OLE Objects category demonstrates how to embed, inspect, extract, and replace OLE content in Word documents using Aspose.Words 26.10.0. Verified examples cover spreadsheet insertion via InsertOleObject, PDF icon embedding with custom icons, legacy package creation, and dimension handling for precise layout.

## Snapshot

- Category: OLE Objects
- Slug: ole-objects
- Total examples: 29
- Publish-ready successful examples: 29 / 29
- Source run: 20261102_031946_08fb5b
- Inspection examples: 7
- Insertion examples: 3
- Online Video examples: 3
- Ole Package examples: 3
- Extraction examples: 7
- Replacement examples: 1
- Batch Processing examples: 2
- Validation examples: 3

## Category rules that shaped these examples

- Do not call undocumented constructors or helpers; only use the documented InsertOleObject overloads and Shape.OleFormat members.
- Create temporary input files in the example directory, insert them with InsertOleObject, and clean up after saving the document.
- Check that Shape.OleFormat is not null before accessing its properties to avoid null-reference warnings.

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
dotnet add package Aspose.Words --version 26.10.0

# PowerShell example
Copy-Item ..\ole-objects\insert-a-spreadsheet-ole-object-into-a-docx-using-documentbuilder-insertoleobject-with-its.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `insert-a-spreadsheet-ole-object-into-a-docx-using-documentbuilder-insertoleobject-with-its.cs` | Insert a spreadsheet OLE object into a DOCX using DocumentBuilder.InsertOleObject with its ProgId. | inspection | docx | existing_repo |
| 2 | `insert-a-pdf-file-as-an-ole-icon-specifying-custom-icon-image-and-display-size.cs` | Insert a PDF file as an OLE icon, specifying custom icon image and display size. | insertion | docx | llm |
| 3 | `insert-an-ole-object-as-an-icon-using-default-system-icon-without-specifying-custom-image.cs` | Insert an OLE object as an icon using default system icon without specifying custom image. | insertion | docx | llm |
| 4 | `set-the-file-name-and-extension-for-an-inserted-ole-object-to-preserve-original-metadata.cs` | Set the file name and extension for an inserted OLE object to preserve original metadata. | inspection | docx | llm |
| 5 | `use-documentbuilder-to-insert-an-ole-object-with-custom-width-and-height-parameters-for-pr.cs` | Use DocumentBuilder to insert an OLE object with custom width and height parameters for precise layout control. | insertion | docx | existing_repo |
| 6 | `insert-an-online-youtube-video-into-a-word-document-using-documentbuilder-insertonlinevide.cs` | Insert an online YouTube video into a Word document using DocumentBuilder.InsertOnlineVideo. | online-video | docx | mcp |
| 7 | `insert-an-online-video-with-specified-width-and-height-parameters-for-precise-placement.cs` | Insert an online video with specified width and height parameters for precise placement. | online-video | docx | llm |
| 8 | `embed-a-legacy-ole-package-when-the-handler-is-unknown-using-the-packager-approach.cs` | Embed a legacy OLE package when the handler is unknown using the Packager approach. | ole-package | docx | llm |
| 9 | `retrieve-the-progid-of-an-inserted-ole-object-and-log-it-for-diagnostic-purposes.cs` | Retrieve the ProgId of an inserted OLE object and log it for diagnostic purposes. | inspection | docx | existing_repo |
| 10 | `retrieve-ole-object-display-width-and-height-after-insertion-and-store-dimensions-for-layo.cs` | Retrieve OLE object display width and height after insertion and store dimensions for layout calculations. | inspection | docx | llm |
| 11 | `adjust-ole-object-icon-dimensions-to-maintain-aspect-ratio-after-resizing-for-document-lay.cs` | Adjust OLE object icon dimensions to maintain aspect ratio after resizing for document layout. | inspection | docx | llm |
| 12 | `retrieve-raw-binary-data-from-an-ole-object-via-its-data-property-for-custom-processing.cs` | Retrieve raw binary data from an OLE object via its Data property for custom processing. | extraction | docx | llm |
| 13 | `save-raw-binary-data-of-an-ole-object-to-a-temporary-file-for-external-analysis.cs` | Save raw binary data of an OLE object to a temporary file for external analysis. | extraction | docx, bin | llm |
| 14 | `use-olepackage-class-to-read-and-modify-properties-of-a-legacy-ole-package-in-net.cs` | Use OlePackage class to read and modify properties of a legacy OLE package in .NET. | ole-package | docx | llm |
| 15 | `use-olepackage-class-to-enumerate-all-parts-of-a-legacy-ole-package-for-inspection.cs` | Use OlePackage class to enumerate all parts of a legacy OLE package for inspection. | ole-package | docx | llm |
| 16 | `read-the-file-name-property-of-an-ole-package-and-compare-it-with-the-original-source-file.cs` | Read the file name property of an OLE package and compare it with the original source file. | inspection | docx | llm |
| 17 | `extract-the-ole-object-stream-from-a-document-and-save-it-to-a-binary-file.cs` | Extract the OLE object stream from a document and save it to a binary file. | extraction | docx, bin | llm |
| 18 | `clone-an-ole-object-from-one-document-and-insert-it-into-another-using-insertoleobject-wit.cs` | Clone an OLE object from one document and insert it into another using InsertOleObject with extracted data. | extraction | docx | existing_repo |
| 19 | `replace-an-existing-ole-object-in-a-doc-file-with-a-new-image-ole-object-using-insertoleob.cs` | Replace an existing OLE object in a DOC file with a new image OLE object using InsertOleObject. | replacement | docx | existing_repo |
| 20 | `batch-insert-the-same-excel-ole-object-into-multiple-word-documents-within-a-folder-using.cs` | Batch insert the same Excel OLE object into multiple Word documents within a folder using a loop. | batch-processing | docx | llm |
| 21 | `batch-insert-identical-excel-ole-object-into-multiple-word-files-using-a-foreach-loop.cs` | Batch insert identical Excel OLE object into multiple Word files using a foreach loop. | batch-processing | docx | llm |
| 22 | `batch-process-a-collection-of-word-files-to-embed-a-common-online-video-into-each-document.cs` | Batch process a collection of Word files to embed a common online video into each document. | online-video | docx | mcp |
| 23 | `validate-the-progid-of-an-ole-object-before-insertion-to-avoid-runtime-errors.cs` | Validate the ProgId of an OLE object before insertion to avoid runtime errors. | validation | docx | llm |
| 24 | `implement-error-handling-for-insertoleobject-when-the-specified-progid-is-not-registered-o.cs` | Implement error handling for InsertOleObject when the specified ProgId is not registered on the system. | validation | docx | llm |
| 25 | `verify-successful-ole-object-insertion-by-checking-the-returned-object-reference-is-not-nu.cs` | Verify successful OLE object insertion by checking the returned object reference is not null. | validation | docx | llm |
| 26 | `extract-ole-object-metadata-such-as-source-file-name-and-size-and-write-to-a-csv-report.cs` | Extract OLE object metadata, such as source file name and size, and write to a CSV report. | extraction | docx, csv | mcp |
| 27 | `create-a-utility-that-extracts-ole-objects-from-documents-and-stores-them-in-a-database-as.cs` | Create a utility that extracts OLE objects from documents and stores them in a database as BLOBs. | extraction | docx, bin | llm |
| 28 | `export-all-ole-objects-from-a-document-into-separate-files-preserving-original-extensions.cs` | Export all OLE objects from a document into separate files preserving original extensions. | extraction | docx, bin | llm |
| 29 | `iterate-through-all-ole-objects-in-a-docx-file-and-log-their-progids-and-display-sizes.cs` | Iterate through all OLE objects in a DOCX file and log their ProgIds and display sizes. | inspection | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Incorrect InsertOleObject overload

- Symptom: Compilation fails because no InsertOleObject method matches the supplied argument count.
- Root cause: The code supplies an overload signature that does not exist in Aspose.Words 26.10.0.
- Preferred fix: Use one of the supported overloads: InsertOleObject(Stream, string progId, bool asIcon, Stream? presentation) or the overload that accepts width and height via the returned Shape.
- Prevention: Consult the Aspose.Words API reference for the exact parameter list before coding.

### Mismatched argument types

- Symptom: CS1503 errors indicate that a string, bool, or int is being passed where a Stream or bool is expected.
- Root cause: Arguments are ordered incorrectly or the wrong variable type is supplied to InsertOleObject.
- Preferred fix: Pass a FileStream or MemoryStream for the data argument, a valid ProgId string, a boolean for asIcon, and a Stream for the optional presentation image.
- Prevention: Declare variables with the correct .NET types and match the overload signature exactly.

### COM type resolution failure

- Symptom: Runtime throws ArgumentNullException because Type.GetTypeFromProgID returns null.
- Root cause: The required COM application (e.g., Word) is not installed or not registered on the current machine.
- Preferred fix: Guard the call with a null check and provide a clear message; run the example only on Windows machines with the appropriate Office component installed.
- Prevention: Document the platform requirement and include a pre-flight check for the ProgID before invoking COM.

### Missing local input asset

- Symptom: Runtime error occurs when the sample spreadsheet, PDF, or icon file cannot be found.
- Root cause: The example assumes the presence of external files that are not created at runtime.
- Preferred fix: Generate minimal placeholder files (e.g., a tiny XLSX ZIP header or a simple PDF text) within the example before insertion.
- Prevention: Always create required input files programmatically at the start of the sample.

### Accessing OleFormat on a non-OLE shape

- Symptom: NullReferenceException is thrown while iterating document shapes.
- Root cause: The code attempts to read OleFormat from a Shape that does not contain OLE data.
- Preferred fix: Check shape.OleFormat != null before accessing its members.
- Prevention: Filter shapes by shape.HasOleObject or by verifying OleFormat is non-null in loops.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.10.0 run.
- Update the example code whenever Aspose.Words introduces new InsertOleObject overloads or deprecates existing ones.
