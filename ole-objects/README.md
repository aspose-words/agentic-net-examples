# OLE Objects Examples for Aspose.Words for .NET

The OLE Objects category covers embedding, inspecting, extracting, and replacing OLE content within Word documents using Aspose.Words. Verified examples demonstrate correct insertion of Excel and PDF OLE objects, icon handling, custom sizing, legacy package embedding, and metadata retrieval.

## Snapshot

- Category: OLE Objects
- Slug: ole-objects
- Total examples: 30
- Publish-ready successful examples: 30 / 30
- Source run: 20261124_190740_1c00f5
- Inspection examples: 8
- Insertion examples: 3
- Online Video examples: 3
- Ole Package examples: 3
- Extraction examples: 7
- Replacement examples: 1
- Batch Processing examples: 2
- Validation examples: 3

## Category rules that shaped these examples

- Never call undocumented constructors or helpers for OLE insertion; only the documented InsertOleObject overloads may be used.
- Bootstrap all input assets within the example code and reference them via streams rather than file paths to avoid missing-file errors.
- Check that any optional icon stream passed to InsertOleObject is non-null when asIcon is true, and verify Shape.OleFormat is not null before accessing its members.

## Prerequisites

- .NET SDK 8.0 or later
- Aspose.Words for .NET `26.11.0`

## Running Examples

Each file in this folder is a **single, standalone `.cs` console example**. To run one example:

```bash
# From the repository root
dotnet new console -n ExampleProject --framework net8.0
cd ExampleProject
dotnet add package Aspose.Words --version 26.11.0

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
dotnet add package Aspose.Words --version 26.11.0

# PowerShell example
Copy-Item ..\ole-objects\insert-a-spreadsheet-ole-object-into-a-docx-using-documentbuilder-insertoleobject-with-its.cs .\Program.cs

dotnet build --configuration Release --verbosity minimal
dotnet run --configuration Release --no-build
```

## File-to-task map

| # | File | Original task | Workflow | Outputs | Selected engine |
|---|---|---|---|---|---|
| 1 | `insert-a-spreadsheet-ole-object-into-a-docx-using-documentbuilder-insertoleobject-with-its.cs` | Insert a spreadsheet OLE object into a DOCX using DocumentBuilder.InsertOleObject with its ProgId. | inspection | docx | llm |
| 2 | `insert-a-pdf-file-as-an-ole-icon-specifying-custom-icon-image-and-display-size.cs` | Insert a PDF file as an OLE icon, specifying custom icon image and display size. | insertion | docx | llm |
| 3 | `insert-an-ole-object-as-an-icon-using-default-system-icon-without-specifying-custom-image.cs` | Insert an OLE object as an icon using default system icon without specifying custom image. | insertion | docx | mcp |
| 4 | `set-the-file-name-and-extension-for-an-inserted-ole-object-to-preserve-original-metadata.cs` | Set the file name and extension for an inserted OLE object to preserve original metadata. | inspection | docx | llm |
| 5 | `use-documentbuilder-to-insert-an-ole-object-with-custom-width-and-height-parameters-for-pr.cs` | Use DocumentBuilder to insert an OLE object with custom width and height parameters for precise layout control. | insertion | docx | existing_repo |
| 6 | `insert-an-online-youtube-video-into-a-word-document-using-documentbuilder-insertonlinevide.cs` | Insert an online YouTube video into a Word document using DocumentBuilder.InsertOnlineVideo. | online-video | docx | mcp |
| 7 | `insert-an-online-video-with-specified-width-and-height-parameters-for-precise-placement.cs` | Insert an online video with specified width and height parameters for precise placement. | online-video | docx | mcp |
| 8 | `embed-a-legacy-ole-package-when-the-handler-is-unknown-using-the-packager-approach.cs` | Embed a legacy OLE package when the handler is unknown using the Packager approach. | ole-package | docx | existing_repo |
| 9 | `retrieve-the-progid-of-an-inserted-ole-object-and-log-it-for-diagnostic-purposes.cs` | Retrieve the ProgId of an inserted OLE object and log it for diagnostic purposes. | inspection | docx | mcp |
| 10 | `retrieve-ole-object-display-width-and-height-after-insertion-and-store-dimensions-for-layo.cs` | Retrieve OLE object display width and height after insertion and store dimensions for layout calculations. | inspection | docx | llm |
| 11 | `adjust-ole-object-icon-dimensions-to-maintain-aspect-ratio-after-resizing-for-document-lay.cs` | Adjust OLE object icon dimensions to maintain aspect ratio after resizing for document layout. | inspection | docx | llm |
| 12 | `retrieve-display-width-and-height-of-an-ole-object-and-adjust-its-size-after-insertion.cs` | Retrieve display width and height of an OLE object and adjust its size after insertion. | inspection | docx | mcp |
| 13 | `retrieve-raw-binary-data-from-an-ole-object-via-its-data-property-for-custom-processing.cs` | Retrieve raw binary data from an OLE object via its Data property for custom processing. | extraction | docx | llm |
| 14 | `save-raw-binary-data-of-an-ole-object-to-a-temporary-file-for-external-analysis.cs` | Save raw binary data of an OLE object to a temporary file for external analysis. | extraction | docx, bin | llm |
| 15 | `use-olepackage-class-to-read-and-modify-properties-of-a-legacy-ole-package-in-net.cs` | Use OlePackage class to read and modify properties of a legacy OLE package in .NET. | ole-package | docx | llm |
| 16 | `use-olepackage-class-to-enumerate-all-parts-of-a-legacy-ole-package-for-inspection.cs` | Use OlePackage class to enumerate all parts of a legacy OLE package for inspection. | ole-package | docx | llm |
| 17 | `read-the-file-name-property-of-an-ole-package-and-compare-it-with-the-original-source-file.cs` | Read the file name property of an OLE package and compare it with the original source file. | inspection | docx | mcp |
| 18 | `extract-the-ole-object-stream-from-a-document-and-save-it-to-a-binary-file.cs` | Extract the OLE object stream from a document and save it to a binary file. | extraction | docx, bin | llm |
| 19 | `clone-an-ole-object-from-one-document-and-insert-it-into-another-using-insertoleobject-wit.cs` | Clone an OLE object from one document and insert it into another using InsertOleObject with extracted data. | extraction | docx | existing_repo |
| 20 | `replace-an-existing-ole-object-in-a-doc-file-with-a-new-image-ole-object-using-insertoleob.cs` | Replace an existing OLE object in a DOC file with a new image OLE object using InsertOleObject. | replacement | docx | existing_repo |
| 21 | `batch-insert-the-same-excel-ole-object-into-multiple-word-documents-within-a-folder-using.cs` | Batch insert the same Excel OLE object into multiple Word documents within a folder using a loop. | batch-processing | docx | llm |
| 22 | `batch-insert-identical-excel-ole-object-into-multiple-word-files-using-a-foreach-loop.cs` | Batch insert identical Excel OLE object into multiple Word files using a foreach loop. | batch-processing | docx | llm |
| 23 | `batch-process-a-collection-of-word-files-to-embed-a-common-online-video-into-each-document.cs` | Batch process a collection of Word files to embed a common online video into each document. | online-video | docx | llm |
| 24 | `validate-the-progid-of-an-ole-object-before-insertion-to-avoid-runtime-errors.cs` | Validate the ProgId of an OLE object before insertion to avoid runtime errors. | validation | docx | mcp |
| 25 | `implement-error-handling-for-insertoleobject-when-the-specified-progid-is-not-registered-o.cs` | Implement error handling for InsertOleObject when the specified ProgId is not registered on the system. | validation | docx | llm |
| 26 | `verify-successful-ole-object-insertion-by-checking-the-returned-object-reference-is-not-nu.cs` | Verify successful OLE object insertion by checking the returned object reference is not null. | validation | docx | llm |
| 27 | `extract-ole-object-metadata-such-as-source-file-name-and-size-and-write-to-a-csv-report.cs` | Extract OLE object metadata, such as source file name and size, and write to a CSV report. | extraction | docx, csv | llm |
| 28 | `create-a-utility-that-extracts-ole-objects-from-documents-and-stores-them-in-a-database-as.cs` | Create a utility that extracts OLE objects from documents and stores them in a database as BLOBs. | extraction | docx, bin | llm |
| 29 | `export-all-ole-objects-from-a-document-into-separate-files-preserving-original-extensions.cs` | Export all OLE objects from a document into separate files preserving original extensions. | extraction | docx, bin | llm |
| 30 | `iterate-through-all-ole-objects-in-a-docx-file-and-log-their-progids-and-display-sizes.cs` | Iterate through all OLE objects in a DOCX file and log their ProgIds and display sizes. | inspection | docx | mcp |

## Common failure patterns seen during generation and how they were corrected

### Incorrect InsertOleObject overload

- Symptom: Compilation fails with CS1501 or CS1503 because the argument list does not match any overload.
- Root cause: The code supplies parameters in the wrong order or uses unsupported types such as string for a Stream argument.
- Preferred fix: Use the overload signature InsertOleObject(Stream data, string progId, bool asIcon, Stream presentation) and pass null for the presentation stream when no custom icon is required.
- Prevention: Consult the Aspose.Words API reference for the exact parameter types and order before writing insertion code.

### Missing required type or namespace

- Symptom: Compilation fails with CS0246 indicating that types like OlePackage, IStorage, or System.Windows.Forms cannot be found.
- Root cause: The example references legacy OLE APIs that are not part of the Aspose.Words assembly or the target .NET framework.
- Preferred fix: Remove references to unsupported types and rely solely on Aspose.Words.Drawing.OlePackage and related classes that are available in the current package version.
- Prevention: Limit code to the namespaces listed in the product documentation and verify package compatibility with the target framework.

### Accessing non-existent OleFormat members

- Symptom: Compilation fails with CS1061 because properties such as FileName, Data, or OleData are not defined on OleFormat.
- Root cause: The code assumes OleFormat exposes file-metadata or raw-data properties that are actually provided by OlePackage.
- Preferred fix: Retrieve file name and binary data through shape.OleFormat.OlePackage.FileName and shape.OleFormat.OlePackage.Data respectively.
- Prevention: Always navigate through OleFormat.OlePackage for package-level information and avoid direct member access on OleFormat.

### Null icon stream when displaying as icon

- Symptom: Runtime throws ArgumentNullException for the parameter 'key' during InsertOleObject execution.
- Root cause: The asIcon flag is true while the icon stream argument is null, causing the API to attempt to load a missing system resource.
- Preferred fix: Pass a valid Stream containing an image when asIcon is true, or set asIcon to false to embed the object without an icon.
- Prevention: Guard the InsertOleObject call with a conditional that supplies a non-null icon stream whenever asIcon is true.

### Missing local input asset

- Symptom: Runtime fails because the sample file (e.g., spreadsheet, PDF, or image) does not exist at the expected path.
- Root cause: Examples rely on external files that are not created or bundled with the code.
- Preferred fix: Generate the required sample content programmatically (e.g., using MemoryStream) or write a temporary file before insertion.
- Prevention: Include asset creation logic at the start of each example to guarantee that all inputs are present.

## See Also

- [`AGENTS.md`](./AGENTS.md) -- category-specific anti-patterns, API surface, and conventions for AI coding agents
- [`../AGENTS.md`](../AGENTS.md) -- repository-wide agent guide
- [`../README.md`](../README.md) -- full category index and project overview
- [Aspose.Words for .NET docs](https://docs.aspose.com/words/net/)

> Each `.cs` file is a standalone, build-validated console example. Drop into a fresh `dotnet new console` project, add the `Aspose.Words` NuGet version listed above, and run.

## Notes for maintainers

- This category is 100% publish-ready for the 26.11.0 run.
- Periodically verify that the InsertOleObject overload signatures have not changed in newer Aspose.Words releases and update the examples accordingly.
