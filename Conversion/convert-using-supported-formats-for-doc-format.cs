using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words;

// Load any supported document (PDF in this example) using default LoadOptions.
string inputFile = @"C:\Docs\input.pdf";
LoadOptions loadOptions = new LoadOptions();               // default options
Document document = new Document(inputFile, loadOptions);

// (Optional) Detect the original format – useful for validation or logging.
FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputFile);
// Example: Console.WriteLine($"Original format: {formatInfo.LoadFormat}");

// Save the loaded document in the legacy DOC format.
// DocSaveOptions allows additional DOC‑specific settings; here we use the constructor that sets SaveFormat.
string outputFile = @"C:\Docs\output.doc";
DocSaveOptions docSaveOptions = new DocSaveOptions(SaveFormat.Doc);
document.Save(outputFile, docSaveOptions);
