using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourceDocPath = @"C:\Docs\Sample.doc";

        // Path to the converted output file (PDF in this example).
        string outputPdfPath = @"C:\Docs\Sample_Converted.pdf";

        // Load the DOC document. The LoadOptions explicitly specify the DOC format,
        // which speeds up loading by avoiding auto‑detection.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Doc, "", "");

        // Use the Document constructor that accepts a file name and LoadOptions.
        Document doc = new Document(sourceDocPath, loadOptions);

        // Save the loaded document to the desired format.
        doc.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
