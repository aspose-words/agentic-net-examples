using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words, e.g., DOCX, PDF, RTF, etc.).
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Path where the MHTML file will be saved.
        string outputPath = @"C:\Docs\OutputDocument.mht";

        // Load the document. The constructor automatically detects the file format.
        Document doc = new Document(inputPath);

        // Save the document as MHTML using the overload that takes a SaveFormat enum.
        doc.Save(outputPath, SaveFormat.Mhtml);

        // If you need additional control over the conversion (e.g., embedding resources as CID URLs),
        // you can use HtmlSaveOptions instead:
        // HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Mhtml)
        // {
        //     ExportCidUrlsForMhtmlResources = true,
        //     PrettyFormat = true
        // };
        // doc.Save(outputPath, options);
    }
}
