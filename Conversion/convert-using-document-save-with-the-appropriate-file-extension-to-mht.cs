using System;
using Aspose.Words;

class ConvertDocmToMhtml
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path to the destination MHTML file.
        string outputPath = @"C:\Docs\ConvertedDocument.mhtml";

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // Save the document in MHTML format using the explicit SaveFormat.
        doc.Save(outputPath, SaveFormat.Mhtml);

        // Alternatively, you could rely on the file extension:
        // doc.Save(outputPath);
    }
}
