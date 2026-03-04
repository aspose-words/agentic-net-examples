using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path where the MHTML (Web archive) will be saved.
        string outputPath = @"C:\Docs\ConvertedDocument.mht";

        // Load the DOCM document.
        Document doc = new Document(inputPath);

        // Save the document in MHTML format.
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
