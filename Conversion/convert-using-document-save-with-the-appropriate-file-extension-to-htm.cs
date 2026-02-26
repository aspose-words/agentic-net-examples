using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocmToHtmlConverter
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Path to the output HTML file. The .html extension tells Aspose.Words the desired format.
        string outputPath = @"C:\Docs\ConvertedDocument.html";

        // Load the DOCM document. The constructor automatically detects the format (DOCM).
        Document doc = new Document(inputPath);

        // Save the document as HTML. Explicitly specify SaveFormat.Html for clarity.
        doc.Save(outputPath, SaveFormat.Html);

        Console.WriteLine("Conversion completed successfully.");
    }
}
