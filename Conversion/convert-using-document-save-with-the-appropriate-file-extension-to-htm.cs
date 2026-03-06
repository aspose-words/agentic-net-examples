using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocmToHtmlConverter
{
    static void Main()
    {
        // Path to the source DOCM file.
        string inputPath = @"C:\Docs\SourceDocument.docm";

        // Load the DOCM document. The format is detected automatically.
        Document doc = new Document(inputPath);

        // Optional: configure HTML save options if needed.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Example option – export headers/footers as part of the body.
            ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection,
            // Example option – embed images as Base64 to keep a single HTML file.
            ExportImagesAsBase64 = true
        };

        // Path to the output HTML file.
        string outputPath = @"C:\Docs\ConvertedDocument.html";

        // Save the document as HTML using the specified options.
        doc.Save(outputPath, saveOptions);

        Console.WriteLine("Conversion completed successfully.");
    }
}
