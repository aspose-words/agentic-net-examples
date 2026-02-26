using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content with a paragraph graphic (underline).
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Underline = Underline.Single; // This creates a paragraph graphic.
        builder.Writeln("This text will be marked as an artifact in the PDF.");

        // Configure PDF save options to export paragraph graphics as artifacts.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true,          // Required for artifact marking to take effect.
            ExportParagraphGraphicsToArtifact = true // Mark paragraph graphics as artifacts.
        };

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", saveOptions);
    }
}
