using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderDocumentWithTitle
{
    static void Main()
    {
        // Define output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Set the built‑in document title. This value will be used for the PDF window title.
        doc.BuiltInDocumentProperties.Title = "My Document Title";

        // Create PdfSaveOptions and enable the DisplayDocTitle flag.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DisplayDocTitle = true // Show the document title in the PDF title bar.
        };

        // Save the document as PDF using the options defined above.
        string outputPath = Path.Combine(artifactsDir, "DocumentWithTitle.pdf");
        doc.Save(outputPath, pdfOptions);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
