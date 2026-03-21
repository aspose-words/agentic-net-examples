using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, world! This is a test document with OpenType features disabled.");

        // Disable OpenType font formatting features for the whole document.
        doc.CompatibilityOptions.DisableOpenTypeFontFormattingFeatures = true;

        // Determine output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ResultDocument.pdf");

        // Save as PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(outputPath, pdfOptions);

        Console.WriteLine($"PDF saved to: {outputPath}");
    }
}
