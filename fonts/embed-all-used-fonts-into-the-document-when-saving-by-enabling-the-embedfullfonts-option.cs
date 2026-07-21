using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class EmbedFontsExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text using different fonts.
        builder.Font.Name = "Arial";
        builder.Writeln("This paragraph uses the Arial font.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("This paragraph uses the Times New Roman font.");

        // Configure PDF save options to embed the full fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EmbeddedFonts.pdf");

        // Save the document as PDF with full font embedding.
        doc.Save(outputPath, pdfOptions);

        // Simple verification that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("PDF saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save PDF.");
        }
    }
}
