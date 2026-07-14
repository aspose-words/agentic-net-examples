using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class EmbedFontsExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text with different fonts.
        builder.Font.Name = "Arial";
        builder.Writeln("This text is in Arial.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("This text is in Times New Roman.");

        // Configure PDF save options to embed full fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Define output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "DocumentWithEmbeddedFonts.pdf");

        // Save the document as PDF with full font embedding.
        doc.Save(outputPath, pdfOptions);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("PDF saved successfully with embedded fonts at:");
            Console.WriteLine(outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the PDF file.");
        }
    }
}
