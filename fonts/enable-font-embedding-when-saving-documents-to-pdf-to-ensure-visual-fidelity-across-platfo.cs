using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text using different fonts.
        builder.Font.Name = "Arial";
        builder.Writeln("This paragraph uses Arial.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("This paragraph uses Times New Roman.");

        builder.Font.Name = "Courier New";
        builder.Writeln("This paragraph uses Courier New.");

        // Configure PDF save options to embed all fonts fully.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true, // Embed the complete font files (no subsetting).
            FontEmbeddingMode = Aspose.Words.Saving.PdfFontEmbeddingMode.EmbedAll // Embed every font used.
        };

        // Determine an output file path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EmbeddedFonts.pdf");

        // Save the document as PDF using the configured options.
        doc.Save(outputPath, pdfOptions);

        // Validate that the PDF file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"PDF saved successfully: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save PDF.");
        }
    }
}
