using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Build a simple document with a couple of different fonts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Name = "Arial";
        builder.Writeln("This line uses Arial.");

        builder.Font.Name = "Courier New";
        builder.Writeln("This line uses Courier New.");

        // Configure PDF save options to embed all fonts fully.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true,                     // Embed the complete font files (no subsetting).
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll // Ensure all fonts are embedded.
        };

        // Save the document as PDF with embedded fonts.
        string outputPath = Path.Combine(artifactsDir, "EmbeddedFonts.pdf");
        doc.Save(outputPath, pdfOptions);

        // Simple verification that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"PDF saved with embedded fonts at: {outputPath}");
        }
    }
}
