using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Build a simple document with a couple of different fonts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Name = "Arial";
        builder.Writeln("This line uses Arial.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("This line uses Times New Roman.");

        // Configure PDF save options to embed all fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed every font used in the document.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,
            // Optional: embed the full font files (no subsetting) for maximum fidelity.
            EmbedFullFonts = true
        };

        // Save the document as PDF with the specified options.
        string pdfPath = Path.Combine(artifactsDir, "EmbeddedFonts.pdf");
        doc.Save(pdfPath, pdfOptions);

        // Simple verification that the file was created.
        if (File.Exists(pdfPath))
        {
            Console.WriteLine($"PDF saved successfully to: {pdfPath}");
        }
        else
        {
            Console.WriteLine("Failed to save PDF.");
        }
    }
}
