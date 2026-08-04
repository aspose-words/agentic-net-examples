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

        // Build a simple document with two different fonts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Name = "Arial";
        builder.Writeln("This text uses Arial.");

        builder.Font.Name = "Courier New";
        builder.Writeln("This text uses Courier New.");

        // Configure PDF save options to embed all fonts fully.
        PdfSaveOptions options = new PdfSaveOptions
        {
            EmbedFullFonts = true,
            FontEmbeddingMode = Aspose.Words.Saving.PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF.
        string pdfPath = Path.Combine(artifactsDir, "EmbeddedFonts.pdf");
        doc.Save(pdfPath, options);

        // Verify that the PDF file was created.
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
