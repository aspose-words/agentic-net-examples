using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new document and add some text with different fonts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Name = "Arial";
        builder.Writeln("This text uses Arial.");

        builder.Font.Name = "Courier New";
        builder.Writeln("This text uses Courier New, a non‑standard font.");

        // Configure PDF save options to embed all fonts fully.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true,
            FontEmbeddingMode = Aspose.Words.Saving.PdfFontEmbeddingMode.EmbedAll
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
