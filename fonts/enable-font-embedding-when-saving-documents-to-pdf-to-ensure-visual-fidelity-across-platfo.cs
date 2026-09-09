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
        builder.Writeln("This paragraph uses Arial.");

        builder.Font.Name = "Courier New";
        builder.Writeln("This paragraph uses Courier New.");

        // Configure PDF save options to embed all fonts in the output PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true,
            FontEmbeddingMode = Aspose.Words.Saving.PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF.
        string pdfPath = Path.Combine(artifactsDir, "EmbeddedFonts.pdf");
        doc.Save(pdfPath, saveOptions);

        // Verify that the PDF file was created.
        Console.WriteLine(File.Exists(pdfPath)
            ? $"PDF saved successfully to: {pdfPath}"
            : "Failed to save PDF.");
    }
}
