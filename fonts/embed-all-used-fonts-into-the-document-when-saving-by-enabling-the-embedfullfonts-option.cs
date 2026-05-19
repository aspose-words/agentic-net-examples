using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a directory for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text using different fonts.
        builder.Font.Name = "Arial";
        builder.Writeln("This paragraph uses the Arial font.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("This paragraph uses the Times New Roman font.");

        // Configure PDF save options to embed full fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Save the document as PDF with the specified options.
        string outputPath = Path.Combine(artifactsDir, "EmbeddedFonts.pdf");
        doc.Save(outputPath, pdfOptions);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the PDF document.");
        }
    }
}
