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

        // Add some text using different fonts.
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world with Arial.");

        builder.Font.Name = "Times New Roman";
        builder.Writeln("Hello world with Times New Roman.");

        // Configure PDF save options to embed the full fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Define the output file path.
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "EmbeddedFonts.pdf");

        // Save the document as PDF with full font embedding.
        doc.Save(outputFile, pdfOptions);

        // Verify that the PDF file was created.
        Console.WriteLine(File.Exists(outputFile)
            ? $"PDF saved successfully: {outputFile}"
            : "Failed to save PDF.");
    }
}
