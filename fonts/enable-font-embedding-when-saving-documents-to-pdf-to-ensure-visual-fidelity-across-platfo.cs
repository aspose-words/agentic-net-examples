using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with sample text.
        builder.Writeln("This is a sample text to demonstrate font embedding in PDF.");

        // Set the font for the paragraph.
        builder.Font.Name = "Arial";

        // Define PDF save options with font embedding enabled.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Embed all fonts used in the document.
            EmbedFullFonts = true
        };

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EmbeddedFontOutput.pdf");

        // Save the document as PDF with the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        bool fileExists = File.Exists(outputPath);
        Console.WriteLine($"PDF saved to '{outputPath}'. File exists: {fileExists}");
    }
}
