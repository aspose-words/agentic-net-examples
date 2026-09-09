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

        // Use DocumentBuilder to add content with different fonts.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with Arial.
        builder.Font.Name = "Arial";
        builder.Writeln("This text is rendered with Arial.");

        // Second paragraph with Times New Roman.
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This text is rendered with Times New Roman.");

        // Third paragraph with a custom font (if available on the system).
        builder.Font.Name = "Courier New";
        builder.Writeln("This text is rendered with Courier New.");

        // Prepare the output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Define the output PDF file path.
        string outputPath = Path.Combine(outputDir, "EmbeddedFonts.pdf");

        // Configure PDF save options to embed full fonts.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Save the document as PDF with the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"PDF successfully saved with embedded fonts at: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the PDF file.");
        }
    }
}
