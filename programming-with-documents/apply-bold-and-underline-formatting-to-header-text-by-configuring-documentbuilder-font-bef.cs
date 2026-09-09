using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure the font: make it bold and apply a single underline.
        builder.Font.Bold = true;
        builder.Font.Underline = Underline.Single;
        builder.Font.Size = 24;          // optional: increase font size
        builder.Font.Name = "Arial";     // optional: set font name

        // Insert the header text with the configured formatting.
        builder.Writeln("Document Header");

        // Reset formatting for any subsequent text (optional).
        builder.Font.ClearFormatting();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document to a file.
        string outputPath = Path.Combine(outputDir, "HeaderFormatted.docx");
        doc.Save(outputPath);
    }
}
