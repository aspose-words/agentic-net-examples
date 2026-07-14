using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure the font: make it bold and underlined.
        builder.Font.Bold = true;
        builder.Font.Underline = Underline.Single;

        // Insert the header text with the configured formatting.
        builder.Writeln("Bold and Underlined Header");

        // Reset formatting for any further content (optional).
        builder.Font.ClearFormatting();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderFormatted.docx");
        doc.Save(outputPath);
    }
}
