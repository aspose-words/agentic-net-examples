using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();

        // Set the document-wide default font to Calibri, size 11.
        doc.Styles.DefaultFont.Name = "Calibri";
        doc.Styles.DefaultFont.Size = 11;

        // Use DocumentBuilder to add new paragraphs.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph using the default font.");
        builder.Writeln("Second paragraph also using the default font.");

        // Save the document.
        string outputPath = Path.Combine(outputDir, "DefaultFont.docx");
        doc.Save(outputPath);
    }
}
