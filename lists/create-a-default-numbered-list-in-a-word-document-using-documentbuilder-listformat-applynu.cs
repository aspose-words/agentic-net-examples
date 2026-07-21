using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Add several list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document to the output folder.
        string outputPath = Path.Combine(outputDir, "DefaultNumberedList.docx");
        doc.Save(outputPath);
    }
}
