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

        // Write a heading before the list.
        builder.Writeln("Bulleted list example:");

        // Start a default bulleted list.
        builder.ListFormat.ApplyBulletDefault();

        // Add list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the bulleted list.
        builder.ListFormat.RemoveNumbers();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "DefaultBulletedList.docx");
        doc.Save(outputPath);
    }
}
