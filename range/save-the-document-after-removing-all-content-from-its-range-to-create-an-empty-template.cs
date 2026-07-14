using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some sample content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This text will be removed.");
        builder.Writeln("Another line to delete.");

        // Delete all characters in the document's range, leaving an empty template.
        doc.Range.Delete();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the empty template.
        string outputPath = Path.Combine(outputDir, "EmptyTemplate.docx");
        doc.Save(outputPath);
    }
}
