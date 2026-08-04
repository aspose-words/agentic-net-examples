using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some sample content that will be removed.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This text will be removed.");

        // Delete all characters in the document's range, leaving an empty template.
        doc.Range.Delete();

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Save the empty template.
        string outputPath = Path.Combine(outputDir, "EmptyTemplate.docx");
        doc.Save(outputPath);
    }
}
