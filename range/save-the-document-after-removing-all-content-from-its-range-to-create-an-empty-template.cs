using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some sample content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This text will be removed to create an empty template.");

        // Delete all characters in the document's range, leaving an empty template.
        doc.Range.Delete();

        // Save the empty template to a file in the current directory.
        string outputPath = "EmptyTemplate.docx";
        doc.Save(outputPath);

        // Optional: indicate completion (no user interaction required).
        Console.WriteLine($"Document saved as '{outputPath}'.");
    }
}
