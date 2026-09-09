using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some sample content so we can demonstrate the removal.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This text will be removed from the document.");

        // Remove all characters (nodes) from the whole document range.
        doc.Range.Delete();

        // Define the output file path (saved in the current working directory).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EmptyTemplate.docx");

        // Save the now empty document as a template.
        doc.Save(outputPath);
    }
}
