using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some sample text to the document body.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the first line.");
        builder.Writeln("This is the second line.");

        // At this point the document contains text.
        // Delete all characters in the whole document by calling Delete on its Range.
        doc.Range.Delete();

        // Define an output path for the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeletedContent.docx");

        // Save the modified (now empty) document.
        doc.Save(outputPath);
    }
}
