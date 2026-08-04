using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample text to the document body.
        builder.Writeln("Hello world!");
        builder.Writeln("This text will be removed.");

        // Delete all characters in the document's range.
        doc.Range.Delete();

        // Save the resulting (empty) document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeletedBody.docx");
        doc.Save(outputPath);
    }
}
