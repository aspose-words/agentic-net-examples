using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add sample text to the document body.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph.");
        builder.Writeln("Another line of text.");

        // Delete all characters in the document's body using the Range.Delete method.
        // The Range obtained from the Document represents the whole document.
        doc.Range.Delete();

        // Save the resulting document (it will be essentially empty).
        string outputPath = "DeletedBody.docx";
        doc.Save(outputPath);
    }
}
