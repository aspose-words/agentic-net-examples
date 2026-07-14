using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text to the document body.
        builder.Writeln("Hello World!");
        builder.Writeln("This text will be deleted.");

        // Save the original document (optional, for verification).
        doc.Save("Original.docx");

        // Delete all characters in the document's body using the Range.Delete method.
        doc.Range.Delete();

        // Save the document after deletion.
        doc.Save("DeletedContent.docx");
    }
}
