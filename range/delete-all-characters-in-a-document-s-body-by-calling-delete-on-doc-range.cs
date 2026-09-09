using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some sample text to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is some sample text that will be deleted.");

        // Delete all characters in the document's body by calling Delete on the document's Range.
        doc.Range.Delete();

        // Save the resulting (empty) document to a file in the current directory.
        doc.Save("DeletedContent.docx");
    }
}
