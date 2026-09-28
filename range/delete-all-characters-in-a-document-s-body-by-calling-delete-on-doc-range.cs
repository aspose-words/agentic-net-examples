using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");
        builder.Writeln("This is a second paragraph.");

        // Save the original document (optional, just for reference).
        doc.Save("original.docx");

        // Delete all characters in the document's body by deleting the whole document range.
        doc.Range.Delete();

        // Save the resulting document after deletion.
        doc.Save("deleted.docx");
    }
}
