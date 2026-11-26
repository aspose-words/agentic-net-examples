using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph.");
        builder.Writeln("Another line of text.");

        // Delete all characters in the document's body by deleting the whole range.
        doc.Range.Delete();

        // Save the resulting (empty) document.
        doc.Save("DeletedContent.docx");

        // Optional verification: the document's range text should be empty.
        Console.WriteLine(string.IsNullOrEmpty(doc.Range.Text) ? "Document is empty." : "Document still contains text.");
    }
}
