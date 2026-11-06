using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is sample content that will be removed.");

        // Remove all content from the document's range to make it empty.
        doc.Range.Delete();

        // Save the empty template.
        doc.Save("EmptyTemplate.docx");
    }
}
