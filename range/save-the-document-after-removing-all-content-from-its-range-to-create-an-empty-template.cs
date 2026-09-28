using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some sample content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is sample content.");

        // Remove all content from the document's range, leaving an empty template.
        doc.Range.Delete();

        // Save the resulting empty document.
        doc.Save("EmptyTemplate.docx");
    }
}
