using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some sample content so that we have something to delete.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This text will be removed.");

        // Delete all characters in the document's range, leaving an empty template.
        doc.Range.Delete();

        // Save the resulting empty document.
        string outputFile = "EmptyTemplate.docx";
        doc.Save(outputFile);
    }
}
