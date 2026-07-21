using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several paragraphs inside a loop.
        for (int i = 1; i <= 5; i++)
        {
            // Writeln adds the supplied text and then creates a new paragraph.
            builder.Writeln($"This is paragraph number {i} inserted in a loop.");
        }

        // Save the document to the file system.
        doc.Save("Output.docx");
    }
}
