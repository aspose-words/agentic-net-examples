using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and a bookmark.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph before the bookmark.");

        // Define a bookmark named "MyBookmark".
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Paragraph inside the bookmark.");
        builder.EndBookmark("MyBookmark");

        // Move the builder to the bookmark location.
        builder.MoveToBookmark("MyBookmark");

        // Insert an empty paragraph after the bookmark.
        builder.InsertParagraph();

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
