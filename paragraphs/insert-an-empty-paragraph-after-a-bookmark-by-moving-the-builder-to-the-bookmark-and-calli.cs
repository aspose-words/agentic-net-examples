using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content.
        builder.Writeln("Paragraph before the bookmark.");

        // Insert a bookmark named "MyBookmark".
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Text inside the bookmark.");
        builder.EndBookmark("MyBookmark");

        // Move the builder to the bookmark.
        builder.MoveToBookmark("MyBookmark");

        // Insert an empty paragraph after the bookmark.
        builder.InsertParagraph();

        // Add more content after the empty paragraph to demonstrate placement.
        builder.Writeln("Paragraph after the empty paragraph.");

        // Save the document.
        doc.Save("Output.docx");
    }
}
