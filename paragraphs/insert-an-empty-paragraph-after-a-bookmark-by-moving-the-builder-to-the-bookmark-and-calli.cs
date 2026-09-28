using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content.
        builder.Writeln("Paragraph before bookmark.");

        // Insert a bookmark named "MyBookmark".
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Text inside bookmark.");
        builder.EndBookmark("MyBookmark");

        // Add more content after the bookmark.
        builder.Writeln("Paragraph after bookmark.");

        // Move the builder to the bookmark.
        builder.MoveToBookmark("MyBookmark");

        // Insert an empty paragraph after the bookmark.
        builder.InsertParagraph();

        // Save the document.
        doc.Save("Output.docx");
    }
}
