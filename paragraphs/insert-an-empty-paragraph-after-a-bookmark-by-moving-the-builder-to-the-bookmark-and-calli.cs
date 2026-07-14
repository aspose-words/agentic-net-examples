using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark named "MyBookmark" with some text inside it.
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Text inside the bookmark.");
        builder.EndBookmark("MyBookmark");

        // Move the builder's cursor to the start of the bookmark.
        builder.MoveToBookmark("MyBookmark");

        // Insert an empty paragraph immediately after the bookmark.
        builder.InsertParagraph();

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
