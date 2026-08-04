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
        builder.Writeln("Text inside bookmark.");
        builder.EndBookmark("MyBookmark");

        // Move the builder's cursor to the start of the bookmark.
        builder.MoveToBookmark("MyBookmark");

        // Insert an empty paragraph immediately after the bookmark.
        // InsertParagraph returns the newly created Paragraph, which will be empty.
        Paragraph emptyParagraph = builder.InsertParagraph();

        // (Optional) Verify that the inserted paragraph is empty.
        // The paragraph text consists only of the paragraph break character.
        // Console.WriteLine($"Inserted paragraph text length: {emptyParagraph.GetText().Length}");

        // Save the document to the current directory.
        doc.Save("ParagraphAfterBookmark.docx");
    }
}
