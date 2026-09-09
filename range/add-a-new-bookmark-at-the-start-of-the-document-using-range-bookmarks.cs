using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some initial content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the first paragraph of the document.");

        // Move the builder cursor to the very start of the document.
        builder.MoveToDocumentStart();

        // Insert a zero‑length bookmark at the start of the document.
        // A valid bookmark requires both a start and an end node with the same name.
        builder.StartBookmark("StartBookmark");
        builder.EndBookmark("StartBookmark");

        // Save the document to verify that the bookmark was added.
        doc.Save("BookmarkAtStart.docx");
    }
}
