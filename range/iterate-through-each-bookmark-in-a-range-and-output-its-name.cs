using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add bookmarks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.StartBookmark("FirstBookmark");
        builder.Writeln("This is the first bookmarked text.");
        builder.EndBookmark("FirstBookmark");

        builder.StartBookmark("SecondBookmark");
        builder.Writeln("Second bookmarked text here.");
        builder.EndBookmark("SecondBookmark");

        // Save the document.
        doc.Save("Bookmarks.docx");

        // Load the document.
        Document loadedDoc = new Document("Bookmarks.docx");

        // Iterate through each bookmark in the document's range and output its name.
        foreach (Bookmark bookmark in loadedDoc.Range.Bookmarks)
        {
            Console.WriteLine($"Bookmark name: {bookmark.Name}");
        }
    }
}
