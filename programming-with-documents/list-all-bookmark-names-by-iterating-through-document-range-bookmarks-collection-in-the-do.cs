using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add first bookmark.
        builder.StartBookmark("FirstBookmark");
        builder.Writeln("This is the first bookmarked paragraph.");
        builder.EndBookmark("FirstBookmark");

        // Add second bookmark.
        builder.StartBookmark("SecondBookmark");
        builder.Writeln("This is the second bookmarked paragraph.");
        builder.EndBookmark("SecondBookmark");

        // List all bookmark names.
        foreach (Bookmark bookmark in doc.Range.Bookmarks)
        {
            Console.WriteLine($"Bookmark name: {bookmark.Name}");
        }

        // Save the document (optional, demonstrates lifecycle).
        doc.Save("BookmarksDocument.docx");
    }
}
