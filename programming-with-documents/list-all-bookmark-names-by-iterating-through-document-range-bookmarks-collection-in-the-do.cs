using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample bookmarks into the document.
        builder.StartBookmark("FirstBookmark");
        builder.Writeln("Hello");
        builder.EndBookmark("FirstBookmark");

        builder.StartBookmark("SecondBookmark");
        builder.Writeln("World");
        builder.EndBookmark("SecondBookmark");

        // Iterate through all bookmarks and output their names.
        foreach (Bookmark bookmark in doc.Range.Bookmarks)
        {
            Console.WriteLine(bookmark.Name);
        }
    }
}
