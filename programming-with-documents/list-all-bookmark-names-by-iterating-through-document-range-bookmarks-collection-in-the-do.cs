using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few bookmarks into the document.
        for (int i = 1; i <= 3; i++)
        {
            string name = $"MyBookmark_{i}";
            builder.StartBookmark(name);
            builder.Write($"Text inside {name}.");
            builder.EndBookmark(name);
            builder.Writeln(); // Add a line break after each bookmark.
        }

        // Save the document (optional, but satisfies the lifecycle rule).
        doc.Save("Bookmarks.docx");

        // Retrieve the collection of bookmarks from the document's range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Iterate through the collection and print each bookmark's name.
        foreach (Bookmark bookmark in bookmarks)
        {
            Console.WriteLine(bookmark.Name);
        }
    }
}
