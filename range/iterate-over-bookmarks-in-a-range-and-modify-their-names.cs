using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a few bookmarks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"MyBookmark_{i}";
            builder.Write($"Text before {bookmarkName}. ");
            builder.StartBookmark(bookmarkName);
            builder.Write($"Content of {bookmarkName}. ");
            builder.EndBookmark(bookmarkName);
            builder.Writeln($" Text after {bookmarkName}.");
        }

        // Iterate over all bookmarks in the document's range and modify their names.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;
        foreach (Bookmark bookmark in bookmarks)
        {
            // Append a suffix to each bookmark name.
            bookmark.Name = $"{bookmark.Name}_Modified";
        }

        // Display the updated bookmark names.
        Console.WriteLine("Updated bookmark names:");
        foreach (Bookmark bookmark in doc.Range.Bookmarks)
        {
            Console.WriteLine(bookmark.Name);
        }

        // Save the modified document.
        doc.Save("ModifiedBookmarks.docx");
    }
}
