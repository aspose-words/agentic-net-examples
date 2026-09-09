using System;
using Aspose.Words;

namespace BookmarkLogger
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a few bookmarks with some surrounding text.
            for (int i = 1; i <= 3; i++)
            {
                string bookmarkName = $"Bookmark_{i}";

                builder.Writeln($"Text before {bookmarkName}.");
                builder.StartBookmark(bookmarkName);
                builder.Writeln($"Content inside {bookmarkName}.");
                builder.EndBookmark(bookmarkName);
                builder.Writeln($"Text after {bookmarkName}.");
            }

            // Retrieve the collection of all bookmarks in the document's range.
            BookmarkCollection bookmarks = doc.Range.Bookmarks;

            // Log each bookmark's name to the console for debugging.
            foreach (Bookmark bookmark in bookmarks)
            {
                Console.WriteLine($"Bookmark name: {bookmark.Name}");
            }
        }
    }
}
