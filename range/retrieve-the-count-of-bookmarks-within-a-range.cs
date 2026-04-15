using System;
using Aspose.Words;

namespace BookmarkCountExample
{
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
                string name = $"Bookmark_{i}";
                builder.StartBookmark(name);
                builder.Write($"Text inside {name}.");
                builder.EndBookmark(name);
                builder.Writeln(); // Add a line break after each bookmark.
            }

            // Retrieve the number of bookmarks that exist in the whole‑document range.
            int bookmarkCount = doc.Range.Bookmarks.Count;

            // Output the count to the console.
            Console.WriteLine($"Total bookmarks in the document range: {bookmarkCount}");
        }
    }
}
