using System;
using Aspose.Words;

namespace BookmarkRangeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a few bookmarks with some text inside each.
            for (int i = 1; i <= 3; i++)
            {
                string bookmarkName = $"Bookmark_{i}";
                builder.StartBookmark(bookmarkName);
                builder.Write($"Text inside {bookmarkName}.");
                builder.EndBookmark(bookmarkName);
                builder.Writeln(); // Add a line break after each bookmark.
            }

            // Get the collection of bookmarks that belong to the whole document range.
            BookmarkCollection bookmarks = doc.Range.Bookmarks;

            // Iterate over each bookmark and modify its name.
            foreach (Bookmark bookmark in bookmarks)
            {
                // Append "_Renamed" to the original bookmark name.
                bookmark.Name = $"{bookmark.Name}_Renamed";
            }

            // Save the modified document to disk.
            doc.Save("ModifiedBookmarks.docx");
        }
    }
}
