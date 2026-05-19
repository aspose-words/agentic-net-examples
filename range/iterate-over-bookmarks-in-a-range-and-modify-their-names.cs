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

            // Add sample bookmarks to the document.
            for (int i = 1; i <= 3; i++)
            {
                string bookmarkName = $"MyBookmark_{i}";
                builder.StartBookmark(bookmarkName);
                builder.Write($"Text inside {bookmarkName}.");
                builder.EndBookmark(bookmarkName);
                builder.Writeln(); // Add a line break after each bookmark.
            }

            // Iterate over all bookmarks in the document's range and modify their names.
            BookmarkCollection bookmarks = doc.Range.Bookmarks;
            foreach (Bookmark bookmark in bookmarks)
            {
                // Append "_Modified" to each bookmark name.
                bookmark.Name = $"{bookmark.Name}_Modified";
            }

            // Save the modified document.
            doc.Save("ModifiedBookmarks.docx");
        }
    }
}
