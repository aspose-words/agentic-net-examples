using System;
using Aspose.Words;

namespace BookmarkLister
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a few sample bookmarks into the document.
            for (int i = 1; i <= 3; i++)
            {
                string bookmarkName = $"SampleBookmark_{i}";
                builder.StartBookmark(bookmarkName);
                builder.Write($"Text inside {bookmarkName}.");
                builder.EndBookmark(bookmarkName);
                builder.Writeln(); // Add a line break after each bookmark.
            }

            // Get the collection of bookmarks from the document's range.
            BookmarkCollection bookmarks = doc.Range.Bookmarks;

            // Iterate through the collection and print each bookmark's name.
            Console.WriteLine("Bookmarks found in the document:");
            foreach (Bookmark bookmark in bookmarks)
            {
                Console.WriteLine(bookmark.Name);
            }

            // The document can be saved if needed.
            // doc.Save("SampleBookmarks.docx");
        }
    }
}
