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

            // Add a few sample bookmarks.
            for (int i = 1; i <= 3; i++)
            {
                string name = $"Bookmark_{i}";
                builder.StartBookmark(name);
                builder.Write($"Text inside {name}.");
                builder.EndBookmark(name);
                builder.Writeln(); // Move to next line.
            }

            // Retrieve the collection of bookmarks from the document.
            BookmarkCollection bookmarks = doc.Range.Bookmarks;

            // Iterate through the collection and print each bookmark's name.
            foreach (Bookmark bm in bookmarks)
            {
                Console.WriteLine(bm.Name);
            }

            // Optional: save the document to verify the bookmarks exist.
            doc.Save("SampleWithBookmarks.docx");
        }
    }
}
