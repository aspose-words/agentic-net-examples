using System;
using Aspose.Words;

namespace RemoveBookmarkExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert three bookmarks with text.
            for (int i = 1; i <= 3; i++)
            {
                string name = $"Bookmark{i}";
                builder.StartBookmark(name);
                builder.Write($"Text inside {name}.");
                builder.EndBookmark(name);
                builder.Writeln(); // Add a line break.
            }

            // Locate the bookmark named "Bookmark2" and remove it.
            Bookmark bookmarkToRemove = doc.Range.Bookmarks["Bookmark2"];
            if (bookmarkToRemove != null)
            {
                // The Remove method deletes the bookmark but keeps its text.
                bookmarkToRemove.Remove();
            }

            // Save the document after removal.
            doc.Save("RemovedBookmark.docx");

            // List remaining bookmarks.
            Console.WriteLine("Remaining bookmarks:");
            foreach (Bookmark bm in doc.Range.Bookmarks)
            {
                Console.WriteLine(bm.Name);
            }
        }
    }
}
