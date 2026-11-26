using System;
using Aspose.Words;

namespace BookmarkLister
{
    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text with bookmarks.
            builder.StartBookmark("FirstBookmark");
            builder.Writeln("This is the first bookmarked paragraph.");
            builder.EndBookmark("FirstBookmark");

            builder.StartBookmark("SecondBookmark");
            builder.Writeln("This is the second bookmarked paragraph.");
            builder.EndBookmark("SecondBookmark");

            // Save the document.
            string filePath = "BookmarksSample.docx";
            doc.Save(filePath);

            // Load the document.
            Document loadedDoc = new Document(filePath);

            // List all bookmark names.
            foreach (Bookmark bookmark in loadedDoc.Range.Bookmarks)
            {
                Console.WriteLine($"Bookmark name: {bookmark.Name}");
            }
        }
    }
}
