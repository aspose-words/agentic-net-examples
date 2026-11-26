using System;
using Aspose.Words;
using AsposeRange = Aspose.Words.Range;

namespace BookmarkLogger
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert first bookmark.
            builder.StartBookmark("FirstBookmark");
            builder.Writeln("This is the first bookmarked text.");
            builder.EndBookmark("FirstBookmark");

            // Insert second bookmark.
            builder.StartBookmark("SecondBookmark");
            builder.Writeln("This is the second bookmarked text.");
            builder.EndBookmark("SecondBookmark");

            // Save the document so we can verify the bookmarks exist if needed.
            doc.Save("SampleWithBookmarks.docx");

            // Retrieve the range that represents the whole document.
            AsposeRange range = doc.Range;

            // Log the names of all bookmarks found in the range.
            foreach (Bookmark bookmark in range.Bookmarks)
            {
                Console.WriteLine($"Bookmark name: {bookmark.Name}");
            }
        }
    }
}
