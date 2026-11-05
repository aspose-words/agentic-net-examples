using System;
using Aspose.Words;

namespace BookmarkLogger
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a builder to add content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add first bookmark.
            builder.StartBookmark("FirstBookmark");
            builder.Writeln("This is the first bookmarked text.");
            builder.EndBookmark("FirstBookmark");

            // Add second bookmark.
            builder.StartBookmark("SecondBookmark");
            builder.Writeln("This is the second bookmarked text.");
            builder.EndBookmark("SecondBookmark");

            // Get the document's range (fully qualified to avoid ambiguity with System.Range).
            Aspose.Words.Range range = doc.Range;

            // Log all bookmark names found in the range.
            foreach (Bookmark bookmark in range.Bookmarks)
            {
                Console.WriteLine($"Bookmark name: {bookmark.Name}");
            }

            // Save the document for verification.
            doc.Save("SampleDocument.docx");
        }
    }
}
