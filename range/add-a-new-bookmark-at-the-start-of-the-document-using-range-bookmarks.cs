using System;
using Aspose.Words;

namespace AsposeWordsBookmarkExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add some content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("First paragraph of the document.");
            builder.Writeln("Second paragraph of the document.");

            // Create a bookmark that starts at the very beginning of the document.
            // An empty bookmark is created by inserting a BookmarkStart and a matching BookmarkEnd
            // at the same location (before the first paragraph).
            BookmarkStart bookmarkStart = new BookmarkStart(doc, "StartBookmark");
            BookmarkEnd bookmarkEnd = new BookmarkEnd(doc, "StartBookmark");

            // Insert the start node before the first paragraph.
            Node firstParagraph = doc.FirstSection.Body.FirstParagraph;
            firstParagraph.ParentNode.InsertBefore(bookmarkStart, firstParagraph);

            // Insert the end node immediately after the start node.
            bookmarkStart.ParentNode.InsertAfter(bookmarkEnd, bookmarkStart);

            // Access the bookmark collection via the document range.
            BookmarkCollection bookmarks = doc.Range.Bookmarks;

            // Save the document to verify the bookmark was added.
            doc.Save("Output.docx");

            // Simple verification output.
            Console.WriteLine($"Bookmarks count: {bookmarks.Count}");
            if (bookmarks.Count > 0)
            {
                Console.WriteLine($"Bookmark name: {bookmarks[0].Name}");
            }
        }
    }
}
