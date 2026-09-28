using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first paragraph with a bookmark.
        builder.StartBookmark("bm1");
        builder.Writeln("First paragraph.");
        builder.EndBookmark("bm1");

        // Insert second paragraph with a bookmark.
        builder.StartBookmark("bm2");
        builder.Writeln("Second paragraph.");
        builder.EndBookmark("bm2");

        // Insert third paragraph with a bookmark.
        builder.StartBookmark("bm3");
        builder.Writeln("Third paragraph.");
        builder.EndBookmark("bm3");

        // Save the document (optional, demonstrates lifecycle handling).
        doc.Save("BookmarksDemo.docx");

        // Get the second paragraph.
        Paragraph secondParagraph = doc.FirstSection.Body.Paragraphs[1];

        // Count bookmarks that are fully contained within the second paragraph.
        int bookmarkCount = 0;
        foreach (Bookmark bookmark in doc.Range.Bookmarks)
        {
            // Determine the paragraph that contains the bookmark start.
            Paragraph bookmarkParagraph = bookmark.BookmarkStart.GetAncestor(NodeType.Paragraph) as Paragraph;

            // If the bookmark start resides in the second paragraph, count it.
            if (bookmarkParagraph == secondParagraph)
                bookmarkCount++;
        }

        // Output the result.
        Console.WriteLine($"Number of bookmarks within the selected range: {bookmarkCount}");
    }
}
