using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section with a bookmark.
        builder.Writeln("First section - start.");
        builder.StartBookmark("FirstSectionBookmark");
        builder.Writeln("This paragraph is inside a bookmark in the first section.");
        builder.EndBookmark("FirstSectionBookmark");
        builder.Writeln("First section - end.");

        // Insert a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Second section - start.");
        builder.StartBookmark("SecondSectionBookmark");
        builder.Writeln("This paragraph is inside a bookmark in the second section.");
        builder.EndBookmark("SecondSectionBookmark");
        builder.Writeln("Second section - end.");

        // Count all bookmarks in the whole document range.
        int totalBookmarks = doc.Range.Bookmarks.Count;
        Console.WriteLine($"Total bookmarks in the document: {totalBookmarks}");

        // Count bookmarks that reside in the first section only.
        Section firstSection = doc.Sections[0];
        int bookmarksInFirstSection = 0;

        foreach (Bookmark bookmark in doc.Range.Bookmarks)
        {
            // Determine if the bookmark's start node is inside the first section.
            Node current = bookmark.BookmarkStart.ParentNode;
            while (current != null && !(current is Section))
                current = current.ParentNode;

            if (current == firstSection)
                bookmarksInFirstSection++;
        }

        Console.WriteLine($"Bookmarks in the first section: {bookmarksInFirstSection}");
    }
}
