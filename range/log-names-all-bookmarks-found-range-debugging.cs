using System;
using System.Collections.Generic;
using Aspose.Words;

class BookmarkDebugger
{
    // Logs the names of all bookmarks within the specified range.
    static void LogBookmarksInRange(Aspose.Words.Range range)
    {
        // Get the collection of bookmarks from the range.
        BookmarkCollection bookmarks = range.Bookmarks;

        // Iterate through the collection using foreach (simpler and safer).
        foreach (Bookmark bookmark in bookmarks)
        {
            // Guard against null entries (should never happen, but defensive).
            if (bookmark != null)
            {
                // Output the bookmark name for debugging purposes.
                Console.WriteLine($"Bookmark name: {bookmark.Name}");
            }
        }
    }

    static void Main()
    {
        // Load an existing document (replace with your actual file path).
        Document doc = new Document("input.docx");

        // Log all bookmark names found in the whole document range.
        LogBookmarksInRange(doc.Range);

        // Optionally save the document after any modifications.
        doc.Save("output.docx");
    }
}
