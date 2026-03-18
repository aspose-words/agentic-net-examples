using System;
using Aspose.Words;

class RemoveBookmarkExample
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Locate the bookmark by its name.
        Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];

        // If the bookmark exists, remove it from the document.
        if (bookmark != null)
        {
            bookmark.Remove(); // Removes the bookmark but keeps its text.
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
