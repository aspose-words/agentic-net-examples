using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Get the collection of bookmarks that belong to the whole document range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Iterate through the collection and change each bookmark's name.
        for (int i = 0; i < bookmarks.Count; i++)
        {
            Bookmark bookmark = bookmarks[i];
            if (bookmark != null)
            {
                // Example modification: append "_Modified" to the original name.
                bookmark.Name = bookmark.Name + "_Modified";
            }
        }

        // Save the document with the updated bookmark names.
        doc.Save("Output.docx");
    }
}
