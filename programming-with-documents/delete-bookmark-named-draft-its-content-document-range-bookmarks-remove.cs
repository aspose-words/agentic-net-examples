using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Locate the bookmark named "Draft".
        Bookmark bookmark = doc.Range.Bookmarks["Draft"];
        if (bookmark != null)
        {
            // Remove the text that the bookmark encloses.
            bookmark.Text = string.Empty;

            // Delete the bookmark itself from the collection.
            doc.Range.Bookmarks.Remove(bookmark);
        }

        // Save the updated document.
        doc.Save("Output.docx");
    }
}
