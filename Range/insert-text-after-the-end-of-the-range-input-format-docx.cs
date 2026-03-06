using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // The range is defined by a bookmark named "InsertHere".
        // Retrieve the bookmark from the document.
        Bookmark bookmark = doc.Range.Bookmarks["InsertHere"];
        if (bookmark != null)
        {
            // Create a DocumentBuilder positioned at the end of the bookmark (range end).
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(bookmark.BookmarkEnd);

            // Insert the desired text immediately after the range.
            builder.Write("This text is inserted after the range.");
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
