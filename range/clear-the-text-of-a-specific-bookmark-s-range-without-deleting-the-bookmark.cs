using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark that contains some text.
        const string bookmarkName = "MyBookmark";
        builder.StartBookmark(bookmarkName);
        builder.Write("This text will be cleared.");
        builder.EndBookmark(bookmarkName);

        // Retrieve the bookmark from the document.
        Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];

        // Clear the text inside the bookmark without removing the bookmark itself.
        bookmark.Text = string.Empty;

        // Save the modified document.
        doc.Save("ClearBookmarkText.docx");
    }
}
