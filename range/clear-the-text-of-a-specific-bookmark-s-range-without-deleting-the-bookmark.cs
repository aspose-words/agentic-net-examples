using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark with some text inside.
        string bookmarkName = "MyBookmark";
        builder.StartBookmark(bookmarkName);
        builder.Write("This text will be cleared.");
        builder.EndBookmark(bookmarkName);

        // Verify the bookmark exists and contains the original text.
        Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];
        Console.WriteLine("Original bookmark text: " + bookmark.Text);

        // Clear the text of the bookmark without removing the bookmark itself.
        bookmark.Text = string.Empty;

        // Verify that the bookmark still exists and its text is now empty.
        Console.WriteLine("Bookmark text after clearing: '" + bookmark.Text + "'");

        // Save the modified document.
        doc.Save("ClearBookmarkText.docx");
    }
}
