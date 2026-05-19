using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark with some text inside it.
        string bookmarkName = "MyBookmark";
        builder.StartBookmark(bookmarkName);
        builder.Write("This text will be cleared.");
        builder.EndBookmark(bookmarkName);

        // Save the original document (optional, for reference).
        doc.Save("Original.docx");

        // Retrieve the bookmark by name.
        Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];

        // Clear the text inside the bookmark without removing the bookmark itself.
        bookmark.Text = string.Empty;

        // Save the modified document.
        doc.Save("ClearedBookmark.docx");

        // Output the document's full text to verify that the bookmark remains but its content is empty.
        Console.WriteLine("Document text after clearing bookmark content:");
        Console.WriteLine(doc.GetText());
    }
}
