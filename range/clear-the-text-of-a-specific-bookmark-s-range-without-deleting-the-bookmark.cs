using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a bookmark with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartBookmark("MyBookmark");
        builder.Writeln("This is the text inside the bookmark.");
        builder.EndBookmark("MyBookmark");

        // Save the original document.
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Load the document from the saved file.
        Document loadedDoc = new Document(originalPath);

        // Retrieve the bookmark and clear its text without deleting the bookmark.
        Bookmark bookmark = loadedDoc.Range.Bookmarks["MyBookmark"];
        if (bookmark != null)
        {
            bookmark.Text = string.Empty; // Clears the bookmark's content.
        }

        // Save the modified document.
        string resultPath = "ClearedBookmark.docx";
        loadedDoc.Save(resultPath);
    }
}
