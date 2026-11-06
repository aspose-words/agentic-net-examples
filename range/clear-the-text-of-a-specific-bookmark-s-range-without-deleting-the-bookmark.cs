using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a bookmark.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document start.");
        builder.StartBookmark("MyBookmark");
        builder.Writeln("This text will be cleared.");
        builder.EndBookmark("MyBookmark");
        builder.Writeln("Document end.");

        // Save the original document.
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Load the document from the saved file.
        Document loadedDoc = new Document(originalPath);

        // Clear the text of the specific bookmark without deleting the bookmark itself.
        Bookmark bookmark = loadedDoc.Range.Bookmarks["MyBookmark"];
        if (bookmark != null)
        {
            bookmark.Text = string.Empty;
        }

        // Save the modified document.
        string outputPath = "ClearedBookmark.docx";
        loadedDoc.Save(outputPath);
    }
}
