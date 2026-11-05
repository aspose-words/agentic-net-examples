using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add first bookmark.
        builder.StartBookmark("Bookmark1");
        builder.Writeln("This is the first bookmarked text.");
        builder.EndBookmark("Bookmark1");

        // Add second bookmark.
        builder.StartBookmark("Bookmark2");
        builder.Writeln("This is the second bookmarked text.");
        builder.EndBookmark("Bookmark2");

        // Save the document (optional, demonstrates lifecycle usage).
        doc.Save("SampleDocument.docx");

        // Get the range that represents the whole document.
        Aspose.Words.Range range = doc.Range;

        // Iterate through each bookmark in the range and output its name.
        foreach (Bookmark bookmark in range.Bookmarks)
        {
            Console.WriteLine(bookmark.Name);
        }
    }
}
