using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for inserting content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first bookmark.
        builder.StartBookmark("First");
        builder.Writeln("This is the first bookmarked paragraph.");
        builder.EndBookmark("First");

        // Insert second bookmark.
        builder.StartBookmark("Second");
        builder.Writeln("This is the second bookmarked paragraph.");
        builder.EndBookmark("Second");

        // Iterate over all bookmarks in the document's range and rename them.
        foreach (Bookmark bookmark in doc.Range.Bookmarks)
        {
            string oldName = bookmark.Name;
            bookmark.Name = "New_" + oldName;
        }

        // Save the modified document.
        doc.Save("ModifiedBookmarks.docx");
    }
}
