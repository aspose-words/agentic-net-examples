using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add first bookmark.
        builder.StartBookmark("Bookmark1");
        builder.Writeln("First bookmark text.");
        builder.EndBookmark("Bookmark1");

        // Add second bookmark.
        builder.StartBookmark("Bookmark2");
        builder.Writeln("Second bookmark text.");
        builder.EndBookmark("Bookmark2");

        // Optionally save the document (not required for iteration, but demonstrates lifecycle).
        doc.Save("Bookmarks.docx");

        // Iterate through each bookmark in the document's range and output its name.
        foreach (Bookmark bookmark in doc.Range.Bookmarks)
        {
            Console.WriteLine(bookmark.Name);
        }
    }
}
