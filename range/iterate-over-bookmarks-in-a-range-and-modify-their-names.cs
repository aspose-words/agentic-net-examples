using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add first bookmarked paragraph.
        builder.StartBookmark("Bookmark1");
        builder.Writeln("This is the first bookmarked paragraph.");
        builder.EndBookmark("Bookmark1");

        // Add second bookmarked paragraph.
        builder.StartBookmark("Bookmark2");
        builder.Writeln("This is the second bookmarked paragraph.");
        builder.EndBookmark("Bookmark2");

        // Save the original document (optional, demonstrates creation).
        string originalPath = Path.Combine(Environment.CurrentDirectory, "Original.docx");
        doc.Save(originalPath);

        // Obtain the document's range using the Aspose.Words.Range type to avoid ambiguity.
        Aspose.Words.Range range = doc.Range;

        // Iterate over all bookmarks in the range and modify their names.
        foreach (Bookmark bookmark in range.Bookmarks)
        {
            // Example modification: prefix each bookmark name with "New_".
            string newName = "New_" + bookmark.Name;
            bookmark.Name = newName;
        }

        // Save the modified document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Modified.docx");
        doc.Save(outputPath);
    }
}
