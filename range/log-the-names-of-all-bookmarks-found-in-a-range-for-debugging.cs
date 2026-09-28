using System;
using Aspose.Words;
using AsposeRange = Aspose.Words.Range;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add first bookmark.
        builder.StartBookmark("FirstBookmark");
        builder.Writeln("This is the first bookmarked paragraph.");
        builder.EndBookmark("FirstBookmark");

        // Add second bookmark.
        builder.StartBookmark("SecondBookmark");
        builder.Writeln("This is the second bookmarked paragraph.");
        builder.EndBookmark("SecondBookmark");

        // Obtain the document's range using the Aspose.Words.Range alias to avoid ambiguity.
        AsposeRange range = doc.Range;

        // Log the names of all bookmarks found in the range.
        foreach (Bookmark bookmark in range.Bookmarks)
        {
            Console.WriteLine($"Bookmark name: {bookmark.Name}");
        }

        // Save the document (optional, demonstrates full lifecycle).
        doc.Save("SampleDocument.docx");
    }
}
