using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three bookmarks with some text.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"Bookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Add a line break.
        }

        // Save the original document (optional, shows the state before removal).
        doc.Save("Original.docx");

        // Locate the specific bookmark to remove.
        Bookmark bookmarkToRemove = doc.Range.Bookmarks["Bookmark_2"];
        if (bookmarkToRemove != null)
        {
            // Remove the bookmark. The text inside the bookmark remains in the document.
            bookmarkToRemove.Remove();
        }

        // Save the modified document.
        doc.Save("Modified.docx");

        // Output the names of the remaining bookmarks to verify removal.
        Console.WriteLine("Remaining bookmarks:");
        foreach (Bookmark bm in doc.Range.Bookmarks)
        {
            Console.WriteLine(bm.Name);
        }
    }
}
