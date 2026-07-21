using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class BookmarkSummaryReport
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample content with three bookmarks.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"MyBookmark_{i}";
            builder.Write($"Text before {bookmarkName}. ");
            builder.StartBookmark(bookmarkName);
            builder.Write($"This is the content of {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // New line after each bookmark.
        }

        // Save the sample document (optional, demonstrates saving).
        doc.Save("SampleDocument.docx");

        // Retrieve the collection of bookmarks from the document's range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Build a textual report of bookmark names and their enclosed text.
        StringBuilder reportBuilder = new StringBuilder();
        reportBuilder.AppendLine("Bookmark Summary Report");
        reportBuilder.AppendLine("-----------------------");

        foreach (Bookmark bookmark in bookmarks)
        {
            reportBuilder.AppendLine($"Name: {bookmark.Name}");
            reportBuilder.AppendLine($"Text: {bookmark.Text}");
            reportBuilder.AppendLine();
        }

        // Write the report to a text file.
        File.WriteAllText("BookmarkReport.txt", reportBuilder.ToString());

        // Output the report path to the console for verification.
        Console.WriteLine("Bookmark report generated: BookmarkReport.txt");
    }
}
