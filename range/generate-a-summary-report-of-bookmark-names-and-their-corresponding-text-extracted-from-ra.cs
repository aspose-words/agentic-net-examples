using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample bookmarks with some text inside each.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"MyBookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Add a line break after each bookmark.
        }

        // Save the sample document (optional, for verification).
        string docPath = Path.Combine(Environment.CurrentDirectory, "SampleDoc.docx");
        doc.Save(docPath);

        // Extract bookmark names and their enclosed text.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;
        StringBuilder reportBuilder = new StringBuilder();

        foreach (Bookmark bm in bookmarks)
        {
            string line = $"Bookmark: {bm.Name}, Text: {bm.Text}";
            reportBuilder.AppendLine(line);
        }

        // Write the summary report to a text file.
        string reportPath = Path.Combine(Environment.CurrentDirectory, "BookmarkReport.txt");
        File.WriteAllText(reportPath, reportBuilder.ToString());

        // Output paths so the user can locate the files (no interactive input required).
        Console.WriteLine($"Document saved to: {docPath}");
        Console.WriteLine($"Bookmark report saved to: {reportPath}");
    }
}
