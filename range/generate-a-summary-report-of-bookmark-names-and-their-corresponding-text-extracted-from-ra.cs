using System;
using System.IO;
using Aspose.Words;

public class BookmarkSummaryReport
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample content with bookmarks.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"Bookmark_{i}";
            builder.Write($"Text before {bookmarkName}. ");
            builder.StartBookmark(bookmarkName);
            builder.Write($"This is the content of {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln($" Text after {bookmarkName}.");
        }

        // Save the sample document (optional, for verification).
        string docPath = "SampleDocument.docx";
        doc.Save(docPath);

        // Build the summary report.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;
        using (StringWriter reportWriter = new StringWriter())
        {
            reportWriter.WriteLine("Bookmark Summary Report");
            reportWriter.WriteLine("-----------------------");
            foreach (Bookmark bookmark in bookmarks)
            {
                // Bookmark.Name gives the name, Bookmark.Text gives the enclosed text.
                reportWriter.WriteLine($"Name: {bookmark.Name}");
                reportWriter.WriteLine($"Text: {bookmark.Text}");
                reportWriter.WriteLine();
            }

            // Save the report to a text file.
            string reportPath = "BookmarkReport.txt";
            File.WriteAllText(reportPath, reportWriter.ToString());

            // Output paths to console for quick reference.
            Console.WriteLine($"Document saved to: {Path.GetFullPath(docPath)}");
            Console.WriteLine($"Report saved to: {Path.GetFullPath(reportPath)}");
        }
    }
}
