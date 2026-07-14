using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some bookmarks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"MyBookmark_{i}";
            builder.Write($"Text before {bookmarkName}. ");
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln($" Text after {bookmarkName}.");
        }

        // Extract bookmark names and their enclosed text.
        StringBuilder report = new StringBuilder();
        foreach (Bookmark bm in doc.Range.Bookmarks)
        {
            string line = $"Bookmark: {bm.Name}, Text: {bm.Text}";
            report.AppendLine(line);
        }

        // Save the report to a local text file.
        string reportPath = Path.Combine(Environment.CurrentDirectory, "BookmarkReport.txt");
        File.WriteAllText(reportPath, report.ToString());

        // Optionally, save the document for inspection.
        string docPath = Path.Combine(Environment.CurrentDirectory, "SampleDocument.docx");
        doc.Save(docPath);
    }
}
