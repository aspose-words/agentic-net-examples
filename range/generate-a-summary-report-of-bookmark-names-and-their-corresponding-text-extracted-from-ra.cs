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

        // Add three bookmarks with distinct text.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"MyBookmark_{i}";
            builder.Write($"Text before {bookmarkName}. ");
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln($" Text after {bookmarkName}.");
        }

        // Save the sample document (optional, demonstrates lifecycle usage).
        string docPath = "SampleDocument.docx";
        doc.Save(docPath);

        // Extract bookmark names and their enclosed text.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;
        string[] reportLines = new string[bookmarks.Count];
        for (int i = 0; i < bookmarks.Count; i++)
        {
            Bookmark bm = bookmarks[i];
            string line = $"Bookmark: {bm.Name}, Text: {bm.Text}";
            reportLines[i] = line;
            Console.WriteLine(line); // Output to console for visibility.
        }

        // Write the summary report to a text file.
        string reportPath = "BookmarkReport.txt";
        File.WriteAllLines(reportPath, reportLines);
    }
}
