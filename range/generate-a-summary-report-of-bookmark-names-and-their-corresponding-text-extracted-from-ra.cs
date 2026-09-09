using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some bookmarks with text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First bookmark.
        builder.StartBookmark("FirstBookmark");
        builder.Write("This is the first bookmark text.");
        builder.EndBookmark("FirstBookmark");
        builder.Writeln(); // New line.

        // Second bookmark.
        builder.StartBookmark("SecondBookmark");
        builder.Write("Second bookmark contains different text.");
        builder.EndBookmark("SecondBookmark");
        builder.Writeln();

        // Third bookmark.
        builder.StartBookmark("ThirdBookmark");
        builder.Write("Third bookmark's content.");
        builder.EndBookmark("ThirdBookmark");
        builder.Writeln();

        // Save the sample document (optional, just for verification).
        string docPath = Path.Combine(Environment.CurrentDirectory, "SampleDocument.docx");
        doc.Save(docPath);

        // Extract bookmark names and their corresponding text.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;
        StringBuilder reportBuilder = new StringBuilder();
        reportBuilder.AppendLine("Bookmark Summary Report");
        reportBuilder.AppendLine("-----------------------");

        foreach (Bookmark bookmark in bookmarks)
        {
            // Bookmark.Name gives the name, Bookmark.Text gives the enclosed text.
            string line = $"Name: {bookmark.Name}, Text: {bookmark.Text}";
            reportBuilder.AppendLine(line);
        }

        // Write the report to a text file.
        string reportPath = Path.Combine(Environment.CurrentDirectory, "BookmarkReport.txt");
        File.WriteAllText(reportPath, reportBuilder.ToString());

        // Also output the report to the console.
        Console.WriteLine(reportBuilder.ToString());
    }
}
