using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample document with bookmarks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First bookmark.
        builder.StartBookmark("First");
        builder.Writeln("This is the first bookmark.");
        builder.EndBookmark("First");

        // Second bookmark.
        builder.StartBookmark("Second");
        builder.Writeln("Second bookmark contains different text.");
        builder.EndBookmark("Second");

        // Save the source document.
        string sourcePath = "Sample.docx";
        doc.Save(sourcePath);

        // Load the document (bootstrap loading rule).
        Document loadedDoc = new Document(sourcePath);

        // Prepare the summary report.
        using (StringWriter reportWriter = new StringWriter())
        {
            foreach (Bookmark bookmark in loadedDoc.Range.Bookmarks)
            {
                string name = bookmark.Name;
                string text = bookmark.Text; // Extract text within the bookmark range.
                reportWriter.WriteLine($"Bookmark: {name}");
                reportWriter.WriteLine($"Text: {text}");
                reportWriter.WriteLine(); // Blank line for readability.
            }

            // Write the report to a text file.
            string reportPath = "BookmarkReport.txt";
            File.WriteAllText(reportPath, reportWriter.ToString());

            // Also output the report to the console.
            Console.WriteLine(reportWriter.ToString());
        }
    }
}
