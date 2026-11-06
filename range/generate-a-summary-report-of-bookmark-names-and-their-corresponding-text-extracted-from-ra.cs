using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

namespace BookmarkRangeReport
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with bookmarks.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First bookmark.
            builder.StartBookmark("FirstBookmark");
            builder.Writeln("This is the text inside the first bookmark.");
            builder.EndBookmark("FirstBookmark");

            // Second bookmark.
            builder.StartBookmark("SecondBookmark");
            builder.Writeln("Second bookmark contains different text.");
            builder.EndBookmark("SecondBookmark");

            // Save the sample document.
            string sourcePath = "Sample.docx";
            doc.Save(sourcePath);

            // Load the document from the saved file.
            Document loadedDoc = new Document(sourcePath);

            // Prepare the summary report.
            using (StringWriter reportWriter = new StringWriter())
            {
                reportWriter.WriteLine("Bookmark Summary Report");
                reportWriter.WriteLine("-----------------------");

                // Iterate through all bookmarks in the document.
                foreach (Bookmark bookmark in loadedDoc.Range.Bookmarks)
                {
                    string name = bookmark.Name;
                    string text = bookmark.Text; // Text inside the bookmark.
                    reportWriter.WriteLine($"Name: {name}");
                    reportWriter.WriteLine($"Text: {text}");
                    reportWriter.WriteLine();
                }

                // Write the report to a text file.
                string reportPath = "BookmarkReport.txt";
                File.WriteAllText(reportPath, reportWriter.ToString());

                // Optionally, output the report path to the console.
                Console.WriteLine($"Bookmark report generated at: {Path.GetFullPath(reportPath)}");
            }
        }
    }
}
