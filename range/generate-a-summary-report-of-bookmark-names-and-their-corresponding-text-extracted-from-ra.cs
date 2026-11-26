using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first bookmark with some text.
        builder.StartBookmark("FirstBookmark");
        builder.Writeln("This is the text inside the first bookmark.");
        builder.EndBookmark("FirstBookmark");

        // Insert second bookmark with some text.
        builder.StartBookmark("SecondBookmark");
        builder.Writeln("Second bookmark contains different text.");
        builder.EndBookmark("SecondBookmark");

        // Insert third bookmark with a single line of text.
        builder.StartBookmark("ThirdBookmark");
        builder.Write("Third bookmark text.");
        builder.EndBookmark("ThirdBookmark");

        // Save the sample document (optional, demonstrates loading later if needed).
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        doc.Save(docPath);

        // Load the document to ensure the loading rule is exercised.
        Document loadedDoc = new Document(docPath);

        // Prepare the report content.
        using (StringWriter reportWriter = new StringWriter())
        {
            // Iterate through all bookmarks in the document's range.
            foreach (Bookmark bookmark in loadedDoc.Range.Bookmarks)
            {
                // Extract bookmark name and its inner text.
                string name = bookmark.Name;
                string text = bookmark.Text;

                // Write a line to the report.
                reportWriter.WriteLine($"Bookmark: {name}, Text: {text}");
            }

            // Save the report to a text file.
            string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "BookmarkReport.txt");
            File.WriteAllText(reportPath, reportWriter.ToString());

            // Optionally, write the report to console.
            Console.WriteLine(reportWriter.ToString());
        }
    }
}
