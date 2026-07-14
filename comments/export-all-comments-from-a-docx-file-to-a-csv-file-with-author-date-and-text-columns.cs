using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class ExportCommentsToCsv
{
    public static void Main()
    {
        // Prepare file paths.
        string workingDir = Directory.GetCurrentDirectory();
        string sampleDocPath = Path.Combine(workingDir, "sample.docx");
        string outputDir = Path.Combine(workingDir, "output");
        string csvPath = Path.Combine(outputDir, "comments.csv");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document with a few comments.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("This is the first paragraph.");
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now.AddDays(-2));
        comment1.SetText("Review the opening sentence.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph with a comment.
        builder.Writeln("Second paragraph contains important data.");
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now.AddDays(-1));
        comment2.SetText("Check the figures, especially the last one.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Third paragraph without a comment.
        builder.Writeln("No comment here.");

        // Save the sample document.
        doc.Save(sampleDocPath);

        // -----------------------------------------------------------------
        // 2. Load the document (simulating a real‑world scenario).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sampleDocPath);

        // -----------------------------------------------------------------
        // 3. Enumerate all top‑level comments.
        // -----------------------------------------------------------------
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        // -----------------------------------------------------------------
        // 4. Export comments to CSV (Author, Date, Text).
        // -----------------------------------------------------------------
        using (var writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8))
        {
            // Write CSV header.
            writer.WriteLine("Author,Date,Text");

            foreach (Comment c in comments)
            {
                string author = c.Author ?? string.Empty;
                string date = c.DateTime.ToString("o"); // ISO 8601 format.
                string text = c.GetText()?.Trim() ?? string.Empty;

                writer.WriteLine($"{EscapeCsv(author)},{EscapeCsv(date)},{EscapeCsv(text)}");
            }
        }

        Console.WriteLine($"Exported {comments.Count} comment(s) to \"{csvPath}\".");
    }

    // Helper method to escape a CSV field according to RFC 4180.
    private static string EscapeCsv(string field)
    {
        if (field.Contains('"'))
            field = field.Replace("\"", "\"\"");

        bool mustQuote = field.Contains(',') || field.Contains('"') || field.Contains('\r') || field.Contains('\n');
        return mustQuote ? $"\"{field}\"" : field;
    }
}
