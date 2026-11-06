using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class ExportCommentsToCsv
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample DOCX file with comments.
        string docPath = Path.Combine(outputDir, "sample.docx");
        CreateSampleDocument(docPath);

        // Load the document.
        Document doc = new Document(docPath);

        // Enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Prepare CSV file.
        string csvPath = Path.Combine(outputDir, "comments.csv");
        using (var writer = new StreamWriter(csvPath))
        {
            // Write CSV header.
            writer.WriteLine("Author,Date,Text");

            // Write each comment as a CSV line.
            foreach (Comment comment in comments)
            {
                string author = comment.Author?.Replace("\"", "\"\"") ?? string.Empty;
                string date = comment.DateTime.ToString("o"); // ISO 8601 format.
                string text = comment.GetText()?.Trim().Replace("\"", "\"\"") ?? string.Empty;

                // Enclose fields containing commas or quotes in double quotes.
                if (author.Contains(',') || author.Contains('\"')) author = $"\"{author}\"";
                if (text.Contains(',') || text.Contains('\"')) text = $"\"{text}\"";

                writer.WriteLine($"{author},{date},{text}");
            }
        }

        // Indicate completion.
        Console.WriteLine($"Exported {comments.Count} comment(s) to \"{csvPath}\".");
    }

    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("This is the first paragraph.");
        Comment comment1 = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now.AddDays(-2)
        };
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "Review the introduction."));
        doc.FirstSection?.Body?.FirstParagraph?.AppendChild(comment1);

        // Second paragraph with a comment.
        builder.Writeln("This is the second paragraph.");
        Comment comment2 = new Comment(doc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now.AddDays(-1)
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Check the data accuracy."));
        doc.FirstSection?.Body?.LastParagraph?.AppendChild(comment2);

        // Save the sample document.
        doc.Save(path);
    }
}
