using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the sample DOCX file.
        string docPath = Path.Combine(outputDir, "Sample.docx");

        // -----------------------------------------------------------------
        // Create a sample document with a few comments.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("First paragraph.");
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("This is Alice's comment.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph with another comment.
        builder.Writeln("Second paragraph.");
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now.AddMinutes(-5));
        comment2.SetText("Bob's comment goes here.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the sample document.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // Load the document and export its comments to CSV.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Enumerate all comment nodes in the document.
        var comments = loadedDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        // Path for the CSV output.
        string csvPath = Path.Combine(outputDir, "Comments.csv");

        // Write CSV header and comment rows.
        using (var writer = new StreamWriter(csvPath))
        {
            writer.WriteLine("Author,Date,Text");

            foreach (Comment c in comments)
            {
                // Escape double quotes in CSV fields.
                string author = (c.Author ?? string.Empty).Replace("\"", "\"\"");
                string date = c.DateTime.ToString("o"); // ISO 8601 format.
                string text = (c.GetText() ?? string.Empty).Trim().Replace("\"", "\"\"");

                writer.WriteLine($"\"{author}\",\"{date}\",\"{text}\"");
            }
        }

        // Indicate completion (optional console output).
        Console.WriteLine($"Exported {comments.Count} comment(s) to \"{csvPath}\".");
    }
}
