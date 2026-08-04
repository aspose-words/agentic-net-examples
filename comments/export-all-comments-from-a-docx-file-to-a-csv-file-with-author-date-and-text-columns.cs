using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class ExportCommentsToCsv
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the sample DOCX file.
        string docPath = Path.Combine(outputDir, "Sample.docx");

        // Create a sample document with comments.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("First paragraph.");
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("First comment.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph with another comment.
        builder.Writeln("Second paragraph.");
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now.AddMinutes(-5));
        comment2.SetText("Second comment, includes a comma.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the sample document.
        doc.Save(docPath);

        // Load the document to demonstrate reading comments.
        Document loadedDoc = new Document(docPath);

        // Enumerate all comments in the document.
        var comments = loadedDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        // Path for the CSV output.
        string csvPath = Path.Combine(outputDir, "Comments.csv");

        // Write comments to CSV with columns: Author, Date, Text.
        using (var writer = new StreamWriter(csvPath))
        {
            writer.WriteLine("Author,Date,Text");
            foreach (Comment c in comments)
            {
                // Ensure CSV fields are properly escaped.
                string author = (c.Author ?? string.Empty).Replace("\"", "\"\"");
                string date = c.DateTime.ToString("o"); // ISO 8601 format.
                string text = c.GetText().Trim().Replace("\"", "\"\"");

                writer.WriteLine($"\"{author}\",\"{date}\",\"{text}\"");
            }
        }

        // Indicate completion (no interactive input required).
        Console.WriteLine($"Exported {comments.Count} comment(s) to \"{csvPath}\".");
    }
}
