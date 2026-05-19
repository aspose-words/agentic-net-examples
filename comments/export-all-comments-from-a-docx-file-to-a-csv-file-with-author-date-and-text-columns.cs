using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class ExportCommentsToCsv
{
    public static void Main()
    {
        // Create an output folder in the current working directory.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputFolder);

        string docPath = Path.Combine(outputFolder, "sample.docx");
        string csvPath = Path.Combine(outputFolder, "comments.csv");

        // -----------------------------------------------------------------
        // Create a sample DOCX file with a few comments.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("First paragraph.");
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("This is Alice's comment."); // SetText creates a paragraph inside the comment.
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

        var comments = loadedDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        using (StreamWriter writer = new StreamWriter(csvPath, false))
        {
            // Write CSV header.
            writer.WriteLine("Author,Date,Text");

            foreach (Comment c in comments)
            {
                // Get the comment text, trim whitespace, and escape double quotes.
                string rawText = c.GetText()?.Trim() ?? string.Empty;
                string escapedText = rawText.Replace("\"", "\"\"");

                // Format the comment date in ISO 8601.
                string isoDate = c.DateTime.ToString("o");

                // Write a CSV line, quoting each field.
                writer.WriteLine($"\"{c.Author}\",\"{isoDate}\",\"{escapedText}\"");
            }
        }

        Console.WriteLine($"Exported {comments.Count} comment(s) to \"{csvPath}\".");
    }
}
