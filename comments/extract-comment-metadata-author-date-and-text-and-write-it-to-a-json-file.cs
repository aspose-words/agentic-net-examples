using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Aspose.Words;

public class Program
{
    // DTO for JSON serialization of comment metadata.
    private class CommentInfo
    {
        public string Author { get; set; } = string.Empty;
        public string DateTime { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a new document and add sample comments.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph.
        builder.Writeln("This is the first paragraph.");

        // Create first comment.
        Comment comment1 = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now.AddDays(-1)
        };
        // A comment must contain at least one paragraph and one run.
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));
        // Attach the comment to the first paragraph.
        doc.FirstSection?.Body?.FirstParagraph?.AppendChild(comment1);

        // Second paragraph.
        builder.Writeln("This is the second paragraph.");

        // Create second comment.
        Comment comment2 = new Comment(doc)
        {
            Author = "Jane Smith",
            Initial = "JS",
            DateTime = DateTime.Now
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Consider rephrasing this sentence."));
        // Attach the comment to the second paragraph (the last paragraph added by the builder).
        Paragraph? secondParagraph = doc.FirstSection?.Body?.LastParagraph;
        secondParagraph?.AppendChild(comment2);

        // -----------------------------------------------------------------
        // 2. Extract comment metadata (author, date, text).
        // -----------------------------------------------------------------
        List<CommentInfo> commentData = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Select(c => new CommentInfo
            {
                Author = c.Author ?? string.Empty,
                DateTime = c.DateTime.ToString("o"),
                Text = c.GetText().Trim()
            })
            .ToList();

        // -----------------------------------------------------------------
        // 3. Serialize the metadata to JSON.
        // -----------------------------------------------------------------
        JsonSerializerOptions options = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(commentData, options);

        // -----------------------------------------------------------------
        // 4. Write JSON to a deterministic file in the current directory.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "comments.json");
        File.WriteAllText(outputPath, json);

        // Inform the user where the file was written.
        Console.WriteLine($"Comment metadata written to: {outputPath}");
    }
}
