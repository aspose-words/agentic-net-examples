using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a sample document with a couple of comments.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("This is the first paragraph.");
        Comment comment1 = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now.AddDays(-1)
        };
        // Add visible text to the comment.
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "Please review this first paragraph."));
        // Attach the comment to the first paragraph.
        doc.FirstSection?.Body?.FirstParagraph?.AppendChild(comment1);

        // Second paragraph with a comment.
        builder.Writeln("This is the second paragraph.");
        Comment comment2 = new Comment(doc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Check the data in this paragraph."));
        // Attach the comment to the second paragraph (the last paragraph added).
        Paragraph? secondParagraph = doc.FirstSection?.Body?.LastParagraph;
        secondParagraph?.AppendChild(comment2);

        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Extract comment metadata.
        List<CommentInfo> commentInfos = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Select(c => new CommentInfo
            {
                Author = c.Author,
                Date = c.DateTime,
                Text = c.GetText()?.Trim() ?? string.Empty
            })
            .ToList();

        // Serialize to JSON.
        JsonSerializerOptions jsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true
        };
        string json = JsonSerializer.Serialize(commentInfos, jsonOptions);

        // Write JSON to file.
        string jsonPath = Path.Combine(outputDir, "comments.json");
        File.WriteAllText(jsonPath, json);
    }

    // Simple DTO for JSON serialization.
    private class CommentInfo
    {
        public string Author { get; set; } = string.Empty;
        public DateTime Date { get; set; }
        public string Text { get; set; } = string.Empty;
    }
}
