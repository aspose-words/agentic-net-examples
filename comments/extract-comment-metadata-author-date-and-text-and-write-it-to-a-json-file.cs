using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph with a comment.");
        builder.Writeln("Second paragraph without a comment.");

        // Create first comment.
        Comment comment1 = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now.AddDays(-1)
        };
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));

        // Attach the comment to the first paragraph.
        Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        firstParagraph?.AppendChild(comment1);

        // Create second comment.
        Comment comment2 = new Comment(doc)
        {
            Author = "Jane Smith",
            Initial = "JS",
            DateTime = DateTime.Now
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Consider adding more details here."));

        // Attach the second comment to the second paragraph.
        Paragraph? secondParagraph = doc.FirstSection?.Body?.Paragraphs[1];
        secondParagraph?.AppendChild(comment2);

        // Extract comment metadata.
        List<CommentInfo> commentInfos = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Select(c => new CommentInfo
            {
                Author = c.Author,
                Date = c.DateTime,
                Text = c.GetText().Trim()
            })
            .ToList();

        // Serialize to JSON.
        string json = JsonSerializer.Serialize(commentInfos, new JsonSerializerOptions { WriteIndented = true });

        // Ensure output directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "comments.json");
        File.WriteAllText(outputPath, json);
    }

    // Simple DTO for JSON serialization.
    private class CommentInfo
    {
        public string Author { get; set; } = string.Empty;
        public DateTime Date { get; set; }
        public string Text { get; set; } = string.Empty;
    }
}
