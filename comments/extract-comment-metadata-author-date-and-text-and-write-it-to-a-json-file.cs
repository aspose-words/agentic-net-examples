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
        // Create a sample document with a couple of comments.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph and its comment.
        builder.Writeln("First paragraph.");
        Paragraph firstPara = doc.FirstSection.Body.FirstParagraph;
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now.AddDays(-1));
        comment1.SetText("Review this paragraph.");
        firstPara.AppendChild(new CommentRangeStart(doc, comment1.Id));
        firstPara.AppendChild(new Run(doc, "Commented text."));
        firstPara.AppendChild(new CommentRangeEnd(doc, comment1.Id));
        firstPara.AppendChild(comment1);

        // Second paragraph and its comment.
        builder.Writeln("Second paragraph.");
        Paragraph secondPara = doc.FirstSection.Body.LastParagraph;
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Check spelling.");
        secondPara.AppendChild(new CommentRangeStart(doc, comment2.Id));
        secondPara.AppendChild(new Run(doc, "Another commented text."));
        secondPara.AppendChild(new CommentRangeEnd(doc, comment2.Id));
        secondPara.AppendChild(comment2);

        // Save the sample document (optional, for verification).
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        doc.Save(docPath);

        // Extract comment metadata.
        List<CommentInfo> commentInfos = doc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Select(c => new CommentInfo
            {
                Author = c.Author,
                Date = c.DateTime,
                Text = c.GetText().Trim()
            })
            .ToList();

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Serialize to JSON with indentation.
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
