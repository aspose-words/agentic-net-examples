using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("This is the first paragraph.");
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now.AddDays(-2));
        comment1.SetText("Review the first paragraph.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph with a comment.
        builder.Writeln("This is the second paragraph.");
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now.AddDays(-1));
        comment2.SetText("Check the second paragraph for accuracy.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the sample document (optional, for verification).
        doc.Save("sample.docx");

        // Extract comment metadata.
        List<CommentInfo> commentInfos = doc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Select(c => new CommentInfo
            {
                Author = c.Author,
                Date = c.DateTime,
                Text = c.GetText()?.Trim() ?? string.Empty
            })
            .ToList();

        // Serialize to JSON with indentation.
        string json = JsonSerializer.Serialize(commentInfos, new JsonSerializerOptions { WriteIndented = true });

        // Write JSON to a file in the working directory.
        File.WriteAllText("comments.json", json);
    }

    // Simple DTO for JSON serialization.
    private class CommentInfo
    {
        public string Author { get; set; } = string.Empty;
        public DateTime Date { get; set; }
        public string Text { get; set; } = string.Empty;
    }
}
