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
        // Create a new document and add some sample comments.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("First paragraph with a comment.");
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now.AddDays(-2));
        comment1.SetText("Review the first paragraph.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph with a comment.
        builder.Writeln("Second paragraph with another comment.");
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now.AddDays(-1));
        comment2.SetText("Check the data in this paragraph.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the document (optional, just to demonstrate saving).
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleComments.docx");
        doc.Save(docPath);

        // Extract comment metadata: author, date, and text.
        var commentInfos = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .Select(c => new CommentInfo
                              {
                                  Author = c.Author,
                                  Date = c.DateTime,
                                  Text = c.GetText().Trim()
                              })
                              .ToList();

        // Serialize the list to JSON.
        string json = JsonSerializer.Serialize(commentInfos, new JsonSerializerOptions { WriteIndented = true });

        // Write JSON to a file.
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "comments.json");
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
