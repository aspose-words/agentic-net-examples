using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace CommentMetadataExport
{
    // Simple DTO for JSON serialization.
    public class CommentInfo
    {
        public string Author { get; set; } = string.Empty;
        public string Date { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // Create a sample document with a few comments.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First paragraph with a comment.
            builder.Writeln("First paragraph.");
            Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now.AddDays(-2));
            comment1.SetText("Review the introduction.");
            builder.CurrentParagraph.AppendChild(comment1);

            // Second paragraph with a comment.
            builder.Writeln("Second paragraph.");
            Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now.AddDays(-1));
            comment2.SetText("Check the data accuracy.");
            builder.CurrentParagraph.AppendChild(comment2);

            // Third paragraph without a comment.
            builder.Writeln("Third paragraph without comment.");

            // Save the sample document (optional, just to illustrate lifecycle).
            string docPath = Path.Combine(outputDir, "sample.docx");
            doc.Save(docPath);

            // Extract comment metadata.
            List<CommentInfo> commentInfos = doc
                .GetChildNodes(NodeType.Comment, true)
                .OfType<Comment>()
                .Select(c => new CommentInfo
                {
                    Author = c.Author,
                    Date = c.DateTime.ToString("o"), // ISO 8601 format.
                    Text = c.GetText().Trim()
                })
                .ToList();

            // Serialize to JSON.
            string json = JsonSerializer.Serialize(commentInfos, new JsonSerializerOptions { WriteIndented = true });

            // Write JSON to file.
            string jsonPath = Path.Combine(outputDir, "comments.json");
            File.WriteAllText(jsonPath, json);
        }
    }
}
