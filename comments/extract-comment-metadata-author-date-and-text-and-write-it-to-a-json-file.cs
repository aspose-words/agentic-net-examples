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
        // Create a sample document with comments.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph.
        builder.Writeln("First paragraph.");

        // Comment for the first paragraph.
        Comment comment1 = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now.AddDays(-1)
        };
        Paragraph commentPara1 = new Paragraph(doc);
        commentPara1.AppendChild(new Run(doc, "Please review this paragraph."));
        comment1.AppendChild(commentPara1);
        doc.FirstSection.Body.FirstParagraph?.AppendChild(comment1);

        // Second paragraph.
        builder.Writeln("Second paragraph.");

        // Comment for the second paragraph.
        Comment comment2 = new Comment(doc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now
        };
        Paragraph commentPara2 = new Paragraph(doc);
        commentPara2.AppendChild(new Run(doc, "Check the data here."));
        comment2.AppendChild(commentPara2);
        Paragraph? secondParagraph = doc.FirstSection.Body.LastParagraph;
        secondParagraph?.AppendChild(comment2);

        // Save the sample document (optional, just for reference).
        doc.Save("sample.docx");

        // Extract comment metadata.
        List<Comment> commentNodes = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        List<CommentInfo> commentInfos = new List<CommentInfo>();
        foreach (Comment c in commentNodes)
        {
            string author = c.Author ?? string.Empty;
            string date = c.DateTime.ToString("o");
            string text = c.GetText()?.Trim() ?? string.Empty;
            commentInfos.Add(new CommentInfo(author, date, text));
        }

        // Serialize metadata to JSON.
        JsonSerializerOptions jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(commentInfos, jsonOptions);

        // Write JSON to file.
        File.WriteAllText("comments.json", json);
    }

    private record CommentInfo(string Author, string DateTime, string Text);
}
