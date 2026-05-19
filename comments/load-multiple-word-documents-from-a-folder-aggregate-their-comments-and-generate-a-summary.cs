using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents containing comments.
        CreateSampleDocument(
            Path.Combine(inputDir, "Doc1.docx"),
            "First document",
            new[]
            {
                new CommentInfo("Alice", "AL", "First comment in Doc1."),
                new CommentInfo("Bob", "BO", "Second comment in Doc1.")
            });

        CreateSampleDocument(
            Path.Combine(inputDir, "Doc2.docx"),
            "Second document",
            new[]
            {
                new CommentInfo("Charlie", "CH", "Only comment in Doc2.")
            });

        // Aggregate comments from all documents in the folder.
        List<AggregatedComment> allComments = new List<AggregatedComment>();

        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(filePath);
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                .OfType<Comment>()
                .Select(c => new AggregatedComment
                {
                    DocumentName = Path.GetFileName(filePath),
                    Author = c.Author,
                    Date = c.DateTime,
                    Text = c.GetText().Trim()
                })
                .ToList();

            allComments.AddRange(comments);
        }

        // Build a summary report document.
        Document report = new Document();
        DocumentBuilder builder = new DocumentBuilder(report);
        builder.Writeln("Comments Summary Report");
        builder.Writeln($"Generated on {DateTime.Now:O}");
        builder.Writeln();

        foreach (var comment in allComments)
        {
            builder.Writeln($"Document: {comment.DocumentName}");
            builder.Writeln($"Author: {comment.Author}");
            builder.Writeln($"Date: {comment.Date:O}");
            builder.Writeln($"Text: {comment.Text}");
            builder.Writeln();
        }

        string reportPath = Path.Combine(outputDir, "CommentsReport.docx");
        report.Save(reportPath);
    }

    // Creates a Word document with the specified title and a set of comments.
    private static void CreateSampleDocument(string filePath, string title, CommentInfo[] commentInfos)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(title);

        foreach (var info in commentInfos)
        {
            // Add a paragraph that will host the comment.
            builder.Writeln($"Paragraph for comment by {info.Author}.");

            // Build the comment node.
            Comment comment = new Comment(doc)
            {
                Author = info.Author,
                Initial = info.Initial,
                DateTime = DateTime.Now
            };
            comment.AppendChild(new Paragraph(doc));
            comment.FirstParagraph.AppendChild(new Run(doc, info.Text));

            // Attach the comment to the current paragraph.
            Paragraph para = builder.CurrentParagraph;
            para.AppendChild(comment);
        }

        doc.Save(filePath);
    }

    // Simple DTO used when creating sample comments.
    private class CommentInfo
    {
        public string Author { get; }
        public string Initial { get; }
        public string Text { get; }

        public CommentInfo(string author, string initial, string text)
        {
            Author = author;
            Initial = initial;
            Text = text;
        }
    }

    // DTO representing a comment extracted from a source document.
    private class AggregatedComment
    {
        public string DocumentName { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
        public DateTime Date { get; set; }
        public string Text { get; set; } = string.Empty;
    }
}
