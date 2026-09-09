using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare folders for input documents and output report.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents with comments.
        CreateSampleDocument(Path.Combine(inputDir, "Doc1.docx"), "Alice", "First comment in Doc1.");
        CreateSampleDocument(Path.Combine(inputDir, "Doc2.docx"), "Bob", "Second comment in Doc2.");
        CreateSampleDocument(Path.Combine(inputDir, "Doc3.docx"), "Charlie", "Third comment in Doc3.");

        // Aggregate comments from all documents in the folder.
        List<AggregatedComment> allComments = new List<AggregatedComment>();
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(filePath);
            // Enumerate comments safely.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment c in comments)
            {
                allComments.Add(new AggregatedComment
                {
                    SourceDocument = Path.GetFileName(filePath),
                    Author = c.Author,
                    Date = c.DateTime,
                    Text = c.GetText().Trim()
                });
            }
        }

        // Create a summary report document.
        Document report = new Document();
        DocumentBuilder builder = new DocumentBuilder(report);
        builder.Writeln("Comments Summary Report");
        builder.Writeln($"Generated on: {DateTime.Now}");
        builder.Writeln();

        foreach (AggregatedComment ac in allComments)
        {
            builder.Writeln($"Document: {ac.SourceDocument}");
            builder.Writeln($"Author: {ac.Author}");
            builder.Writeln($"Date: {ac.Date}");
            builder.Writeln($"Comment: {ac.Text}");
            builder.Writeln(); // Blank line between entries.
        }

        // Save the report.
        string reportPath = Path.Combine(outputDir, "CommentsReport.docx");
        report.Save(reportPath);
    }

    // Helper method to create a simple document with a single comment.
    private static void CreateSampleDocument(string filePath, string author, string commentText)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will be commented.
        builder.Writeln($"This is a sample paragraph for {author}.");

        // Create a comment anchored to the current paragraph.
        Comment comment = new Comment(doc, author, author.Substring(0, 1), DateTime.Now);
        // Append a paragraph inside the comment to hold its text.
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph.AppendChild(new Run(doc, commentText));

        // Attach the comment to the paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document.
        doc.Save(filePath);
    }

    // Simple DTO to hold aggregated comment information.
    private class AggregatedComment
    {
        public string SourceDocument { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
        public DateTime Date { get; set; }
        public string Text { get; set; } = string.Empty;
    }
}
