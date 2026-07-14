using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    // DTO to hold comment data together with its source file name.
    private class CommentInfo
    {
        public string SourceFile { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
        public DateTime DateTime { get; set; }
        public string Text { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare a temporary folder with sample Word documents.
        // -----------------------------------------------------------------
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "TempData");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create two sample documents, each containing a few comments.
        CreateSampleDocument(Path.Combine(inputDir, "Doc1.docx"), "Alice", "Bob");
        CreateSampleDocument(Path.Combine(inputDir, "Doc2.docx"), "Charlie", "Dana");

        // -----------------------------------------------------------------
        // 2. Load each document, enumerate its comments and collect data.
        // -----------------------------------------------------------------
        List<CommentInfo> allComments = new List<CommentInfo>();

        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Enumerate comments using the approved pattern.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment comment in comments)
            {
                // Guard against possible nulls.
                string commentText = comment?.GetText()?.Trim() ?? string.Empty;
                string author = comment?.Author ?? string.Empty;
                DateTime date = comment?.DateTime ?? DateTime.MinValue;

                allComments.Add(new CommentInfo
                {
                    SourceFile = Path.GetFileName(filePath),
                    Author = author,
                    DateTime = date,
                    Text = commentText
                });
            }
        }

        // -----------------------------------------------------------------
        // 3. Create a summary report document that lists all collected comments.
        // -----------------------------------------------------------------
        Document reportDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(reportDoc);

        builder.Writeln("Comments Summary Report");
        builder.Writeln($"Generated on: {DateTime.Now:O}");
        builder.Writeln();

        foreach (CommentInfo info in allComments)
        {
            builder.Writeln($"File   : {info.SourceFile}");
            builder.Writeln($"Author : {info.Author}");
            builder.Writeln($"Date   : {info.DateTime:O}");
            builder.Writeln($"Text   : {info.Text}");
            builder.Writeln(); // Blank line between entries.
        }

        // Save the report.
        string reportPath = Path.Combine(outputDir, "CommentsReport.docx");
        reportDoc.Save(reportPath);

        // -----------------------------------------------------------------
        // 4. Clean up temporary files (optional). Comment out if you want to inspect them.
        // -----------------------------------------------------------------
        // Directory.Delete(baseDir, true);
    }

    // -----------------------------------------------------------------
    // Helper method to create a sample document with two comments.
    // -----------------------------------------------------------------
    private static void CreateSampleDocument(string filePath, string author1, string author2)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("First paragraph with a comment.");
        Comment comment1 = new Comment(doc)
        {
            Author = author1,
            Initial = author1.Substring(0, 1).ToUpper(),
            DateTime = DateTime.Now
        };
        // Ensure the comment has at least one paragraph.
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph.AppendChild(new Run(doc, $"Comment by {author1}."));
        // Attach the comment to the first paragraph.
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment1);

        // Second paragraph with a comment.
        builder.Writeln("Second paragraph with another comment.");
        Comment comment2 = new Comment(doc)
        {
            Author = author2,
            Initial = author2.Substring(0, 1).ToUpper(),
            DateTime = DateTime.Now.AddMinutes(-5)
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph.AppendChild(new Run(doc, $"Comment by {author2}."));
        // The paragraph we just added is the last one in the body.
        Paragraph? secondPara = doc.FirstSection.Body.Paragraphs.Last() as Paragraph;
        if (secondPara != null)
        {
            secondPara.AppendChild(comment2);
        }

        // Save the sample document.
        doc.Save(filePath);
    }
}
