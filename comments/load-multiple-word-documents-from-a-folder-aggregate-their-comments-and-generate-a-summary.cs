using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    // Simple DTO to hold comment data.
    private class CommentInfo
    {
        public string SourceFile { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
        public DateTime DateTime { get; set; }
        public string Text { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Folder that will contain the sample documents.
        string tempFolder = Path.Combine(Directory.GetCurrentDirectory(), "TempDocs");
        Directory.CreateDirectory(tempFolder);

        // Create a few sample documents with comments.
        CreateSampleDocument(Path.Combine(tempFolder, "Doc1.docx"), "Author1", "A1", "First comment", DateTime.Now.AddDays(-2));
        CreateSampleDocument(Path.Combine(tempFolder, "Doc2.docx"), "Author2", "A2", "Second comment", DateTime.Now.AddDays(-1));
        CreateSampleDocument(Path.Combine(tempFolder, "Doc3.docx"), "Author3", "A3", "Third comment", DateTime.Now);

        // Aggregate comments from all documents in the folder.
        List<CommentInfo> allComments = new List<CommentInfo>();

        foreach (string filePath in Directory.GetFiles(tempFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Enumerate comment nodes safely.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment comment in comments)
            {
                // Ensure GetText() is not null.
                string commentText = comment.GetText()?.Trim() ?? string.Empty;

                allComments.Add(new CommentInfo
                {
                    SourceFile = Path.GetFileName(filePath),
                    Author = comment.Author,
                    DateTime = comment.DateTime,
                    Text = commentText
                });
            }
        }

        // Create a summary report document.
        Document report = new Document();
        DocumentBuilder builder = new DocumentBuilder(report);

        builder.Writeln("Comments Summary Report");
        builder.Writeln($"Generated on {DateTime.Now:u}");
        builder.Writeln();

        foreach (CommentInfo info in allComments)
        {
            builder.Writeln($"Source File : {info.SourceFile}");
            builder.Writeln($"Author      : {info.Author}");
            builder.Writeln($"Date/Time   : {info.DateTime:u}");
            builder.Writeln($"Comment Text: {info.Text}");
            builder.Writeln(); // Blank line between entries.
        }

        // Save the report to the working directory.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentsReport.docx");
        report.Save(reportPath);
    }

    // Helper method to create a document with a single comment.
    private static void CreateSampleDocument(string filePath, string author, string initials, string commentText, DateTime commentDate)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content.
        builder.Writeln($"This is the content of {Path.GetFileName(filePath)}.");

        // Create a comment and attach it to the current paragraph.
        Comment comment = new Comment(doc, author, initials, commentDate);
        comment.SetText(commentText);
        builder.CurrentParagraph.AppendChild(comment);

        // Ensure the comment has at least one paragraph (required for visibility).
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph.AppendChild(new Run(doc, commentText));

        doc.Save(filePath);
    }
}
