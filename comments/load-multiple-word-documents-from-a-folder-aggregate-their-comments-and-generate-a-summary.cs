using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace CommentAggregator
{
    public class Program
    {
        public static void Main()
        {
            // Create a temporary folder for sample documents.
            string tempFolder = Path.Combine(Directory.GetCurrentDirectory(), "TempDocs");
            Directory.CreateDirectory(tempFolder);

            // Generate sample Word documents with comments.
            CreateSampleDocument(Path.Combine(tempFolder, "Document1.docx"), "First document content.", "Alice", "AL", "Comment on first document.");
            CreateSampleDocument(Path.Combine(tempFolder, "Document2.docx"), "Second document content.", "Bob", "BO", "Comment on second document.");

            // Aggregate comments from all documents in the folder.
            List<CommentInfo> aggregatedComments = new List<CommentInfo>();

            foreach (string filePath in Directory.GetFiles(tempFolder, "*.docx"))
            {
                Document doc = new Document(filePath);
                var comments = doc.GetChildNodes(NodeType.Comment, true)
                                  .OfType<Comment>()
                                  .ToList();

                foreach (Comment comment in comments)
                {
                    string text = comment.GetText()?.Trim() ?? string.Empty;
                    aggregatedComments.Add(new CommentInfo
                    {
                        SourceFile = Path.GetFileName(filePath),
                        Author = comment.Author,
                        DateTime = comment.DateTime,
                        Text = text
                    });
                }
            }

            // Create a summary report document.
            Document reportDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(reportDoc);

            builder.Writeln("Comments Summary Report");
            builder.Writeln($"Generated on: {DateTime.Now:O}");
            builder.Writeln();

            foreach (CommentInfo info in aggregatedComments)
            {
                builder.Writeln($"Source File: {info.SourceFile}");
                builder.Writeln($"Author     : {info.Author}");
                builder.Writeln($"Date       : {info.DateTime:O}");
                builder.Writeln($"Text       : {info.Text}");
                builder.Writeln(); // Blank line between entries.
            }

            // Save the report.
            string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentsReport.docx");
            reportDoc.Save(reportPath);
        }

        // Helper method to create a sample document with a single comment.
        private static void CreateSampleDocument(string filePath, string bodyText, string author, string initials, string commentText)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write main content.
            builder.Writeln(bodyText);

            // Create a comment node.
            Comment comment = new Comment(doc)
            {
                Author = author,
                Initial = initials,
                DateTime = DateTime.Now
            };

            // Add a paragraph and run inside the comment to hold its text.
            comment.AppendChild(new Paragraph(doc));
            Paragraph? commentParagraph = comment.FirstParagraph;
            if (commentParagraph != null)
            {
                commentParagraph.AppendChild(new Run(doc, commentText));
            }

            // Attach the comment to the first paragraph of the document body.
            Paragraph? targetParagraph = doc.FirstSection?.Body?.FirstParagraph;
            if (targetParagraph != null)
            {
                targetParagraph.AppendChild(comment);
            }

            // Save the document.
            doc.Save(filePath);
        }

        // Simple DTO to hold comment information for the report.
        private class CommentInfo
        {
            public string SourceFile { get; set; } = string.Empty;
            public string Author { get; set; } = string.Empty;
            public DateTime DateTime { get; set; }
            public string Text { get; set; } = string.Empty;
        }
    }
}
