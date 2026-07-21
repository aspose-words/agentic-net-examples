using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeCommentsDemo
{
    public class Program
    {
        // DTO to hold comment information together with its source file name.
        private class AggregatedComment
        {
            public string SourceFile { get; set; } = string.Empty;
            public string Author { get; set; } = string.Empty;
            public DateTime DateTime { get; set; }
            public string Text { get; set; } = string.Empty;
        }

        public static void Main()
        {
            // 1. Prepare a temporary folder for the demo files.
            string demoFolder = Path.Combine(Path.GetTempPath(), "AsposeCommentsDemo");
            Directory.CreateDirectory(demoFolder);

            // 2. Create a few sample Word documents with comments.
            CreateSampleDocuments(demoFolder);

            // 3. Load all documents from the folder and aggregate their comments.
            List<AggregatedComment> allComments = LoadAndAggregateComments(demoFolder);

            // 4. Generate a summary report document containing the aggregated comment data.
            GenerateReport(demoFolder, allComments);

            // 5. Inform the user (via console) where the report was saved.
            Console.WriteLine($"Comments summary report generated at: {Path.Combine(demoFolder, "CommentsReport.docx")}");
        }

        private static void CreateSampleDocuments(string folderPath)
        {
            // Create three sample documents.
            for (int i = 1; i <= 3; i++)
            {
                Document doc = new Document();
                // Ensure the document has at least one paragraph.
                doc.EnsureMinimum();

                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln($"This is the content of sample document {i}.");

                // Add a comment to the current paragraph.
                string author = $"Author{i}";
                string initials = $"A{i}";
                DateTime commentDate = DateTime.Now.AddDays(-i);
                Comment comment = new Comment(doc, author, initials, commentDate);
                // Append the comment node to the paragraph.
                builder.CurrentParagraph.AppendChild(comment);
                // Add a paragraph inside the comment to hold the comment text.
                Paragraph commentParagraph = new Paragraph(doc);
                commentParagraph.AppendChild(new Run(doc, $"This is comment {i} text."));
                comment.AppendChild(commentParagraph);

                // Save the document.
                string fileName = Path.Combine(folderPath, $"Sample{i}.docx");
                doc.Save(fileName);
            }
        }

        private static List<AggregatedComment> LoadAndAggregateComments(string folderPath)
        {
            var aggregated = new List<AggregatedComment>();

            // Get all .docx files in the folder.
            string[] docFiles = Directory.GetFiles(folderPath, "*.docx");
            foreach (string filePath in docFiles)
            {
                // Load the document.
                Document doc = new Document(filePath);

                // Enumerate all comment nodes in the document.
                var comments = doc.GetChildNodes(NodeType.Comment, true)
                                  .OfType<Comment>()
                                  .ToList();

                foreach (Comment c in comments)
                {
                    // Guard against possible nulls (unlikely but satisfies nullable rules).
                    string author = c.Author ?? string.Empty;
                    DateTime date = c.DateTime;
                    string text = c.GetText()?.Trim() ?? string.Empty;

                    aggregated.Add(new AggregatedComment
                    {
                        SourceFile = Path.GetFileName(filePath),
                        Author = author,
                        DateTime = date,
                        Text = text
                    });
                }
            }

            return aggregated;
        }

        private static void GenerateReport(string folderPath, List<AggregatedComment> comments)
        {
            Document report = new Document();
            // Ensure the report has a body to write into.
            report.EnsureMinimum();

            DocumentBuilder builder = new DocumentBuilder(report);
            builder.Writeln("Comments Summary Report");
            builder.Writeln($"Generated on: {DateTime.Now:O}");
            builder.Writeln();

            if (comments.Count == 0)
            {
                builder.Writeln("No comments were found in the processed documents.");
            }
            else
            {
                // Group comments by source file for clearer organization.
                var commentsByFile = comments.GroupBy(c => c.SourceFile);
                foreach (var fileGroup in commentsByFile)
                {
                    builder.Writeln($"Document: {fileGroup.Key}");
                    foreach (AggregatedComment ac in fileGroup)
                    {
                        builder.Writeln($"  Author : {ac.Author}");
                        builder.Writeln($"  Date   : {ac.DateTime:O}");
                        builder.Writeln($"  Text   : {ac.Text}");
                        builder.Writeln();
                    }
                }
            }

            // Save the report in the same demo folder.
            string reportPath = Path.Combine(folderPath, "CommentsReport.docx");
            report.Save(reportPath);
        }
    }
}
