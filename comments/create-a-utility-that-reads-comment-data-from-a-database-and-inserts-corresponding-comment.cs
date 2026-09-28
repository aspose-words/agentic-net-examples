using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;

namespace CommentInsertionExample
{
    // Simple record representing a comment retrieved from a database.
    public class CommentRecord
    {
        public string Author { get; set; } = "";
        public string Initial { get; set; } = "";
        public DateTime DateTime { get; set; }
        public string Text { get; set; } = "";
        // Zero‑based index of the paragraph where the comment should be attached.
        public int ParagraphIndex { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a template document with a few paragraphs.
            string templatePath = Path.Combine(outputDir, "template.docx");
            CreateTemplateDocument(templatePath);

            // 2. Simulate reading comment data from a database.
            List<CommentRecord> commentData = GetSampleCommentData();

            // 3. Load the template document.
            Document doc = new Document(templatePath);

            // 4. Insert comments into the document according to the simulated data.
            InsertCommentsIntoDocument(doc, commentData);

            // 5. Save the resulting document.
            string resultPath = Path.Combine(outputDir, "document-with-comments.docx");
            doc.Save(resultPath);

            // 6. Enumerate and display inserted comments (optional verification).
            EnumerateComments(doc);
        }

        private static void CreateTemplateDocument(string path)
        {
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Paragraph 1: Introduction.");
            builder.Writeln("Paragraph 2: Details.");
            builder.Writeln("Paragraph 3: Conclusion.");

            template.Save(path);
        }

        private static List<CommentRecord> GetSampleCommentData()
        {
            // In a real scenario this data would come from a database query.
            return new List<CommentRecord>
            {
                new CommentRecord
                {
                    Author = "Alice",
                    Initial = "AL",
                    DateTime = DateTime.Now.AddMinutes(-30),
                    Text = "Please review the introduction.",
                    ParagraphIndex = 0 // First paragraph
                },
                new CommentRecord
                {
                    Author = "Bob",
                    Initial = "BO",
                    DateTime = DateTime.Now.AddMinutes(-20),
                    Text = "Add more technical details here.",
                    ParagraphIndex = 1 // Second paragraph
                },
                new CommentRecord
                {
                    Author = "Carol",
                    Initial = "CA",
                    DateTime = DateTime.Now.AddMinutes(-10),
                    Text = "Conclusion looks good.",
                    ParagraphIndex = 2 // Third paragraph
                }
            };
        }

        private static void InsertCommentsIntoDocument(Document doc, List<CommentRecord> records)
        {
            // Retrieve all paragraphs in the main story.
            List<Paragraph> paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                                            .OfType<Paragraph>()
                                            .ToList();

            foreach (CommentRecord record in records)
            {
                // Guard against an invalid paragraph index.
                if (record.ParagraphIndex < 0 || record.ParagraphIndex >= paragraphs.Count)
                    continue;

                Paragraph targetParagraph = paragraphs[record.ParagraphIndex];

                // Create a new comment node.
                Comment comment = new Comment(doc)
                {
                    Author = record.Author,
                    Initial = record.Initial,
                    DateTime = record.DateTime
                };

                // Add visible text to the comment body.
                Paragraph commentParagraph = new Paragraph(doc);
                Run commentRun = new Run(doc, record.Text);
                commentParagraph.AppendChild(commentRun);
                comment.AppendChild(commentParagraph);

                // Attach the comment to the target paragraph.
                // The comment is added as a child of the paragraph node.
                targetParagraph.AppendChild(comment);
            }
        }

        private static void EnumerateComments(Document doc)
        {
            List<Comment> comments = doc.GetChildNodes(NodeType.Comment, true)
                                        .OfType<Comment>()
                                        .ToList();

            foreach (Comment c in comments)
            {
                string text = c.GetText().Trim();
                Console.WriteLine($"{c.Author} ({c.Initial}) at {c.DateTime:u}: {text}");
            }
        }
    }
}
