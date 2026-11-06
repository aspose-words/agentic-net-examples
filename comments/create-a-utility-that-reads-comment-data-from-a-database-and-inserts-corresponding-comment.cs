using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

#nullable enable

namespace CommentInsertionExample
{
    // Simple record that mimics a row read from a database.
    public class CommentRecord
    {
        public int ParagraphIndex { get; set; }   // Zero‑based index of the paragraph to comment.
        public string Author { get; set; } = "";
        public string Initial { get; set; } = "";
        public DateTime DateTime { get; set; }
        public string Text { get; set; } = "";
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

            // 4. Insert comments based on the simulated data.
            foreach (CommentRecord record in commentData)
            {
                // Validate paragraph index.
                Paragraph? targetParagraph = GetParagraphByIndex(doc, record.ParagraphIndex);
                if (targetParagraph == null)
                    continue; // Skip invalid indices.

                // Create a new comment.
                Comment comment = new Comment(doc)
                {
                    Author = record.Author,
                    Initial = record.Initial,
                    DateTime = record.DateTime
                };

                // Add a paragraph and run so the comment has visible text.
                Paragraph commentParagraph = new Paragraph(doc);
                commentParagraph.AppendChild(new Run(doc, record.Text));
                comment.AppendChild(commentParagraph);

                // Append the comment to the target paragraph.
                targetParagraph.AppendChild(comment);
            }

            // 5. Save the resulting document.
            string resultPath = Path.Combine(outputDir, "filled.docx");
            doc.Save(resultPath);
        }

        // Creates a simple template document with three paragraphs.
        private static void CreateTemplateDocument(string path)
        {
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("First paragraph of the template.");
            builder.Writeln("Second paragraph of the template.");
            builder.Writeln("Third paragraph of the template.");

            template.Save(path);
        }

        // Returns a list of comment records that would normally come from a database.
        private static List<CommentRecord> GetSampleCommentData()
        {
            return new List<CommentRecord>
            {
                new CommentRecord
                {
                    ParagraphIndex = 0,
                    Author = "Alice",
                    Initial = "AL",
                    DateTime = DateTime.Now.AddMinutes(-15),
                    Text = "Please review the opening statement."
                },
                new CommentRecord
                {
                    ParagraphIndex = 1,
                    Author = "Bob",
                    Initial = "BO",
                    DateTime = DateTime.Now.AddMinutes(-10),
                    Text = "Consider adding more details here."
                },
                new CommentRecord
                {
                    ParagraphIndex = 2,
                    Author = "Carol",
                    Initial = "CA",
                    DateTime = DateTime.Now.AddMinutes(-5),
                    Text = "Check the formatting of this paragraph."
                }
            };
        }

        // Retrieves a paragraph by its zero‑based index safely.
        private static Paragraph? GetParagraphByIndex(Document doc, int index)
        {
            if (doc.FirstSection?.Body?.Paragraphs == null)
                return null;

            if (index < 0 || index >= doc.FirstSection.Body.Paragraphs.Count)
                return null;

            return doc.FirstSection.Body.Paragraphs[index];
        }
    }
}
