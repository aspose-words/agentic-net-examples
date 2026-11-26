using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace CommentInsertionExample
{
    // Simple POCO representing a comment record that might come from a database.
    public class CommentData
    {
        public string Author { get; set; } = "";
        public string Initial { get; set; } = "";
        public DateTime DateTime { get; set; }
        public string Text { get; set; } = "";
        // Zero‑based index of the paragraph in the main body to which the comment will be attached.
        public int ParagraphIndex { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Simulate reading comment data from a database.
            List<CommentData> commentDataList = GetSampleCommentData();

            // 2. Create a simple template document with a few paragraphs.
            string templatePath = "template.docx";
            CreateTemplateDocument(templatePath);

            // 3. Load the template document.
            Document doc = new Document(templatePath);

            // 4. Insert comments based on the simulated database records.
            InsertCommentsIntoDocument(doc, commentDataList);

            // 5. Save the resulting document.
            string outputPath = "output-with-comments.docx";
            doc.Save(outputPath);
        }

        // Simulated database retrieval.
        private static List<CommentData> GetSampleCommentData()
        {
            return new List<CommentData>
            {
                new CommentData
                {
                    Author = "Alice",
                    Initial = "AL",
                    DateTime = DateTime.Now.AddDays(-2),
                    Text = "Please review this opening paragraph.",
                    ParagraphIndex = 0
                },
                new CommentData
                {
                    Author = "Bob",
                    Initial = "BO",
                    DateTime = DateTime.Now.AddDays(-1),
                    Text = "Consider adding more details here.",
                    ParagraphIndex = 1
                },
                new CommentData
                {
                    Author = "Charlie",
                    Initial = "CH",
                    DateTime = DateTime.Now,
                    Text = "Typo detected.",
                    ParagraphIndex = 2
                }
            };
        }

        // Creates a basic template document with three paragraphs.
        private static void CreateTemplateDocument(string filePath)
        {
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("This is the first paragraph of the template.");
            builder.Writeln("This is the second paragraph of the template.");
            builder.Writeln("This is the third paragraph of the template.");

            template.Save(filePath);
        }

        // Inserts comments into the specified document.
        private static void InsertCommentsIntoDocument(Document doc, List<CommentData> comments)
        {
            // Ensure the document has at least one section and a body.
            if (doc.FirstSection?.Body == null)
                return;

            // Retrieve all paragraphs in the main story.
            List<Paragraph> paragraphs = doc.FirstSection.Body.Paragraphs.Cast<Paragraph>().ToList();

            foreach (CommentData data in comments)
            {
                // Validate the target paragraph index.
                if (data.ParagraphIndex < 0 || data.ParagraphIndex >= paragraphs.Count)
                    continue; // Skip invalid indices.

                Paragraph targetParagraph = paragraphs[data.ParagraphIndex];
                if (targetParagraph == null)
                    continue;

                // Create a new comment node.
                Comment comment = new Comment(doc)
                {
                    Author = data.Author,
                    Initial = data.Initial,
                    DateTime = data.DateTime
                };

                // The comment must contain at least one paragraph and run with visible text.
                Paragraph commentParagraph = new Paragraph(doc);
                Run commentRun = new Run(doc, data.Text);
                commentParagraph.AppendChild(commentRun);
                comment.AppendChild(commentParagraph);

                // Attach the comment to the target paragraph.
                targetParagraph.AppendChild(comment);
            }
        }
    }
}
