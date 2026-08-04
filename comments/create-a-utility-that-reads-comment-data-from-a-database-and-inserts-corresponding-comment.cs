using System;
using System.Collections.Generic;
using Aspose.Words;

namespace CommentInsertionUtility
{
    // Simple data model representing a comment record that could come from a database.
    public class CommentRecord
    {
        public string Author { get; set; } = "";
        public string Initial { get; set; } = "";
        public DateTime DateTime { get; set; }
        public string Text { get; set; } = "";
        // Zero‑based index of the paragraph in the template where the comment will be attached.
        public int ParagraphIndex { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Simulate reading comment data from a database.
            List<CommentRecord> commentData = GetSampleCommentData();

            // Create a simple template document with a few paragraphs.
            Document template = CreateTemplateDocument();

            // Insert comments into the template based on the simulated data.
            InsertCommentsIntoDocument(template, commentData);

            // Save the resulting document.
            const string outputPath = "TemplateWithComments.docx";
            template.Save(outputPath);
        }

        // Returns a hard‑coded list of comment records.
        private static List<CommentRecord> GetSampleCommentData()
        {
            return new List<CommentRecord>
            {
                new CommentRecord
                {
                    Author = "Alice Johnson",
                    Initial = "AJ",
                    DateTime = DateTime.Now.AddDays(-2),
                    Text = "Please verify the figures in this paragraph.",
                    ParagraphIndex = 0
                },
                new CommentRecord
                {
                    Author = "Bob Smith",
                    Initial = "BS",
                    DateTime = DateTime.Now.AddDays(-1),
                    Text = "Consider rephrasing this sentence for clarity.",
                    ParagraphIndex = 1
                },
                new CommentRecord
                {
                    Author = "Carol Lee",
                    Initial = "CL",
                    DateTime = DateTime.Now,
                    Text = "Add a reference to the source material.",
                    ParagraphIndex = 2
                }
            };
        }

        // Creates a basic document that will serve as the template.
        private static Document CreateTemplateDocument()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Paragraph 1: Introduction to the report.");
            builder.Writeln("Paragraph 2: Detailed analysis of the data.");
            builder.Writeln("Paragraph 3: Conclusions and recommendations.");

            return doc;
        }

        // Inserts comments into the specified document according to the provided records.
        private static void InsertCommentsIntoDocument(Document doc, List<CommentRecord> records)
        {
            // Ensure the document has at least one section and a body.
            if (doc.FirstSection?.Body == null)
                return;

            // Iterate over each comment record.
            foreach (CommentRecord record in records)
            {
                // Validate the paragraph index.
                if (record.ParagraphIndex < 0 ||
                    record.ParagraphIndex >= doc.FirstSection.Body.Paragraphs.Count)
                {
                    // Skip invalid indices.
                    continue;
                }

                // Retrieve the target paragraph.
                Paragraph? targetParagraph = doc.FirstSection.Body.Paragraphs[record.ParagraphIndex];
                if (targetParagraph == null)
                    continue;

                // Create a new comment node with the required metadata.
                Comment comment = new Comment(doc, record.Author, record.Initial, record.DateTime);
                comment.SetText(record.Text);

                // Append the comment to the paragraph.
                targetParagraph.AppendChild(comment);
            }
        }
    }
}
