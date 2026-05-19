using System;
using System.Collections.Generic;
using Aspose.Words;

namespace CommentInsertionExample
{
    // Simple data model representing a comment record from a database.
    public class CommentRecord
    {
        public string Author { get; set; } = "";
        public string Initial { get; set; } = "";
        public string Text { get; set; } = "";
        public DateTime DateTime { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Simulate retrieving comment data from a database.
            List<CommentRecord> commentData = GetSampleCommentData();

            // Create a new blank document that will serve as the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add an introductory paragraph to the document.
            builder.Writeln("Document generated with comments from simulated database:");

            // Insert a paragraph and a comment for each record.
            foreach (CommentRecord record in commentData)
            {
                // Add a paragraph that the comment will be attached to.
                builder.Writeln($"Paragraph for comment by {record.Author}:");

                // Ensure the builder's current paragraph is not null.
                Paragraph? targetParagraph = builder.CurrentParagraph;
                if (targetParagraph == null)
                {
                    // If for some reason there is no current paragraph, create one.
                    targetParagraph = new Paragraph(doc);
                    doc.FirstSection.Body.AppendChild(targetParagraph);
                }

                // Create a new comment node and set its metadata.
                Comment comment = new Comment(doc)
                {
                    Author = record.Author,
                    Initial = record.Initial,
                    DateTime = record.DateTime
                };

                // Add at least one paragraph and run inside the comment so it has visible text.
                comment.AppendChild(new Paragraph(doc));
                comment.FirstParagraph.AppendChild(new Run(doc, record.Text));

                // Append the comment to the paragraph.
                targetParagraph.AppendChild(comment);
            }

            // Save the resulting document.
            const string outputPath = "CommentsInserted.docx";
            doc.Save(outputPath);

            // Optional: enumerate and display the inserted comments in the console.
            var comments = doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>();
            foreach (Comment c in comments)
            {
                Console.WriteLine($"{c.Author} ({c.Initial}) on {c.DateTime:u}: {c.GetText().Trim()}");
            }
        }

        // Generates sample comment records to mimic database rows.
        private static List<CommentRecord> GetSampleCommentData()
        {
            return new List<CommentRecord>
            {
                new CommentRecord
                {
                    Author = "Alice Johnson",
                    Initial = "AJ",
                    Text = "Please review this section.",
                    DateTime = DateTime.Now.AddDays(-2)
                },
                new CommentRecord
                {
                    Author = "Bob Smith",
                    Initial = "BS",
                    Text = "Consider rephrasing the previous sentence.",
                    DateTime = DateTime.Now.AddDays(-1)
                },
                new CommentRecord
                {
                    Author = "Carol Lee",
                    Initial = "CL",
                    Text = "Add a reference here.",
                    DateTime = DateTime.Now
                }
            };
        }
    }
}
