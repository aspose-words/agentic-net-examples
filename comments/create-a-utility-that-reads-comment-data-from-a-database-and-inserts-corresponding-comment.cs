using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace CommentInsertionExample
{
    // Simple data model representing a comment record from a database.
    public class CommentData
    {
        public string Author { get; set; } = "";
        public string Initial { get; set; } = "";
        public DateTime DateTime { get; set; }
        public string Text { get; set; } = "";
        // Zero‑based index of the paragraph to which the comment will be attached.
        public int ParagraphIndex { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Simulated database records.
            List<CommentData> commentRecords = new List<CommentData>
            {
                new CommentData
                {
                    Author = "Alice",
                    Initial = "A",
                    DateTime = DateTime.Now,
                    Text = "Review this opening paragraph.",
                    ParagraphIndex = 0
                },
                new CommentData
                {
                    Author = "Bob",
                    Initial = "B",
                    DateTime = DateTime.Now.AddMinutes(-5),
                    Text = "Consider rephrasing this sentence.",
                    ParagraphIndex = 2
                },
                new CommentData
                {
                    Author = "Carol",
                    Initial = "C",
                    DateTime = DateTime.Now.AddHours(-1),
                    Text = "Add a reference here.",
                    ParagraphIndex = 4
                }
            };

            // Create a blank document and add a few paragraphs that will serve as the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 0; i < 5; i++)
            {
                builder.Writeln($"Paragraph {i + 1}: This is sample text for paragraph {i + 1}.");
            }

            // Insert comments based on the simulated data.
            foreach (CommentData record in commentRecords)
            {
                // Safely obtain the target paragraph.
                Paragraph? targetParagraph = doc.FirstSection?.Body?.Paragraphs.ElementAtOrDefault(record.ParagraphIndex) as Paragraph;
                if (targetParagraph == null)
                    continue; // Skip if the index is out of range.

                // Move the builder to the target paragraph.
                builder.MoveTo(targetParagraph);

                // Create a new comment node with metadata.
                Comment comment = new Comment(doc, record.Author, record.Initial, record.DateTime);
                // Set the comment text (adds a paragraph internally).
                comment.SetText(record.Text);
                // Append the comment to the current paragraph.
                builder.CurrentParagraph.AppendChild(comment);
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentsInserted.docx");
            doc.Save(outputPath);

            // Enumerate and display the inserted comments to verify.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment c in comments)
            {
                Console.WriteLine($"{c.Author} ({c.Initial}) on {c.DateTime}: {c.GetText().Trim()}");
            }

            // Indicate completion.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
