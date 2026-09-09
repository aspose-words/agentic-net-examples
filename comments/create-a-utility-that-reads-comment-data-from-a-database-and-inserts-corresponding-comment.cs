using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;

namespace CommentInsertionExample
{
    // Simple POCO representing a comment record that might come from a database.
    public class CommentData
    {
        public string Author { get; set; } = "";
        public string Initial { get; set; } = "";
        public DateTime DateTime { get; set; }
        public string Text { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Simulate retrieving comment data from a database.
            List<CommentData> commentRecords = GetSampleCommentData();

            // Create a template document with a few paragraphs.
            Document template = CreateTemplateDocument();

            // Insert comments from the simulated database into the template.
            InsertCommentsIntoDocument(template, commentRecords);

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DocumentWithComments.docx");
            template.Save(outputPath);

            // Load the saved document and enumerate the comments to verify insertion.
            Document loadedDoc = new Document(outputPath);
            EnumerateComments(loadedDoc);
        }

        // Returns a list of sample comment data.
        private static List<CommentData> GetSampleCommentData()
        {
            return new List<CommentData>
            {
                new CommentData
                {
                    Author = "Alice Johnson",
                    Initial = "AJ",
                    DateTime = DateTime.Now.AddDays(-2),
                    Text = "Review the introduction."
                },
                new CommentData
                {
                    Author = "Bob Smith",
                    Initial = "BS",
                    DateTime = DateTime.Now.AddDays(-1),
                    Text = "Consider adding more examples here."
                },
                new CommentData
                {
                    Author = "Carol Lee",
                    Initial = "CL",
                    DateTime = DateTime.Now,
                    Text = "Check the formatting of this section."
                }
            };
        }

        // Creates a simple template document with three paragraphs.
        private static Document CreateTemplateDocument()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Paragraph 1: This is the first paragraph of the template.");
            builder.Writeln("Paragraph 2: This is the second paragraph of the template.");
            builder.Writeln("Paragraph 3: This is the third paragraph of the template.");

            return doc;
        }

        // Inserts each comment into a corresponding paragraph of the document.
        private static void InsertCommentsIntoDocument(Document doc, List<CommentData> comments)
        {
            // Ensure the document has at least as many paragraphs as comments.
            int paragraphCount = doc.FirstSection?.Body?.Paragraphs?.Count ?? 0;
            int requiredCount = comments.Count;
            if (paragraphCount < requiredCount)
            {
                DocumentBuilder extraBuilder = new DocumentBuilder(doc);
                for (int i = paragraphCount; i < requiredCount; i++)
                {
                    extraBuilder.Writeln($"Additional paragraph {i + 1}.");
                }
            }

            // Attach each comment to the paragraph with the same index.
            for (int i = 0; i < comments.Count; i++)
            {
                CommentData data = comments[i];
                Paragraph? paragraph = doc.FirstSection?.Body?.Paragraphs[i];
                if (paragraph == null)
                    continue; // Safety check; should not happen.

                // Create a new comment node.
                Comment comment = new Comment(doc, data.Author, data.Initial, data.DateTime);
                comment.SetText(data.Text);

                // Append the comment to the paragraph.
                paragraph.AppendChild(comment);
            }
        }

        // Enumerates all comments in the document and writes their details to the console.
        private static void EnumerateComments(Document doc)
        {
            var commentNodes = doc.GetChildNodes(NodeType.Comment, true)
                                  .OfType<Comment>()
                                  .ToList();

            foreach (Comment c in commentNodes)
            {
                string author = c.Author ?? "Unknown";
                string text = c.GetText().Trim();
                DateTime date = c.DateTime;
                Console.WriteLine($"Comment by {author} on {date:G}: \"{text}\"");
            }
        }
    }
}
