using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace CommentRemovalExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add first paragraph with a recent comment.
            builder.Writeln("First paragraph with a recent comment.");
            Comment recentComment = new Comment(doc)
            {
                Author = "Alice",
                Initial = "A",
                DateTime = DateTime.Now.AddDays(-2) // Recent comment.
            };
            // Add visible text to the comment.
            recentComment.AppendChild(new Paragraph(doc));
            recentComment.FirstParagraph?.AppendChild(new Run(doc, "This is a recent comment."));
            // Attach the comment to the first paragraph.
            doc.FirstSection?.Body?.FirstParagraph?.AppendChild(recentComment);

            // Add second paragraph with an old comment.
            builder.Writeln("Second paragraph with an old comment.");
            Comment oldComment = new Comment(doc)
            {
                Author = "Bob",
                Initial = "B",
                DateTime = DateTime.Now.AddDays(-10) // Old comment.
            };
            oldComment.AppendChild(new Paragraph(doc));
            oldComment.FirstParagraph?.AppendChild(new Run(doc, "This is an old comment."));
            // Attach the comment to the second paragraph.
            Paragraph? secondParagraph = doc.FirstSection?.Body?.LastParagraph;
            secondParagraph?.AppendChild(oldComment);

            // Define the date threshold: comments older than this will be removed.
            DateTime threshold = DateTime.Now.AddDays(-5);

            // Enumerate all comments safely.
            var allComments = doc.GetChildNodes(NodeType.Comment, true)
                                 .OfType<Comment>()
                                 .ToList();

            // Remove comments older than the threshold.
            foreach (Comment comment in allComments)
            {
                if (comment.DateTime < threshold)
                {
                    comment.Remove();
                }
            }

            // Ensure output directory exists.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
            doc.Save(outputPath);
        }
    }
}
