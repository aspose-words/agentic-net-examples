using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace CommentRemovalExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add first paragraph and an old comment (10 days ago).
            builder.Writeln("First paragraph with an old comment.");
            Comment oldComment = new Comment(doc, "Alice", "A", DateTime.Now.AddDays(-10));
            oldComment.SetText("This comment is older than the threshold.");
            // Append the comment to the current paragraph.
            builder.CurrentParagraph.AppendChild(oldComment);

            // Add second paragraph and a recent comment (today).
            builder.Writeln("Second paragraph with a recent comment.");
            Comment recentComment = new Comment(doc, "Bob", "B", DateTime.Now);
            recentComment.SetText("This comment is within the threshold.");
            builder.CurrentParagraph.AppendChild(recentComment);

            // Define the date threshold: comments older than this date will be removed.
            DateTime threshold = DateTime.Now.AddDays(-5);

            // Enumerate all comment nodes safely using ToList to avoid modifying the collection during iteration.
            var allComments = doc.GetChildNodes(NodeType.Comment, true)
                                 .OfType<Comment>()
                                 .ToList();

            foreach (Comment comment in allComments)
            {
                // Remove comment if its DateTime is earlier than the threshold.
                if (comment.DateTime < threshold)
                {
                    comment.Remove();
                }
            }

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentsFiltered.docx");
            doc.Save(outputPath);
        }
    }
}
