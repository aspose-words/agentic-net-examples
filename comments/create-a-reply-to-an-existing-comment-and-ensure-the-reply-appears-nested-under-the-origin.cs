using System;
using System.IO;
using System.Linq;
using Aspose.Words;

namespace CommentReplyExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for convenient document editing.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a paragraph that will host the comment.
            builder.Writeln("This paragraph will have a comment with a reply.");

            // Ensure the paragraph was created.
            Paragraph? paragraph = builder.CurrentParagraph;
            if (paragraph == null)
                throw new InvalidOperationException("Failed to create a paragraph.");

            // Create a top‑level comment.
            Comment topComment = new Comment(doc, "Alice", "A", DateTime.Now);
            topComment.SetText("Original comment text.");

            // Append the comment to the paragraph. It will appear in the margin of the document.
            paragraph.AppendChild(topComment);

            // Add a reply to the top‑level comment. The reply will be nested under the original comment.
            topComment.AddReply("Bob", "B", DateTime.Now, "This is a reply to Alice's comment.");

            // Prepare an output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "CommentReplyExample.docx");
            doc.Save(outputPath);

            // Optional: enumerate comments and write their hierarchy to the console.
            var allComments = doc.GetChildNodes(NodeType.Comment, true)
                                 .OfType<Comment>()
                                 .ToList();

            foreach (Comment comment in allComments.Where(c => c.Ancestor == null))
            {
                Console.WriteLine($"Top‑level comment by {comment.Author}: {comment.GetText().Trim()}");
                foreach (Comment reply in comment.Replies)
                {
                    Console.WriteLine($"  Reply by {reply.Author}: {reply.GetText().Trim()}");
                }
            }

            // Indicate completion.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
