using System;
using Aspose.Words;

namespace CommentReplyDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX document.
            // Replace with the actual path to your input file.
            string inputPath = "Input.docx";
            Document doc = new Document(inputPath);

            // Create a DocumentBuilder to work with the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Ensure there is a paragraph to attach the comment to.
            builder.Writeln("This is a sample paragraph.");

            // Create a top‑level comment.
            Comment comment = new Comment(doc, "John Doe", "J.D.", DateTime.Now);
            comment.SetText("Original comment text.");
            // Append the comment to the current paragraph.
            builder.CurrentParagraph.AppendChild(comment);

            // Add a reply to the comment.
            Comment reply = comment.AddReply(
                author: "Jane Smith",
                initial: "J.S.",
                dateTime: DateTime.Now,
                text: "This is a reply to the original comment.");

            // At this point the document contains the comment and one reply.
            // -----------------------------------------------------------------
            // Example 1: Remove a specific reply.
            comment.RemoveReply(reply);

            // Example 2: Add two replies and then remove all of them at once.
            comment.AddReply("Alice Brown", "A.B.", DateTime.Now, "First additional reply.");
            comment.AddReply("Bob White", "B.W.", DateTime.Now, "Second additional reply.");
            // Remove all replies from the comment.
            comment.RemoveAllReplies();

            // Save the modified document.
            // Replace with the desired output path.
            string outputPath = "Output.docx";
            doc.Save(outputPath);
        }
    }
}
