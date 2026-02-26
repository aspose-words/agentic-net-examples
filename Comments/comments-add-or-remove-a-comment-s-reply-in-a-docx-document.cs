using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCommentDemo
{
    class Program
    {
        static void Main()
        {
            // Path where the demo files will be saved.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -------------------------------------------------
            // 1. Create a new document and add a comment with a reply.
            // -------------------------------------------------
            Document doc = new Document();                     // Create a blank document.
            DocumentBuilder builder = new DocumentBuilder(doc); // Helper to add content.

            // Add some text to comment.
            builder.Writeln("This is a paragraph that will have a comment.");

            // Create a top‑level comment.
            Comment topComment = new Comment(doc, "Alice", "A", DateTime.Now);
            topComment.SetText("Initial comment text.");
            // Append the comment to the current paragraph.
            builder.CurrentParagraph.AppendChild(topComment);

            // Add a reply to the comment.
            topComment.AddReply("Bob", "B", DateTime.Now, "This is a reply to the comment.");

            // Save the document that now contains a comment with a reply.
            string addReplyPath = Path.Combine(outputDir, "AddReply.docx");
            doc.Save(addReplyPath); // Save using the overload that determines format from extension.

            // -------------------------------------------------
            // 2. Load the document, remove the reply, and save again.
            // -------------------------------------------------
            Document loadedDoc = new Document(addReplyPath); // Load the previously saved document.

            // Retrieve the first comment in the document.
            Comment comment = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                       .Cast<Comment>()
                                       .FirstOrDefault();

            if (comment != null && comment.Replies.Count > 0)
            {
                // Remove the first reply.
                comment.RemoveReply(comment.Replies[0]);
                // Alternatively, to remove all replies at once:
                // comment.RemoveAllReplies();
            }

            // Save the document after the reply has been removed.
            string removeReplyPath = Path.Combine(outputDir, "RemoveReply.docx");
            loadedDoc.Save(removeReplyPath);
        }
    }
}
