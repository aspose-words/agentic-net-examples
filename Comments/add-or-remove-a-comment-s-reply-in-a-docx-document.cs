using System;
using Aspose.Words;

namespace CommentReplyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some text to the document.
            builder.Writeln("This is a sample paragraph.");

            // Create a top‑level comment and attach it to the current paragraph.
            Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
            comment.SetText("Initial comment.");
            builder.CurrentParagraph.AppendChild(comment);

            // Add a reply to the comment.
            Comment reply = comment.AddReply("Bob", "B", DateTime.Now, "First reply.");

            // Add a second reply for demonstration.
            comment.AddReply("Charlie", "C", DateTime.Now, "Second reply");

            // At this point the comment has two replies.
            Console.WriteLine($"Replies before removal: {comment.Replies.Count}");

            // Remove the first reply using RemoveReply.
            comment.RemoveReply(reply);

            // Verify that only one reply remains.
            Console.WriteLine($"Replies after RemoveReply: {comment.Replies.Count}");

            // Remove all remaining replies using RemoveAllReplies.
            comment.RemoveAllReplies();

            // Verify that no replies remain.
            Console.WriteLine($"Replies after RemoveAllReplies: {comment.Replies.Count}");

            // Save the document with the modifications.
            doc.Save("CommentReplyDemo.docx");
        }
    }
}
