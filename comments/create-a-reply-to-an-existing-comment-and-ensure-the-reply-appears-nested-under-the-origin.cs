using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold the comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a top‑level comment.
        Comment topComment = new Comment(doc, "Alice", "A", DateTime.Now);
        topComment.SetText("Please review this paragraph.");

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(topComment);

        // Add a reply to the top‑level comment.
        topComment.AddReply("Bob", "B", DateTime.Now, "I have reviewed it, looks good.");

        // Enumerate comments to demonstrate the nesting.
        var comments = doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>();
        foreach (Comment comment in comments.Where(c => c.Ancestor == null))
        {
            Console.WriteLine($"Comment by {comment.Author}: {comment.GetText().Trim()}");
            foreach (Comment reply in comment.Replies)
            {
                Console.WriteLine($"\tReply by {reply.Author}: {reply.GetText().Trim()}");
            }
        }

        // Save the document with the comment and its reply.
        doc.Save("CommentReplyExample.docx");
    }
}
