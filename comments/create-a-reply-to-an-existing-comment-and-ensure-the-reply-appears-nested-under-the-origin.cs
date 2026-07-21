using System;
using System.Linq;
using Aspose.Words;

public class CommentReplyExample
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
        topComment.AddReply("Bob", "B", DateTime.Now, "Reviewed, looks good.");

        // Enumerate all comments (including replies) to demonstrate nesting.
        var allComments = doc.GetChildNodes(NodeType.Comment, true)
                             .OfType<Comment>()
                             .ToList();

        Console.WriteLine($"Total comment nodes (including replies): {allComments.Count}");
        foreach (Comment c in allComments)
        {
            string level = c.Ancestor == null ? "Top‑level" : "Reply";
            Console.WriteLine($"{level} comment by {c.Author}: {c.GetText().Trim()}");
        }

        // Save the document to the working directory.
        doc.Save("CommentReplyExample.docx");
    }
}
