using System;
using Aspose.Words;

class CommentReplyExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some text to the first paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph that will have a comment.");

        // Create a top‑level comment anchored to the current paragraph.
        Comment topComment = new Comment(doc, "Alice Smith", "A.S.", DateTime.Now);
        topComment.SetText("Please review this sentence.");
        builder.CurrentParagraph.AppendChild(topComment);

        // Add a reply to the top‑level comment. The reply will appear nested under the original comment.
        Comment reply = topComment.AddReply(
            author: "Bob Johnson",
            initial: "B.J.",
            dateTime: DateTime.Now,
            text: "I have checked it, looks good.");

        // Optional: verify the hierarchy (reply.Ancestor should be the top‑level comment).
        if (reply.Ancestor != null && reply.Ancestor == topComment)
        {
            Console.WriteLine("Reply successfully nested under the original comment.");
        }

        // Save the document containing the comment and its reply.
        doc.Save("CommentWithReply.docx");
    }
}
