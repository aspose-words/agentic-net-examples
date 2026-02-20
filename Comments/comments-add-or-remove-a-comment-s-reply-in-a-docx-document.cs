using System;
using Aspose.Words;

class CommentReplyExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Find the first top‑level comment in the document.
        Comment topComment = doc.GetChildNodes(NodeType.Comment, true)[0] as Comment;
        if (topComment == null)
        {
            Console.WriteLine("No comments found in the document.");
            return;
        }

        // -------------------- Add a reply --------------------
        // Create a new comment that will act as a reply.
        Comment reply = new Comment(doc, "Jane Doe", "J.D.", DateTime.Now);
        // Set the reply text.
        reply.SetText("This is a reply to the original comment.");
        // Link the reply to its parent comment.
        reply.ParentId = topComment.Id;
        // Add the reply to the parent comment's Replies collection.
        topComment.Replies.Add(reply);

        // Save the document with the added reply.
        doc.Save("Output_WithReply.docx");

        // -------------------- Remove the reply --------------------
        // Remove the reply comment from the parent comment's Replies collection.
        topComment.Replies.Remove(reply);

        // Save the document after removing the reply.
        doc.Save("Output_ReplyRemoved.docx");
    }
}
