using System;
using Aspose.Words;

class ReadCommentReplies
{
    static void Main()
    {
        // Load the DOCX file that contains comments and replies.
        // The Document constructor with a file path loads the document.
        Document doc = new Document("Comments.docx");

        // Get all comment nodes in the document (both top‑level comments and replies).
        NodeCollection allComments = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate through the collection and process only top‑level comments (those without an ancestor).
        foreach (Comment topLevelComment in allComments.OfType<Comment>())
        {
            if (topLevelComment.Ancestor != null)
                continue; // Skip reply comments; we will handle them via the Replies collection.

            Console.WriteLine($"Top‑level comment by {topLevelComment.Author}:");
            Console.WriteLine($"    \"{topLevelComment.GetText().Trim()}\"");

            // The Replies property returns a CommentCollection of immediate child replies.
            foreach (Comment reply in topLevelComment.Replies)
            {
                // GetText() returns the full text of the reply comment.
                string replyText = reply.GetText().Trim();
                Console.WriteLine($"    Reply by {reply.Author}: \"{replyText}\"");
            }

            Console.WriteLine(); // Blank line between comment groups.
        }
    }
}
