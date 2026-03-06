using System;
using System.Linq;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file that contains comments and replies.
        Document doc = new Document("Comments.docx");

        // Retrieve all comment nodes in the document (including replies).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate only top‑level comments (those without an ancestor comment).
        foreach (Comment comment in commentNodes.OfType<Comment>().Where(c => c.Ancestor == null))
        {
            // Print the text and author of the top‑level comment.
            Console.WriteLine($"Comment: \"{comment.GetText().Trim()}\", Author: {comment.Author}");

            // Iterate through any replies to this comment.
            foreach (Comment reply in comment.Replies)
            {
                // Print the text and author of each reply.
                Console.WriteLine($"\tReply: \"{reply.GetText().Trim()}\", Author: {reply.Author}");
            }
        }
    }
}
