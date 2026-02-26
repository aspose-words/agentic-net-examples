using System;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Linq;

class Program
{
    static void Main()
    {
        // Load the DOCX file that contains comments.
        Document doc = new Document("Comments.docx");

        // Retrieve all comment nodes in the document (including replies).
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate through top‑level comments only (those without an ancestor comment).
        foreach (Comment comment in commentNodes.OfType<Comment>().Where(c => c.Ancestor == null))
        {
            // Print the main comment text and author.
            Console.WriteLine($"Comment: \"{comment.GetText().Trim()}\" by {comment.Author}");

            // Iterate over each reply to the current comment.
            foreach (Comment reply in comment.Replies)
            {
                // Print the reply text and author.
                Console.WriteLine($"\tReply: \"{reply.GetText().Trim()}\" by {reply.Author}");
            }
        }
    }
}
