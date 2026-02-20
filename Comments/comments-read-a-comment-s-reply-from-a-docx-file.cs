using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file.
        Document doc = new Document("CommentsSample.docx");

        // Get all top‑level comments in the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

        foreach (Node node in commentNodes)
        {
            Comment comment = (Comment)node;
            // Print the main comment text.
            Console.WriteLine($"Comment by {comment.Author}: {comment.GetText().Trim()}");

            // Iterate through any replies to this comment.
            foreach (Comment reply in comment.Replies)
            {
                // Print the reply text.
                Console.WriteLine($"\tReply by {reply.Author}: {reply.GetText().Trim()}");
            }
        }
    }
}
