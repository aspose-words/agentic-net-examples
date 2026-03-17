using System;
using System.Linq;
using Aspose.Words;

class CommentProcessor
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Get all comments in the document. Comment class lives directly in Aspose.Words namespace.
        Comment[] allComments = doc.GetChildNodes(NodeType.Comment, true)
                                   .Cast<Comment>()
                                   .ToArray();

        foreach (Comment comment in allComments)
        {
            // Keep comments authored by "John Doe"; remove all others.
            if (!string.Equals(comment.Author, "John Doe", StringComparison.OrdinalIgnoreCase))
            {
                comment.Remove();
            }
            // Matching author – comment is kept (accepted).
        }

        // Accept any tracked changes that may exist in the document.
        doc.AcceptAllRevisions();

        // Save the revised document.
        doc.Save("Revised.docx");
    }
}
