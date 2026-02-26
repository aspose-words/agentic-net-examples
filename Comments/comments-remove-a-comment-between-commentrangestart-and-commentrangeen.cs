using System;
using Aspose.Words;

class RemoveCommentExample
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Get all top‑level comments in the document.
        NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate over a copy of the collection because we will modify the document while iterating.
        foreach (Comment comment in comments.ToArray())
        {
            // The comment range start node is three siblings before the comment node:
            // CommentRangeStart -> Run (commented text) -> CommentRangeEnd -> Comment
            Node commentRangeEnd = comment.PreviousSibling;
            Node commentRangeStart = commentRangeEnd?.PreviousSibling?.PreviousSibling;

            // Remove the comment range start node if it exists.
            commentRangeStart?.Remove();

            // Remove the comment range end node if it exists.
            commentRangeEnd?.Remove();

            // Finally remove the comment itself.
            comment.Remove();
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
