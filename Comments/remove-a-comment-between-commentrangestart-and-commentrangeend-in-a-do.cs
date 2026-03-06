using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

class RemoveCommentExample
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Get all top‑level comments in the document.
        NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

        // Iterate over a copy of the collection because we will modify the document.
        foreach (Comment comment in comments.Cast<Comment>().ToList())
        {
            // Remove the comment node itself.
            comment.Remove();

            // Find the matching CommentRangeStart node (same Id as the comment).
            CommentRangeStart rangeStart = (CommentRangeStart)doc.GetChild(
                NodeType.CommentRangeStart, comment.Id, true);

            // Find the matching CommentRangeEnd node (same Id as the comment).
            CommentRangeEnd rangeEnd = (CommentRangeEnd)doc.GetChild(
                NodeType.CommentRangeEnd, comment.Id, true);

            // Remove the range start and end nodes if they exist.
            rangeStart?.Remove();
            rangeEnd?.Remove();
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
