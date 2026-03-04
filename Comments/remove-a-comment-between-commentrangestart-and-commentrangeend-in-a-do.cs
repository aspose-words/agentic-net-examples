using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;

class RemoveCommentExample
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Collect all top‑level comments.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
        List<Comment> comments = new List<Comment>();
        foreach (Comment c in commentNodes)
        {
            // Only top‑level comments have no ancestor comment.
            if (c.Ancestor == null)
                comments.Add(c);
        }

        // For each comment remove its range start, range end and the comment itself.
        foreach (Comment comment in comments)
        {
            int commentId = comment.Id;

            // Find and remove the matching CommentRangeStart node.
            NodeCollection rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true);
            foreach (CommentRangeStart start in rangeStarts)
            {
                if (start.Id == commentId)
                {
                    start.Remove();
                    break; // IDs are unique, safe to exit loop.
                }
            }

            // Find and remove the matching CommentRangeEnd node.
            NodeCollection rangeEnds = doc.GetChildNodes(NodeType.CommentRangeEnd, true);
            foreach (CommentRangeEnd end in rangeEnds)
            {
                if (end.Id == commentId)
                {
                    end.Remove();
                    break;
                }
            }

            // Finally remove the comment node itself.
            comment.Remove();
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
