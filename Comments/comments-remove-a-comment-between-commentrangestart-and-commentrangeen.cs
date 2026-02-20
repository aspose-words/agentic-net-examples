using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

class RemoveCommentExample
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Collect all CommentRangeStart nodes first (so we can modify the document while iterating).
        NodeCollection rangeStartNodes = doc.GetChildNodes(NodeType.CommentRangeStart, true);
        List<CommentRangeStart> rangeStarts = rangeStartNodes
            .Cast<CommentRangeStart>()
            .ToList();

        foreach (CommentRangeStart rangeStart in rangeStarts)
        {
            int commentId = rangeStart.Id;

            // Find the Comment node with the same Id.
            Comment commentNode = doc.GetChildNodes(NodeType.Comment, true)
                .Cast<Comment>()
                .FirstOrDefault(c => c.Id == commentId);

            // Remove the comment node if it exists.
            commentNode?.Remove();

            // Remove the CommentRangeStart node.
            rangeStart.Remove();

            // Find and remove the matching CommentRangeEnd node.
            CommentRangeEnd rangeEnd = doc.GetChildNodes(NodeType.CommentRangeEnd, true)
                .Cast<CommentRangeEnd>()
                .FirstOrDefault(e => e.Id == commentId);
            rangeEnd?.Remove();
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
