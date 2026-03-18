using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

class CommentPositionSynchronizer
{
    static void Main()
    {
        // Load the source document.
        string inputPath = @"Input.docx";
        Document doc = new Document(inputPath);

        // -------------------------------------------------
        // Reorder sections: move the last section to the front.
        // -------------------------------------------------
        if (doc.Sections.Count > 1)
        {
            // Remove the last section from its current position.
            Section lastSection = doc.Sections[doc.Sections.Count - 1];
            lastSection.Remove();

            // Insert it at the beginning of the document.
            doc.Sections.Insert(0, lastSection);
        }

        // -------------------------------------------------
        // Synchronize comment anchors after the reorder.
        // -------------------------------------------------
        // Gather all comment range start/end nodes once for efficiency.
        var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true)
                             .Cast<CommentRangeStart>()
                             .ToDictionary(r => r.Id);
        var rangeEnds = doc.GetChildNodes(NodeType.CommentRangeEnd, true)
                           .Cast<CommentRangeEnd>()
                           .ToDictionary(r => r.Id);

        // Iterate over every comment in the document.
        foreach (Comment comment in doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>())
        {
            // Try to locate the matching start and end nodes by Id.
            if (!rangeStarts.TryGetValue(comment.Id, out CommentRangeStart start) ||
                !rangeEnds.TryGetValue(comment.Id, out CommentRangeEnd end))
                continue; // No matching range – nothing to synchronize.

            // Ensure the range end is placed immediately before the comment node.
            if (comment.PreviousSibling != end)
            {
                // Detach the end node from its current location.
                end.Remove();

                // Insert it right before the comment.
                comment.ParentNode.InsertBefore(end, comment);
            }

            // Ensure the range start is placed immediately before the range end.
            if (start.NextSibling != end)
            {
                // Detach the start node from its current location.
                start.Remove();

                // Insert it right before the range end.
                end.ParentNode.InsertBefore(start, end);
            }
        }

        // -------------------------------------------------
        // Save the modified document.
        // -------------------------------------------------
        string outputPath = @"Output.docx";
        doc.Save(outputPath);
    }
}
