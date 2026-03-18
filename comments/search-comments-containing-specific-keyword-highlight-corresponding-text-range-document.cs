using System;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

class HighlightCommentedText
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = "input.docx";
        // Path where the modified document will be saved.
        string outputPath = "output.docx";
        // Keyword to search for inside comment texts.
        string keyword = "TODO";

        // Load the document.
        Document doc = new Document(inputPath);

        // Iterate through all top‑level comments in the document.
        foreach (Comment comment in doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>())
        {
            // Check if the comment text contains the desired keyword.
            if (comment.GetText().Contains(keyword))
            {
                int commentId = comment.Id;

                // Locate the start and end nodes that delimit the commented range.
                CommentRangeStart rangeStart = doc.GetChildNodes(NodeType.CommentRangeStart, true)
                    .Cast<CommentRangeStart>()
                    .FirstOrDefault(s => s.Id == commentId);

                CommentRangeEnd rangeEnd = doc.GetChildNodes(NodeType.CommentRangeEnd, true)
                    .Cast<CommentRangeEnd>()
                    .FirstOrDefault(e => e.Id == commentId);

                if (rangeStart == null || rangeEnd == null)
                    continue; // Safety check.

                // Walk through the nodes that lie between the start and end markers.
                Node curNode = rangeStart.NextSibling;
                while (curNode != null && curNode != rangeEnd)
                {
                    // Apply highlight to each Run node within the range.
                    if (curNode.NodeType == NodeType.Run)
                    {
                        ((Run)curNode).Font.HighlightColor = Color.Yellow;
                    }
                    curNode = curNode.NextSibling;
                }
            }
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
