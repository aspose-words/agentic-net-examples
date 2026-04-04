using System;
using System.Drawing;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment that contains the keyword "TODO".
        builder.Writeln("This is a sample paragraph with a TODO comment.");

        // Retrieve the paragraph we just added.
        Paragraph firstPara = doc.FirstSection.Body.FirstParagraph;

        // Create the comment.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("TODO: Review this sentence.");

        // Anchor the comment to the word "TODO".
        int commentId1 = comment1.Id;
        firstPara.AppendChild(new CommentRangeStart(doc, commentId1));
        firstPara.AppendChild(new Run(doc, "TODO"));
        firstPara.AppendChild(new CommentRangeEnd(doc, commentId1));
        firstPara.AppendChild(comment1);

        // Second paragraph with a comment that does NOT contain the keyword.
        builder.Writeln("Another line without the keyword.");

        // Retrieve the second paragraph.
        Paragraph secondPara = doc.FirstSection.Body.LastParagraph;

        // Create the second comment.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Note: No action needed.");

        // Anchor the second comment to the word "Note".
        int commentId2 = comment2.Id;
        secondPara.AppendChild(new CommentRangeStart(doc, commentId2));
        secondPara.AppendChild(new Run(doc, "Note"));
        secondPara.AppendChild(new CommentRangeEnd(doc, commentId2));
        secondPara.AppendChild(comment2);

        // Define the keyword to search for inside comment texts.
        const string keyword = "TODO";

        // Enumerate all comments in the document.
        var allComments = doc.GetChildNodes(NodeType.Comment, true)
                             .OfType<Comment>()
                             .ToList();

        foreach (Comment comment in allComments)
        {
            // Check if the comment text contains the keyword (case‑insensitive).
            if (comment.GetText().IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Find the corresponding CommentRangeStart node by matching Id.
                CommentRangeStart? rangeStart = doc.GetChildNodes(NodeType.CommentRangeStart, true)
                                                  .OfType<CommentRangeStart>()
                                                  .FirstOrDefault(crs => crs.Id == comment.Id);

                if (rangeStart == null)
                    continue; // Safety check – should not happen.

                // Walk through the nodes between the start and end of the comment range.
                Node curNode = rangeStart.NextSibling;
                while (curNode != null)
                {
                    // Stop when we reach the matching CommentRangeEnd.
                    if (curNode.NodeType == NodeType.CommentRangeEnd &&
                        ((CommentRangeEnd)curNode).Id == comment.Id)
                    {
                        break;
                    }

                    // Highlight any Run nodes inside the range.
                    if (curNode.NodeType == NodeType.Run)
                    {
                        ((Run)curNode).Font.HighlightColor = Color.Yellow;
                    }

                    curNode = curNode.NextSibling;
                }
            }
        }

        // Save the resulting document.
        const string outputPath = "HighlightedComments.docx";
        doc.Save(outputPath);
    }
}
