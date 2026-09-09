using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        // Use DocumentBuilder for convenience when adding content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Create sample paragraphs with comments.
        // -----------------------------------------------------------------
        // First comment – contains the keyword "TODO".
        Paragraph para1 = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(para1);
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("TODO: verify this text.");
        // Build the comment range: [CommentRangeStart] Run [CommentRangeEnd] Comment
        para1.AppendChild(new CommentRangeStart(doc, comment1.Id));
        para1.AppendChild(new Run(doc, "Important text that needs review."));
        para1.AppendChild(new CommentRangeEnd(doc, comment1.Id));
        para1.AppendChild(comment1);

        // Second comment – does NOT contain the keyword.
        Paragraph para2 = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(para2);
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("General note.");
        para2.AppendChild(new CommentRangeStart(doc, comment2.Id));
        para2.AppendChild(new Run(doc, "Additional information."));
        para2.AppendChild(new CommentRangeEnd(doc, comment2.Id));
        para2.AppendChild(comment2);

        // -----------------------------------------------------------------
        // Search for comments containing a specific keyword and highlight
        // the text range they are attached to.
        // -----------------------------------------------------------------
        const string keyword = "TODO";

        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            // Check if the comment text contains the keyword (case‑insensitive).
            if (c.GetText().IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Locate the start and end markers of the comment range.
                CommentRangeStart? rangeStart = c.PreviousSibling?.PreviousSibling?.PreviousSibling as CommentRangeStart;
                CommentRangeEnd?   rangeEnd   = c.PreviousSibling as CommentRangeEnd;

                if (rangeStart != null && rangeEnd != null)
                {
                    // Iterate over all nodes between the start and end markers.
                    Node? curNode = rangeStart.NextSibling;
                    while (curNode != null && curNode != rangeEnd)
                    {
                        if (curNode is Run run)
                        {
                            // Apply yellow highlight to the run's font.
                            run.Font.HighlightColor = Color.Yellow;
                        }
                        curNode = curNode.NextSibling;
                    }
                }
            }
        }

        // -----------------------------------------------------------------
        // Save the resulting document.
        // -----------------------------------------------------------------
        const string outputPath = "CommentsHighlighted.docx";
        doc.Save(outputPath);
        // Inform the user (no interactive input required).
        Console.WriteLine($"Document saved to '{Path.GetFullPath(outputPath)}'.");
    }
}
