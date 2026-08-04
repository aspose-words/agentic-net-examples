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

        // First paragraph (no comment).
        builder.Writeln("This is a regular paragraph without comments.");

        // Insert first comment with the keyword "important".
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("This is an important comment.");

        // Add a commented range that contains the keyword.
        builder.Writeln(); // creates a new empty paragraph.
        Paragraph para1 = doc.FirstSection.Body.LastParagraph;
        para1.AppendChild(new CommentRangeStart(doc, comment1.Id));
        para1.AppendChild(new Run(doc, "Important text inside comment range."));
        para1.AppendChild(new CommentRangeEnd(doc, comment1.Id));
        para1.AppendChild(comment1);

        // Insert second comment without the keyword.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Just a regular comment.");

        builder.Writeln(); // another new paragraph.
        Paragraph para2 = doc.FirstSection.Body.LastParagraph;
        para2.AppendChild(new CommentRangeStart(doc, comment2.Id));
        para2.AppendChild(new Run(doc, "Some other text."));
        para2.AppendChild(new CommentRangeEnd(doc, comment2.Id));
        para2.AppendChild(comment2);

        // Keyword to search for inside comment texts.
        const string keyword = "important";

        // Highlight the text ranges of comments that contain the keyword.
        HighlightCommentRanges(doc, keyword);

        // Save the resulting document.
        doc.Save("HighlightedComments.docx");
    }

    private static void HighlightCommentRanges(Document doc, string keyword)
    {
        // Enumerate all top‑level comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment comment in comments)
        {
            // Skip replies – they have an ancestor comment.
            if (comment.Ancestor != null)
                continue;

            // Check if the comment text contains the keyword (case‑insensitive).
            string commentText = comment.GetText() ?? string.Empty;
            if (commentText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) < 0)
                continue;

            // Locate the matching comment range start/end nodes.
            CommentRangeStart? start = doc.GetChildNodes(NodeType.CommentRangeStart, true)
                                          .OfType<CommentRangeStart>()
                                          .FirstOrDefault(s => s.Id == comment.Id);
            CommentRangeEnd? end = doc.GetChildNodes(NodeType.CommentRangeEnd, true)
                                      .OfType<CommentRangeEnd>()
                                      .FirstOrDefault(e => e.Id == comment.Id);

            if (start == null || end == null)
                continue; // Safety check – malformed comment.

            // Walk through the nodes that lie between the start and end markers.
            Node? node = start.NextSibling;
            while (node != null && node != end)
            {
                if (node is Run run)
                {
                    // Apply yellow highlight to the run.
                    run.Font.HighlightColor = Color.Yellow;
                }

                node = node.NextSibling;
            }
        }
    }
}
