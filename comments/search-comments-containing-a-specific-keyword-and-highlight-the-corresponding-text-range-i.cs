using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment that contains the keyword "keyword".
        builder.Writeln("This is the first paragraph.");
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Please review this keyword.");
        Paragraph para1 = doc.FirstSection.Body.FirstParagraph;
        para1.AppendChild(new CommentRangeStart(doc, comment1.Id));
        para1.AppendChild(new Run(doc, "Text to be highlighted."));
        para1.AppendChild(new CommentRangeEnd(doc, comment1.Id));
        para1.AppendChild(comment1);

        // Second paragraph with a comment that does NOT contain the keyword.
        builder.Writeln("This is the second paragraph.");
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("No relevant term here.");
        Paragraph para2 = doc.FirstSection.Body.Paragraphs[2]; // Index after the two Writeln calls.
        para2.AppendChild(new CommentRangeStart(doc, comment2.Id));
        para2.AppendChild(new Run(doc, "Another piece of text."));
        para2.AppendChild(new CommentRangeEnd(doc, comment2.Id));
        para2.AppendChild(comment2);

        // Keyword to search for inside comment texts.
        const string keyword = "keyword";

        // Enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment comment in comments)
        {
            // Check if the comment text contains the keyword (case‑insensitive).
            if (comment.GetText().IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Find the matching CommentRangeStart node by comment Id.
                CommentRangeStart? start = doc.GetChildNodes(NodeType.CommentRangeStart, true)
                                             .OfType<CommentRangeStart>()
                                             .FirstOrDefault(r => r.Id == comment.Id);

                // Find the matching CommentRangeEnd node by comment Id.
                CommentRangeEnd? end = doc.GetChildNodes(NodeType.CommentRangeEnd, true)
                                         .OfType<CommentRangeEnd>()
                                         .FirstOrDefault(r => r.Id == comment.Id);

                if (start == null || end == null)
                    continue; // Safety check.

                // Iterate over all nodes between start and end (exclusive) and highlight runs.
                Node? curNode = start.NextSibling;
                while (curNode != null && curNode != end)
                {
                    if (curNode is Run run)
                    {
                        run.Font.HighlightColor = Color.Yellow;
                    }
                    curNode = curNode.NextSibling;
                }
            }
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
