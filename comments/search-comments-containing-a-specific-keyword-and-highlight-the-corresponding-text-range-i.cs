using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

#nullable enable

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with the keyword "important".
        builder.Writeln("This is a sample paragraph with an important note.");

        // Locate the run that contains the word "important".
        Run? importantRun = FindRunContainingText(builder.CurrentParagraph, "important");
        if (importantRun != null)
        {
            // Create a comment anchored to the word "important".
            Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
            comment.SetText("This comment contains the keyword.");

            // Insert the comment range start before the run.
            builder.CurrentParagraph.InsertBefore(new CommentRangeStart(doc, comment.Id), importantRun);
            // Insert the comment range end after the run.
            builder.CurrentParagraph.InsertAfter(new CommentRangeEnd(doc, comment.Id), importantRun);
            // Append the comment node after the range end.
            builder.CurrentParagraph.AppendChild(comment);
        }

        // Second paragraph without the keyword.
        builder.Writeln("Another paragraph with a regular comment.");

        // Create a comment for the whole paragraph.
        Comment otherComment = new Comment(doc, "Bob", "B", DateTime.Now);
        otherComment.SetText("Just a regular comment.");

        // Anchor the comment to the entire paragraph.
        Paragraph secondPara = builder.CurrentParagraph;
        secondPara.InsertBefore(new CommentRangeStart(doc, otherComment.Id), secondPara.FirstChild);
        secondPara.AppendChild(new CommentRangeEnd(doc, otherComment.Id));
        secondPara.AppendChild(otherComment);

        // Define the keyword to search for in comments.
        const string keyword = "keyword";

        // Enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        foreach (Comment c in comments)
        {
            // Check if the comment text contains the keyword (case‑insensitive).
            if (c.GetText().IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Find the matching CommentRangeStart node by Id.
                CommentRangeStart? rangeStart = doc.GetChildNodes(NodeType.CommentRangeStart, true)
                    .OfType<CommentRangeStart>()
                    .FirstOrDefault(crs => crs.Id == c.Id);

                if (rangeStart != null)
                {
                    // Walk through the nodes between the start and the corresponding end.
                    Node? node = rangeStart.NextSibling;
                    while (node != null)
                    {
                        // Stop when we reach the matching CommentRangeEnd.
                        if (node.NodeType == NodeType.CommentRangeEnd && ((CommentRangeEnd)node).Id == c.Id)
                            break;

                        // Highlight any Run nodes inside the range.
                        if (node.NodeType == NodeType.Run)
                        {
                            ((Run)node).Font.HighlightColor = Color.Yellow;
                        }

                        node = node.NextSibling;
                    }
                }
            }
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the modified document.
        string outputPath = Path.Combine(outputDir, "HighlightedComments.docx");
        doc.Save(outputPath);
    }

    // Helper method to find the first Run in a paragraph that contains the specified text.
    private static Run? FindRunContainingText(Paragraph paragraph, string text)
    {
        foreach (Run run in paragraph.Runs)
        {
            if (run.Text != null && run.Text.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                return run;
        }
        return null;
    }
}
