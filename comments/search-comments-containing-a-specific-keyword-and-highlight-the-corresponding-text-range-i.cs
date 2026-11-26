using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace CommentKeywordHighlighter
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with a comment.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Paragraph without comment.
            builder.Writeln("First paragraph without comment.");

            // Insert a comment range start.
            const int commentId = 1;
            CommentRangeStart rangeStart = new CommentRangeStart(doc, commentId);
            builder.CurrentParagraph.AppendChild(rangeStart);

            // Text that will be inside the comment range.
            builder.Writeln("This text is inside the comment range and will be highlighted if the comment contains the keyword: highlight.");

            // Insert a comment range end.
            CommentRangeEnd rangeEnd = new CommentRangeEnd(doc, commentId);
            builder.CurrentParagraph.AppendChild(rangeEnd);

            // Create the comment node itself.
            Comment comment = new Comment(doc)
            {
                Id = commentId,
                Author = "John Doe",
                Initial = "JD",
                DateTime = DateTime.Now
            };
            // Add visible content to the comment.
            comment.AppendChild(new Paragraph(doc));
            comment.FirstParagraph.AppendChild(new Run(doc, "This is a comment containing the keyword: highlight."));

            // Append the comment to the document's comments story.
            doc.FirstSection.Body.FirstParagraph.AppendChild(comment);

            // Save the original document.
            const string originalPath = "original.docx";
            doc.Save(originalPath);
            Console.WriteLine($"Original document saved to '{originalPath}'.");

            // Keyword to search for inside comment text.
            const string keyword = "keyword";

            // Find all comments that contain the keyword (case‑insensitive).
            List<Comment> matchingComments = doc.GetChildNodes(NodeType.Comment, true)
                .OfType<Comment>()
                .Where(c => c.GetText()?.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            // Highlight the text range associated with each matching comment.
            foreach (Comment matchingComment in matchingComments)
            {
                // Locate the corresponding CommentRangeStart and CommentRangeEnd nodes by Id.
                CommentRangeStart? startNode = doc.GetChildNodes(NodeType.CommentRangeStart, true)
                    .OfType<CommentRangeStart>()
                    .FirstOrDefault(s => s.Id == matchingComment.Id);
                CommentRangeEnd? endNode = doc.GetChildNodes(NodeType.CommentRangeEnd, true)
                    .OfType<CommentRangeEnd>()
                    .FirstOrDefault(e => e.Id == matchingComment.Id);

                if (startNode == null || endNode == null)
                    continue; // Skip if the range nodes are missing.

                // Iterate over all nodes between the start and end nodes (exclusive).
                for (Node cur = startNode.NextSibling; cur != null && cur != endNode; cur = cur.NextSibling)
                {
                    // Highlight runs directly.
                    if (cur is Run run)
                    {
                        run.Font.HighlightColor = Color.Yellow;
                    }
                    // If the node is a paragraph, highlight all runs it contains.
                    else if (cur is Paragraph para)
                    {
                        foreach (Run innerRun in para.GetChildNodes(NodeType.Run, true).OfType<Run>())
                        {
                            innerRun.Font.HighlightColor = Color.Yellow;
                        }
                    }
                }
            }

            // Save the modified document.
            const string outputPath = "highlighted.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document with highlighted ranges saved to '{outputPath}'.");
        }
    }
}
