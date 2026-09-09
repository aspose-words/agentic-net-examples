using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace CommentSyncExample
{
    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a sample document with two sections, each containing a paragraph and a comment.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ----- Section 1 -----
            builder.Writeln("Section 1: This is the first section.");
            Paragraph? para1 = doc.FirstSection?.Body?.FirstParagraph;
            if (para1 != null)
                AddCommentToParagraph(doc, para1, "Alice", "A", "Comment for Section 1.");

            // Insert a new section break.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // ----- Section 2 -----
            builder.Writeln("Section 2: This is the second section.");
            Paragraph? para2 = doc.LastSection?.Body?.LastParagraph;
            if (para2 != null)
                AddCommentToParagraph(doc, para2, "Bob", "B", "Comment for Section 2.");

            // Save the original document.
            string originalPath = Path.Combine(outputDir, "Original.docx");
            doc.Save(originalPath);

            // ----- Reorder Sections -----
            // Swap the two sections so that Section 2 becomes the first one.
            if (doc.Sections.Count >= 2)
            {
                Section first = doc.Sections[0];
                Section second = doc.Sections[1];

                // Remove the second section first, then insert it at the beginning.
                doc.Sections.RemoveAt(1);
                doc.Sections.Insert(0, second);
                // The original first section now follows the second section automatically.
            }

            // Synchronize comment positions after reordering.
            SynchronizeComments(doc);

            // Save the reordered document.
            string reorderedPath = Path.Combine(outputDir, "Reordered.docx");
            doc.Save(reorderedPath);
        }

        // Adds a comment to the specified paragraph.
        private static void AddCommentToParagraph(Document doc, Paragraph paragraph, string author, string initials, string commentText)
        {
            // Create a comment with metadata.
            Comment comment = new Comment(doc, author, initials, DateTime.Now);
            comment.SetText(commentText); // This creates the comment's internal paragraphs.

            // Insert the comment range start, the commented text, the range end, and finally the comment node.
            CommentRangeStart rangeStart = new CommentRangeStart(doc, comment.Id);
            CommentRangeEnd rangeEnd = new CommentRangeEnd(doc, comment.Id);

            paragraph.AppendChild(rangeStart);
            paragraph.AppendChild(new Run(doc, "Commented text."));
            paragraph.AppendChild(rangeEnd);
            paragraph.AppendChild(comment);
        }

        // Ensures that each comment's range start/end IDs match the comment's Id
        // and that the comment node appears immediately after its range end.
        private static void SynchronizeComments(Document doc)
        {
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment comment in comments)
            {
                // Find the CommentRangeEnd node that should precede this comment.
                CommentRangeEnd? rangeEnd = comment.PreviousSibling as CommentRangeEnd;

                // If the immediate previous sibling is not the range end, search backwards.
                if (rangeEnd == null)
                {
                    Node? node = comment.PreviousSibling;
                    while (node != null && !(node is CommentRangeEnd))
                    {
                        node = node.PreviousSibling;
                    }
                    rangeEnd = node as CommentRangeEnd;
                }

                // Find the matching CommentRangeStart node by scanning backwards from the range end.
                CommentRangeStart? rangeStart = null;
                if (rangeEnd != null)
                {
                    Node? node = rangeEnd.PreviousSibling;
                    while (node != null && !(node is CommentRangeStart))
                    {
                        node = node.PreviousSibling;
                    }
                    rangeStart = node as CommentRangeStart;
                }

                // Synchronize IDs if both range nodes are found.
                if (rangeStart != null && rangeEnd != null)
                {
                    rangeStart.Id = comment.Id;
                    rangeEnd.Id = comment.Id;
                }

                // Ensure the comment node is placed immediately after its range end.
                if (rangeEnd != null && comment.ParentNode != null && comment.PreviousSibling != rangeEnd)
                {
                    comment.Remove();
                    rangeEnd.ParentNode?.InsertAfter(comment, rangeEnd);
                }
            }
        }
    }
}
