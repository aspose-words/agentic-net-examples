using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create an output folder for the generated documents.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------------------
        // 1. Build a sample document with two sections, each containing a
        //    paragraph and an anchored comment.
        // -------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Section 1 -----
        builder.Writeln("Section 1: This paragraph will have a comment.");

        // Create comment for the first paragraph.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Comment on Section 1.");

        // Anchor the comment to the paragraph using a comment range.
        Paragraph para1 = doc.FirstSection.Body.FirstParagraph;
        para1.AppendChild(new CommentRangeStart(doc, comment1.Id));
        para1.AppendChild(new Run(doc, "Commented text in Section 1."));
        para1.AppendChild(new CommentRangeEnd(doc, comment1.Id));
        para1.AppendChild(comment1);

        // Insert a page break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ----- Section 2 -----
        builder.Writeln("Section 2: This paragraph will also have a comment.");

        // Create comment for the second paragraph.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Comment on Section 2.");

        Paragraph? para2 = doc.LastSection.Body.FirstParagraph;
        if (para2 != null)
        {
            para2.AppendChild(new CommentRangeStart(doc, comment2.Id));
            para2.AppendChild(new Run(doc, "Commented text in Section 2."));
            para2.AppendChild(new CommentRangeEnd(doc, comment2.Id));
            para2.AppendChild(comment2);
        }

        // Save the original document.
        string originalPath = Path.Combine(outputDir, "Original.docx");
        doc.Save(originalPath);

        // -------------------------------------------------------------
        // 2. Reorder sections: move the second section before the first.
        // -------------------------------------------------------------
        if (doc.Sections.Count >= 2)
        {
            Section secondSection = doc.Sections[1];
            // Remove the second section and insert it at the beginning.
            doc.Sections.RemoveAt(1);
            doc.Sections.Insert(0, secondSection);
        }

        // -------------------------------------------------------------
        // 3. Synchronize comment IDs with their range markers after reordering.
        //    This ensures that each CommentRangeStart/End has the same Id as
        //    its associated Comment node.
        // -------------------------------------------------------------
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment comment in comments)
        {
            // The expected layout is:
            // CommentRangeStart -> Run (commented text) -> CommentRangeEnd -> Comment
            // Navigate backwards from the comment to locate the range nodes.
            Node? endNode = comment.PreviousSibling;
            Node? startNode = null;

            if (endNode is CommentRangeEnd endRange)
            {
                // The start node is three positions before the comment.
                startNode = endRange.PreviousSibling?.PreviousSibling?.PreviousSibling;
            }

            // If we successfully located the range nodes, synchronize their Ids.
            if (startNode is CommentRangeStart startRange && endNode is CommentRangeEnd endRangeNode)
            {
                if (startRange.Id != comment.Id)
                    startRange.Id = comment.Id;
                if (endRangeNode.Id != comment.Id)
                    endRangeNode.Id = comment.Id;
            }
        }

        // Save the reordered document.
        string reorderedPath = Path.Combine(outputDir, "Reordered.docx");
        doc.Save(reorderedPath);
    }
}
