using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Build a sample document with two sections,
        //    each containing a comment.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1
        builder.Writeln("Section 1 - First paragraph.");
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Comment attached to Section 1.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - First paragraph.");
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Comment attached to Section 2.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the original document.
        string originalPath = Path.Combine(outputDir, "Original.docx");
        doc.Save(originalPath);

        // -------------------------------------------------
        // 2. Reorder sections: move Section 2 before Section 1.
        // -------------------------------------------------
        // Ensure there are at least two sections.
        if (doc.Sections.Count >= 2)
        {
            // Remove the second section from its current position.
            Section secondSection = doc.Sections[1];
            doc.Sections.RemoveAt(1);

            // Insert it at the beginning of the collection.
            doc.Sections.Insert(0, secondSection);
        }

        // -------------------------------------------------
        // 3. Synchronize comment identifiers after reordering.
        //    Assign new sequential IDs and update the related
        //    CommentRangeStart/CommentRangeEnd nodes.
        // -------------------------------------------------
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true)
                             .OfType<CommentRangeStart>()
                             .ToList();

        var rangeEnds = doc.GetChildNodes(NodeType.CommentRangeEnd, true)
                           .OfType<CommentRangeEnd>()
                           .ToList();

        for (int i = 0; i < comments.Count; i++)
        {
            int oldId = comments[i].Id;
            int newId = i + 1;

            // Update the comment's Id.
            comments[i].Id = newId;

            // Update matching range start nodes.
            foreach (var start in rangeStarts.Where(rs => rs.Id == oldId))
                start.Id = newId;

            // Update matching range end nodes.
            foreach (var end in rangeEnds.Where(re => re.Id == oldId))
                end.Id = newId;
        }

        // Save the reordered document.
        string reorderedPath = Path.Combine(outputDir, "Reordered.docx");
        doc.Save(reorderedPath);
    }
}
