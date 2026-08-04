using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a deterministic output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Build a sample document with two sections,
        //    each containing a comment anchored to a range.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- First section -----
        builder.Writeln("Section 1 - Introduction");
        builder.Writeln("This paragraph will have a comment.");

        // Create the first comment.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("First comment.");

        // Anchor the comment to a range inside the current paragraph.
        Paragraph para1 = builder.CurrentParagraph;
        para1.AppendChild(new CommentRangeStart(doc, comment1.Id));
        para1.AppendChild(new Run(doc, "Commented text."));
        para1.AppendChild(new CommentRangeEnd(doc, comment1.Id));
        para1.AppendChild(comment1);

        // Insert a section break (new page) to start the second section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ----- Second section -----
        builder.Writeln("Section 2 - Details");
        builder.Writeln("Another paragraph with a comment.");

        // Create the second comment.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Second comment.");

        // Anchor the second comment.
        Paragraph para2 = builder.CurrentParagraph;
        para2.AppendChild(new CommentRangeStart(doc, comment2.Id));
        para2.AppendChild(new Run(doc, "Commented text."));
        para2.AppendChild(new CommentRangeEnd(doc, comment2.Id));
        para2.AppendChild(comment2);

        // Save the original document.
        string originalPath = Path.Combine(outputDir, "Original.docx");
        doc.Save(originalPath);

        // -------------------------------------------------
        // 2. Reorder sections: move the second section before the first.
        // -------------------------------------------------
        Document reorderedDoc = (Document)doc.Clone(true);

        if (reorderedDoc.Sections.Count >= 2)
        {
            Section secondSection = reorderedDoc.Sections[1];
            reorderedDoc.Sections.RemoveAt(1);
            reorderedDoc.Sections.Insert(0, secondSection);
        }

        // -------------------------------------------------
        // 3. Synchronize comment IDs with their associated range nodes.
        //    After moving sections, the comment IDs may no longer be sequential.
        // -------------------------------------------------
        // Collect all top‑level comments.
        List<Comment> comments = reorderedDoc.GetChildNodes(NodeType.Comment, true)
                                            .OfType<Comment>()
                                            .Where(c => c.Ancestor == null)
                                            .ToList();

        // Collect all range start/end nodes once for efficiency.
        List<CommentRangeStart> rangeStarts = reorderedDoc.GetChildNodes(NodeType.CommentRangeStart, true)
                                                         .OfType<CommentRangeStart>()
                                                         .ToList();
        List<CommentRangeEnd> rangeEnds = reorderedDoc.GetChildNodes(NodeType.CommentRangeEnd, true)
                                                     .OfType<CommentRangeEnd>()
                                                     .ToList();

        int nextId = 1;
        foreach (Comment comment in comments)
        {
            int oldId = comment.Id;
            comment.Id = nextId;

            foreach (CommentRangeStart start in rangeStarts.Where(r => r.Id == oldId))
                start.Id = nextId;

            foreach (CommentRangeEnd end in rangeEnds.Where(r => r.Id == oldId))
                end.Id = nextId;

            nextId++;
        }

        // -------------------------------------------------
        // 4. Save the reordered document with synchronized comments.
        // -------------------------------------------------
        string reorderedPath = Path.Combine(outputDir, "Reordered_Synchronized.docx");
        reorderedDoc.Save(reorderedPath);

        Console.WriteLine($"Original document saved to: {originalPath}");
        Console.WriteLine($"Reordered document saved to: {reorderedPath}");
    }
}
