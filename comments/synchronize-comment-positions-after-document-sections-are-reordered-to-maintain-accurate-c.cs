using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Section 1 ----------
        builder.Writeln("Section 1 - Paragraph 1.");

        // Create a comment for the first paragraph.
        Comment comment1 = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        // Add visible text to the comment.
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "Comment for Section 1."));

        // Anchor the comment to the current paragraph.
        Paragraph para1 = builder.CurrentParagraph;
        para1?.AppendChild(comment1);

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2 ----------
        builder.Writeln("Section 2 - Paragraph 1.");

        // Create a comment for the second paragraph.
        Comment comment2 = new Comment(doc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Comment for Section 2."));

        // Anchor the comment to the current paragraph.
        Paragraph para2 = builder.CurrentParagraph;
        para2?.AppendChild(comment2);

        // Save the original document.
        string originalPath = "original.docx";
        doc.Save(originalPath);

        // ---------- Reorder Sections ----------
        // Capture references to the existing sections.
        Section firstSection = doc.Sections[0];
        Section secondSection = doc.Sections[1];

        // Remove sections from the document (order matters).
        doc.Sections.RemoveAt(1); // Remove second section.
        doc.Sections.RemoveAt(0); // Remove first section.

        // Insert them back in swapped order.
        doc.Sections.Insert(0, secondSection);
        doc.Sections.Insert(1, firstSection);

        // After reordering, verify that each comment is still anchored to a paragraph
        // within the correct (new) section.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        Console.WriteLine("Comments after section reordering:");
        foreach (Comment c in comments)
        {
            // The comment should be a child of a Paragraph.
            Paragraph? parentParagraph = c.ParentNode as Paragraph;
            if (parentParagraph == null)
            {
                Console.WriteLine($"Comment by {c.Author} is not attached to a paragraph.");
                continue;
            }

            // Find the section that contains this paragraph.
            Section? parentSection = parentParagraph.ParentNode as Section;
            int sectionIndex = parentSection != null ? doc.Sections.IndexOf(parentSection) + 1 : -1;

            Console.WriteLine($"Author: {c.Author}, Section: {sectionIndex}, Text: {c.GetText().Trim()}");
        }

        // Save the reordered document.
        string reorderedPath = "reordered.docx";
        doc.Save(reorderedPath);
    }
}
