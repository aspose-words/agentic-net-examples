using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Section 1 ----------
        // Write a paragraph in the first section.
        builder.Writeln("This is the first paragraph in Section 1.");

        // Create a comment for the paragraph we just added.
        Paragraph para1 = doc.FirstSection.Body.LastParagraph;
        Comment comment1 = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "First section comment."));
        para1?.AppendChild(comment1);

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2 ----------
        // Write a paragraph in the second section.
        builder.Writeln("This is the first paragraph in Section 2.");

        // Create a comment for the paragraph we just added.
        // The new paragraph is the last paragraph of the second section.
        Paragraph para2 = doc.Sections[1].Body.LastParagraph;
        Comment comment2 = new Comment(doc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now.AddMinutes(-5)
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Second section comment."));
        para2?.AppendChild(comment2);

        // Save the original document.
        string originalPath = "original.docx";
        doc.Save(originalPath);

        // ---------- Reorder Sections ----------
        // Move the first section to the end of the document.
        if (doc.Sections.Count >= 2)
        {
            Section firstSection = doc.Sections[0];
            doc.Sections.RemoveAt(0);
            doc.Sections.Add(firstSection);
        }

        // Verify that comments are still correctly anchored.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            // The comment is attached to a paragraph (its parent node).
            Paragraph? attachedParagraph = c.ParentNode as Paragraph;
            string paragraphText = attachedParagraph?.GetText().Trim() ?? "(no paragraph found)";
            Console.WriteLine($"Comment by {c.Author} attached to paragraph: \"{paragraphText}\"");
        }

        // Save the reordered document.
        string reorderedPath = "reordered.docx";
        doc.Save(reorderedPath);
    }
}
