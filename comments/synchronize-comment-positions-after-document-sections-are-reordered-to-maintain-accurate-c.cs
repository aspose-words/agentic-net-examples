using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document with two sections, each containing a paragraph and a comment.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1
        builder.Writeln("Paragraph in Section 1.");
        Comment comment1 = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "Comment for Section 1."));
        // Anchor the comment to the current paragraph.
        builder.CurrentParagraph?.AppendChild(comment1);

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Paragraph in Section 2.");
        Comment comment2 = new Comment(doc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now.AddMinutes(-5)
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Comment for Section 2."));
        builder.CurrentParagraph?.AppendChild(comment2);

        // Save the original document (optional, for verification).
        doc.Save("OriginalDocument.docx");

        // -----------------------------------------------------------------
        // Reorder sections: move the second section to be the first one.
        // -----------------------------------------------------------------
        if (doc.Sections.Count >= 2)
        {
            Section secondSection = doc.Sections[1];
            doc.Sections.RemoveAt(1);
            doc.Sections.Insert(0, secondSection);
        }

        // -----------------------------------------------------------------
        // Synchronize comment positions after reordering.
        // Capture existing comment data, remove all comments,
        // then re‑insert them anchored to the first paragraph of each section
        // in the new order.
        // -----------------------------------------------------------------

        // Capture comment data.
        var capturedComments = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Select(c => new
            {
                Text = c.GetText().Trim(),
                Author = c.Author,
                Initial = c.Initial,
                DateTime = c.DateTime
            })
            .ToList();

        // Remove existing comments safely.
        foreach (Comment c in doc.GetChildNodes(NodeType.Comment, true).OfType<Comment>().ToList())
        {
            c.Remove();
        }

        // Re‑insert comments, one per section, anchored to the first paragraph.
        int commentIndex = 0;
        foreach (Section section in doc.Sections)
        {
            Paragraph? firstParagraph = section.Body?.FirstParagraph;
            if (firstParagraph == null || commentIndex >= capturedComments.Count)
                continue;

            var data = capturedComments[commentIndex];

            Comment newComment = new Comment(doc)
            {
                Author = data.Author,
                Initial = data.Initial,
                DateTime = data.DateTime
            };
            newComment.AppendChild(new Paragraph(doc));
            newComment.FirstParagraph?.AppendChild(new Run(doc, data.Text));

            // Anchor the comment to the first paragraph of the current section.
            firstParagraph.AppendChild(newComment);

            commentIndex++;
        }

        // Save the final document with reordered sections and synchronized comments.
        doc.Save("ReorderedComments.docx");
    }
}
