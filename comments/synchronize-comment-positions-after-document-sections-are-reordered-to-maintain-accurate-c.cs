using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Section 1 ----------
        builder.Writeln("This is the first section.");
        // Add a comment anchored to the paragraph we just created.
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Comment for the first section.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2 ----------
        builder.Writeln("This is the second section.");
        // Add a comment anchored to the second paragraph.
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Comment for the second section.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the original document for reference.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string originalPath = Path.Combine(outputDir, "Original.docx");
        doc.Save(originalPath);

        // ---------- Reorder Sections ----------
        // Move the second section before the first one.
        // Sections are children of the Document node; InsertBefore moves the node.
        Section secondSection = doc.Sections[1];
        Section firstSection = doc.Sections[0];
        firstSection.ParentNode.InsertBefore(secondSection, firstSection);

        // Save the reordered document.
        string reorderedPath = Path.Combine(outputDir, "Reordered.docx");
        doc.Save(reorderedPath);

        // ---------- Verify Comment Anchoring ----------
        // Enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        Console.WriteLine("Comments after reordering sections:");
        foreach (Comment comment in comments)
        {
            // The paragraph that the comment is attached to.
            Paragraph? parentParagraph = comment.ParentParagraph;
            string paragraphText = parentParagraph?.GetText().Trim() ?? "(no paragraph)";
            Console.WriteLine($"Author: {comment.Author}");
            Console.WriteLine($"Comment Text: {comment.GetText().Trim()}");
            Console.WriteLine($"Anchored Paragraph: \"{paragraphText}\"");
            Console.WriteLine();
        }

        // The program finishes without waiting for user input.
    }
}
