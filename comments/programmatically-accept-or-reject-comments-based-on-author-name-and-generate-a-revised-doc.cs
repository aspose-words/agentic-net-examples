using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several comments from different authors.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment from Alice.
        builder.Writeln("This is the first paragraph.");
        Comment commentAlice = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        commentAlice.AppendChild(new Paragraph(doc));
        commentAlice.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));
        // Attach the comment to the first paragraph.
        doc.FirstSection.Body.FirstParagraph?.AppendChild(commentAlice);

        // Second paragraph with a comment from Charlie.
        builder.Writeln("This is the second paragraph.");
        Comment commentCharlie = new Comment(doc)
        {
            Author = "Charlie",
            Initial = "C",
            DateTime = DateTime.Now
        };
        commentCharlie.AppendChild(new Paragraph(doc));
        commentCharlie.FirstParagraph?.AppendChild(new Run(doc, "Consider rephrasing this sentence."));
        // Attach the comment to the second paragraph.
        doc.FirstSection.Body.Paragraphs[1]?.AppendChild(commentCharlie);

        // Third paragraph with a comment from Bob.
        builder.Writeln("This is the third paragraph.");
        Comment commentBob = new Comment(doc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now
        };
        commentBob.AppendChild(new Paragraph(doc));
        commentBob.FirstParagraph?.AppendChild(new Run(doc, "Add a citation here."));
        // Attach the comment to the third paragraph.
        doc.FirstSection.Body.Paragraphs[2]?.AppendChild(commentBob);

        // Save the original document.
        string originalPath = "original.docx";
        doc.Save(originalPath);

        // Load the document for processing (simulating a separate workflow).
        Document processedDoc = new Document(originalPath);

        // Define authors whose comments should be accepted (kept).
        string[] acceptedAuthors = { "Alice", "Bob" };

        // Find comments that do NOT match the accepted authors.
        var commentsToRemove = processedDoc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Where(c => !acceptedAuthors.Contains(c.Author, StringComparer.OrdinalIgnoreCase))
            .ToList();

        // Remove the unwanted comments safely.
        foreach (Comment c in commentsToRemove)
        {
            c.Remove();
        }

        // Save the revised document.
        string revisedPath = "revised.docx";
        processedDoc.Save(revisedPath);
    }
}
