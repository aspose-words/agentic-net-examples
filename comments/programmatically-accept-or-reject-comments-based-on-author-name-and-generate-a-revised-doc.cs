using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a few paragraphs and comments.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment from Alice.
        builder.Writeln("First paragraph with a comment.");
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
        builder.Writeln("Second paragraph with another comment.");
        Comment commentCharlie = new Comment(doc)
        {
            Author = "Charlie",
            Initial = "C",
            DateTime = DateTime.Now.AddDays(-1)
        };
        commentCharlie.AppendChild(new Paragraph(doc));
        commentCharlie.FirstParagraph?.AppendChild(new Run(doc, "Consider rephrasing this sentence."));
        // Attach the comment to the second paragraph.
        doc.FirstSection.Body.LastParagraph?.AppendChild(commentCharlie);

        // Third paragraph with a comment from Bob.
        builder.Writeln("Third paragraph with a third comment.");
        Comment commentBob = new Comment(doc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now.AddHours(-5)
        };
        commentBob.AppendChild(new Paragraph(doc));
        commentBob.FirstParagraph?.AppendChild(new Run(doc, "Check the data accuracy."));
        // Attach the comment to the third paragraph.
        doc.FirstSection.Body.LastParagraph?.AppendChild(commentBob);

        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the original document.
        string originalPath = Path.Combine(outputDir, "OriginalDocument.docx");
        doc.Save(originalPath);

        // Define authors whose comments should be accepted (kept).
        HashSet<string> acceptedAuthors = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Alice",
            "Bob"
        };

        // Enumerate all comments safely.
        List<Comment> allComments = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        // Remove comments whose author is not in the accepted list.
        foreach (Comment c in allComments)
        {
            string author = c.Author ?? string.Empty;
            if (!acceptedAuthors.Contains(author))
            {
                c.Remove();
            }
        }

        // Save the revised document with only accepted comments.
        string revisedPath = Path.Combine(outputDir, "RevisedDocument.docx");
        doc.Save(revisedPath);
    }
}
