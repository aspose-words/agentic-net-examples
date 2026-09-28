using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample paragraphs.
        builder.Writeln("First paragraph with a comment.");
        builder.Writeln("Second paragraph without a comment.");
        builder.Writeln("Third paragraph with another comment.");

        // Add a comment to the first paragraph.
        Comment comment1 = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "Please review the first paragraph."));

        // Attach the comment to the first paragraph node.
        Paragraph? firstPara = doc.FirstSection?.Body?.FirstParagraph;
        firstPara?.AppendChild(comment1);

        // Add a comment to the third paragraph.
        Paragraph? thirdPara = doc.FirstSection?.Body?.Paragraphs[2];
        if (thirdPara != null)
        {
            Comment comment2 = new Comment(doc)
            {
                Author = "Bob",
                Initial = "B",
                DateTime = DateTime.Now.AddMinutes(-5)
            };
            comment2.AppendChild(new Paragraph(doc));
            comment2.FirstParagraph?.AppendChild(new Run(doc, "Consider rephrasing this sentence."));

            thirdPara.AppendChild(comment2);
        }

        // Move builder to the end of the document to add a footnote summary.
        builder.MoveToDocumentEnd();
        builder.Writeln();
        builder.Writeln("Comments presented as footnotes:");

        // Enumerate all comments safely.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        foreach (Comment c in comments)
        {
            // Write a brief label before the footnote.
            builder.Write($"Comment by {c.Author}: ");
            // Insert a footnote containing the comment text.
            string commentText = c.GetText()?.Trim() ?? string.Empty;
            builder.InsertFootnote(FootnoteType.Footnote, commentText);
            builder.Writeln();
        }

        // Ensure output directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentsFootnotes.docx");
        doc.Save(outputPath);
    }
}
