using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        const string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a sample document with a paragraph and a comment
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will receive a comment
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a comment node
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // Add visible text to the comment
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));

        // Attach the comment to the first paragraph of the document
        Paragraph? targetParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Save the original document (optional, for reference)
        string originalPath = Path.Combine(outputDir, "original.docx");
        doc.Save(originalPath);

        // Extract comments and embed them as footnotes in the same document
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Add a heading for the footnote summary
        builder.Writeln();
        builder.Writeln("Comments Summary (as footnotes):");

        foreach (Comment c in comments)
        {
            string commentText = c.GetText()?.Trim() ?? string.Empty;
            builder.Write($"Comment by {c.Author}: ");
            builder.InsertFootnote(FootnoteType.Footnote, commentText);
            builder.Writeln();
        }

        // Save the final document with footnotes
        string finalPath = Path.Combine(outputDir, "comments-with-footnotes.docx");
        doc.Save(finalPath);
    }
}
