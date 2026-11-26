using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with an old comment (40 days ago).
        builder.Writeln("First paragraph with an old comment.");
        Paragraph firstParagraph = builder.CurrentParagraph;

        Comment oldComment = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now.AddDays(-40)
        };
        // Add visible text to the comment.
        oldComment.AppendChild(new Paragraph(doc));
        oldComment.FirstParagraph?.AppendChild(new Run(doc, "This comment is older than the threshold."));
        // Attach the comment to the paragraph.
        firstParagraph.AppendChild(oldComment);

        // Second paragraph with a recent comment (10 days ago).
        builder.Writeln("Second paragraph with a recent comment.");
        Paragraph secondParagraph = builder.CurrentParagraph;

        Comment recentComment = new Comment(doc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now.AddDays(-10)
        };
        recentComment.AppendChild(new Paragraph(doc));
        recentComment.FirstParagraph?.AppendChild(new Run(doc, "This comment is newer than the threshold."));
        secondParagraph.AppendChild(recentComment);

        // Define the date threshold (30 days ago).
        DateTime threshold = DateTime.Now.AddDays(-30);

        // Safely collect all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        // Remove comments older than the threshold.
        foreach (Comment c in comments)
        {
            if (c.DateTime < threshold)
            {
                c.Remove();
            }
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the filtered document.
        string outputPath = Path.Combine(outputDir, "filtered-comments.docx");
        doc.Save(outputPath);
    }
}
