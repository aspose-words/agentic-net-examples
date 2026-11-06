using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class DeleteCommentsByAuthor
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with comments from two different authors.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment from John Doe.
        builder.Writeln("This is the first paragraph.");
        Paragraph? firstPara = builder.CurrentParagraph;
        if (firstPara != null)
        {
            Comment commentJohn = new Comment(doc)
            {
                Author = "John Doe",
                Initial = "JD",
                DateTime = DateTime.Now
            };
            commentJohn.AppendChild(new Paragraph(doc));
            commentJohn.FirstParagraph?.AppendChild(new Run(doc, "John's comment."));
            firstPara.AppendChild(commentJohn);
        }

        // Second paragraph with a comment from Jane Smith.
        builder.Writeln("This is the second paragraph.");
        Paragraph? secondPara = builder.CurrentParagraph;
        if (secondPara != null)
        {
            Comment commentJane = new Comment(doc)
            {
                Author = "Jane Smith",
                Initial = "JS",
                DateTime = DateTime.Now
            };
            commentJane.AppendChild(new Paragraph(doc));
            commentJane.FirstParagraph?.AppendChild(new Run(doc, "Jane's comment."));
            secondPara.AppendChild(commentJane);
        }

        // Save the original document.
        string originalPath = Path.Combine(outputDir, "original.docx");
        doc.Save(originalPath);

        // Load the document for processing.
        Document loadedDoc = new Document(originalPath);

        // Author whose comments should be removed.
        string targetAuthor = "John Doe";

        // Find all comments authored by the target author.
        var commentsToRemove = loadedDoc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Remove each matching comment safely.
        foreach (Comment comment in commentsToRemove)
        {
            comment.Remove();
        }

        // Save the modified document.
        string modifiedPath = Path.Combine(outputDir, "modified.docx");
        loadedDoc.Save(modifiedPath);
    }
}
