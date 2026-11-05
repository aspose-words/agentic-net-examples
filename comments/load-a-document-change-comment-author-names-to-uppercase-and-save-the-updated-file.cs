using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare file names.
        string inputPath = "input.docx";
        string outputPath = "output.docx";

        // Create a sample document with a comment.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph with a comment.");

        // Create a comment node.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // Add visible text to the comment.
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph?.AppendChild(comment);

        // Save the initial document.
        doc.Save(inputPath);

        // Load the document from file.
        Document loadedDoc = new Document(inputPath);

        // Enumerate all comments in the document.
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        // Convert each comment author name to uppercase.
        foreach (Comment c in comments)
        {
            if (!string.IsNullOrEmpty(c.Author))
            {
                c.Author = c.Author.ToUpperInvariant();
            }
        }

        // Save the updated document.
        loadedDoc.Save(outputPath);
    }
}
