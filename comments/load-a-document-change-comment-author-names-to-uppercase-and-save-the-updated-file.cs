using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // Create a sample document with a comment.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph.");

        // Create a comment and attach it to the first paragraph.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // Add a paragraph and run to the comment so it has visible text.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "Sample comment text."));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the first paragraph of the document body.
        Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph != null)
        {
            firstParagraph.AppendChild(comment);
        }

        // Save the sample document.
        doc.Save(inputPath);

        // Load the document from disk.
        Document loadedDoc = new Document(inputPath);

        // Enumerate all comments in the document.
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        // Convert each comment's author name to uppercase.
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
