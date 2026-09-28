using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with a comment
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample paragraph with a comment.");

        // Create a comment node
        Comment comment = new Comment(sampleDoc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // Add visible text to the comment
        comment.AppendChild(new Paragraph(sampleDoc));
        comment.FirstParagraph?.AppendChild(new Run(sampleDoc, "Please review this paragraph."));

        // Attach the comment to the first paragraph
        Paragraph? firstParagraph = sampleDoc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph != null)
        {
            firstParagraph.AppendChild(comment);
        }

        // Save the sample document
        string inputPath = Path.Combine(outputDir, "input.docx");
        sampleDoc.Save(inputPath);

        // Load the document from file
        Document loadedDoc = new Document(inputPath);

        // Enumerate all comments in the document
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        // Convert each comment author name to uppercase
        foreach (Comment c in comments)
        {
            if (!string.IsNullOrEmpty(c.Author))
            {
                c.Author = c.Author.ToUpperInvariant();
            }
        }

        // Save the updated document
        string outputPath = Path.Combine(outputDir, "output.docx");
        loadedDoc.Save(outputPath);
    }
}
