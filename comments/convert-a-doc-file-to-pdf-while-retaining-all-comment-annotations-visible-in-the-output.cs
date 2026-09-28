using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for the sample DOC and resulting PDF.
        string docPath = Path.Combine(outputDir, "sample.doc");
        string pdfPath = Path.Combine(outputDir, "sample.pdf");

        // -----------------------------------------------------------------
        // 1. Create a DOC file with a comment.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will hold the comment.
        builder.Writeln("This is a paragraph that contains a comment.");

        // Create a comment node.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // The comment must contain at least one paragraph and run with visible text.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "Please review this paragraph."));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the last paragraph of the document.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.LastParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Save the DOC file.
        doc.Save(docPath, SaveFormat.Doc);

        // -----------------------------------------------------------------
        // 2. Load the DOC file and convert it to PDF while preserving comments.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Optional: enumerate comments to demonstrate they are present.
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        foreach (Comment c in comments)
        {
            Console.WriteLine($"{c.Author}: {c.GetText().Trim()}");
        }

        // Save as PDF. Comments are rendered as PDF annotations by default.
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);
    }
}
