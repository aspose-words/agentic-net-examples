using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create an output folder in the current working directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for the intermediate DOC file and the final PDF file.
        string docPath = Path.Combine(outputDir, "sample.doc");
        string pdfPath = Path.Combine(outputDir, "sample.pdf");

        // -----------------------------------------------------------------
        // Create a new document and add a paragraph with a comment.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a comment node.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // The comment must contain at least one paragraph and run with visible text.
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph for accuracy."));

        // Attach the comment to the paragraph that was just created.
        Paragraph? targetParagraph = builder.CurrentParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Save the document as a temporary DOC file.
        doc.Save(docPath, SaveFormat.Doc);

        // -----------------------------------------------------------------
        // Load the DOC file and convert it to PDF.
        // Comments are preserved as PDF annotations by default.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);
    }
}
