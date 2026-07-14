using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names for the sample files.
        const string pdfPath = "sample.pdf";
        const string xpsPath = "sample.xps";

        // -----------------------------------------------------------------
        // 1. Create a sample Word document with a comment (annotation).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Document created for conversion demo.");

        // Create a comment node and attach it to the current paragraph.
        // The comment must contain at least one paragraph with a run.
        Comment comment = new Comment(sourceDoc, "Reviewer", "RV", DateTime.Now);
        Paragraph commentParagraph = new Paragraph(sourceDoc);
        commentParagraph.AppendChild(new Run(sourceDoc, "This paragraph contains a review comment."));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the paragraph that was just written.
        builder.CurrentParagraph.AppendChild(comment);

        // Ensure comments are rendered as PDF annotations.
        sourceDoc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Save the document as PDF – this embeds the comment as a PDF annotation.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create PDF file '{pdfPath}'.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and convert it to XPS while preserving annotations.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // XpsSaveOptions preserve annotations by default.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        pdfDoc.Save(xpsPath, xpsOptions);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException($"Failed to create XPS file '{xpsPath}'.");

        Console.WriteLine("PDF successfully converted to XPS with annotations preserved.");
    }
}
