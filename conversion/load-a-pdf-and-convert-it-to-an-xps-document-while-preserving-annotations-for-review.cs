using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a comment.");

        // Create a comment that will become an annotation in PDF/XPS.
        Comment comment = new Comment(doc, "Reviewer", "RV", DateTime.Now);

        // A Comment must contain at least one Paragraph before adding Runs.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "Please review this paragraph."));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Render comments as annotations when saving to PDF.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Save the document as PDF (the comment becomes a PDF annotation).
        string pdfPath = "sample.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the generated PDF.
        Document pdfDoc = new Document(pdfPath);

        // Convert the PDF to XPS while preserving annotations.
        string xpsPath = "output.xps";
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        pdfDoc.Save(xpsPath, xpsOptions);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("The XPS file was not created.");

        // Optional cleanup of the intermediate PDF.
        // File.Delete(pdfPath);
    }
}
