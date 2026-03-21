using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

class ConvertDocToPdfWithComments
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text.
        builder.Writeln("This is a sample paragraph.");

        // Add a comment to the last paragraph.
        Comment comment = new Comment(doc, "Author", "A", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc) { ParagraphFormat = { StyleIdentifier = StyleIdentifier.CommentText } });
        comment.Paragraphs[0].AppendChild(new Run(doc, "This is a comment that will appear as a PDF annotation."));
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);

        // Make comments visible as PDF annotations.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;

        // Rebuild the layout after changing the comment display mode.
        doc.UpdatePageLayout();

        // Determine output path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConvertedDocument.pdf");

        // Save the document as PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(outputPath, pdfOptions);

        Console.WriteLine($"PDF saved to: {outputPath}");
    }
}
