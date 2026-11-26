using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain a comment.
        builder.Writeln("This is a paragraph that will contain a comment.");

        // Create a comment node with author information.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // Add visible text to the comment (a paragraph with a run).
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));

        // Attach the comment to the first paragraph of the document.
        Paragraph firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph != null)
        {
            firstParagraph.AppendChild(comment);
        }

        // Save the document as a DOC file.
        const string docPath = "sample.doc";
        doc.Save(docPath);

        // Load the document back from the file (demonstrates loading).
        Document loadedDoc = new Document(docPath);

        // Prepare PDF save options. In recent versions comments are exported
        // as PDF annotations by default, so no special flag is required.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF with comments preserved.
        const string pdfPath = "sample_with_comments.pdf";
        loadedDoc.Save(pdfPath, pdfOptions);
    }
}
