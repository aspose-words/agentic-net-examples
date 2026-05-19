using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will have a comment.
        builder.Writeln("This paragraph will have a comment containing a hyperlink.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Move the builder into the comment story to add content.
        builder.MoveTo(comment.AppendChild(new Paragraph(doc)));
        // Insert a hyperlink field inside the comment.
        builder.InsertHyperlink("Aspose.Words", "https://www.aspose.com/words", false);
        builder.Writeln(); // End the paragraph inside the comment.

        // Save the document to PDF with default options (hyperlinks are preserved).
        const string pdfPath = "CommentWithHyperlink.pdf";
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Reload the PDF to verify that the comment contains a hyperlink field.
        Document pdfDoc = new Document(pdfPath);
        var comments = pdfDoc.GetChildNodes(NodeType.Comment, true)
                             .OfType<Comment>()
                             .ToList();

        foreach (Comment c in comments)
        {
            bool hasHyperlink = c.GetChildNodes(NodeType.FieldStart, true)
                                 .OfType<FieldStart>()
                                 .Any(fs => fs.FieldType == FieldType.FieldHyperlink);
            Console.WriteLine($"Comment by {c.Author} contains hyperlink: {hasHyperlink}");
        }
    }
}
