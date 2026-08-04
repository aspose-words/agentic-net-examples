using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a simple paragraph to host the comment.
        builder.Writeln("This paragraph will have a comment with a hyperlink.");

        // Create a top‑level comment.
        Comment comment = new Comment(doc, "Jane Doe", "JD", DateTime.Now);
        // Ensure the comment has at least one paragraph.
        comment.AppendChild(new Paragraph(doc));
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Move the builder into the comment's first paragraph to add content.
        builder.MoveTo(comment.FirstParagraph);
        // Insert a hyperlink field inside the comment.
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink("Aspose website", "https://www.aspose.com", false);
        builder.Font.ClearFormatting();

        // Show comments as PDF annotations.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;
        // Rebuild layout after changing display mode.
        doc.UpdatePageLayout();

        // Save the document to PDF.
        string pdfPath = "CommentWithHyperlink.pdf";
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Optional: open hyperlinks in a new window/tab.
            OpenHyperlinksInNewWindow = true
        };
        doc.Save(pdfPath, pdfOptions);

        // Reload the PDF to verify the hyperlink inside the comment.
        Document pdfDoc = new Document(pdfPath);
        var comments = pdfDoc.GetChildNodes(NodeType.Comment, true)
                             .OfType<Comment>()
                             .ToList();

        foreach (Comment c in comments)
        {
            // Look for hyperlink fields inside the comment's range.
            var hyperlinkField = c.Range.Fields
                                    .OfType<FieldHyperlink>()
                                    .FirstOrDefault();

            if (hyperlinkField != null)
            {
                Console.WriteLine($"Comment by {c.Author} contains hyperlink: {hyperlinkField.Address}");
            }
            else
            {
                Console.WriteLine($"Comment by {c.Author} does not contain a hyperlink.");
            }
        }
    }
}
