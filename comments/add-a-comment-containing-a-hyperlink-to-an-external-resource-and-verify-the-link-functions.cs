using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold the comment.
        builder.Writeln("This paragraph will have a comment containing a hyperlink.");

        // Create a top‑level comment.
        Comment comment = new Comment(doc, "Jane Doe", "JD", DateTime.Now);
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Inside the comment, add a paragraph and insert a hyperlink.
        Paragraph commentParagraph = (Paragraph)comment.AppendChild(new Paragraph(doc));
        builder.MoveTo(commentParagraph);
        // Insert a hyperlink that points to an external URL.
        builder.InsertHyperlink("Aspose.Words", "https://www.aspose.com/words", false);

        // Optional: add some explanatory text after the hyperlink.
        builder.Write(" – official documentation.");

        // Ensure that comments are rendered as annotations in the PDF.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;
        // Rebuild the layout after changing the display mode.
        doc.UpdatePageLayout();

        // Configure PDF save options (e.g., open hyperlinks in a new window).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OpenHyperlinksInNewWindow = true
        };

        // Save the document as PDF.
        const string outputPath = "CommentWithHyperlink.pdf";
        doc.Save(outputPath, pdfOptions);

        // Verify that the comment was added (output to console for demonstration).
        var comments = doc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment c in comments)
        {
            Console.WriteLine($"Comment by {c.Author} on {c.DateTime}:");
            Console.WriteLine(c.GetText().Trim());
        }
    }
}
