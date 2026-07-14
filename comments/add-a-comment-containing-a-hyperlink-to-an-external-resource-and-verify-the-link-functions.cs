using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will host the comment.
        builder.Writeln("Paragraph with a comment containing a hyperlink.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);

        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Move the builder into the comment's story to add content.
        builder.MoveTo(comment.AppendChild(new Paragraph(doc)));

        // Insert a hyperlink inside the comment.
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink("Aspose.Words", "https://www.aspose.com/words", false);
        builder.Font.ClearFormatting();

        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string pdfPath = Path.Combine(outputDir, "CommentWithHyperlink.pdf");

        // Save the document as PDF, preserving hyperlinks.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OpenHyperlinksInNewWindow = false
        };
        doc.Save(pdfPath, pdfOptions);
    }
}
