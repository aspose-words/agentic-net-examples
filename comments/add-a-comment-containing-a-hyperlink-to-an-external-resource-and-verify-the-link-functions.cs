using System;
using System.IO;
using System.Drawing;
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

        // Create a comment and attach it to the paragraph.
        Comment comment = new Comment(doc, "Jane Doe", "JD", DateTime.Now);
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
        paragraph.AppendChild(comment);

        // Inside the comment, add a paragraph and a hyperlink to an external URL.
        Paragraph commentParagraph = (Paragraph)comment.AppendChild(new Paragraph(doc));
        DocumentBuilder commentBuilder = new DocumentBuilder(doc);
        commentBuilder.MoveTo(commentParagraph);
        commentBuilder.Font.Color = Color.Blue;
        commentBuilder.Font.Underline = Underline.Single;
        commentBuilder.InsertHyperlink("Aspose.Words", "https://www.aspose.com/words", false);
        commentBuilder.Font.ClearFormatting();

        // Configure the document to show comments as PDF annotations.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;
        doc.UpdatePageLayout();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document as PDF.
        string pdfPath = Path.Combine(outputDir, "CommentWithHyperlink.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Simple verification that the PDF file was created and is not empty.
        if (File.Exists(pdfPath) && new FileInfo(pdfPath).Length > 0)
        {
            Console.WriteLine("PDF saved successfully with a comment containing a hyperlink.");
        }
        else
        {
            Console.WriteLine("Failed to create the PDF file.");
        }
    }
}
