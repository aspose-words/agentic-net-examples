using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold the comment.
        builder.Writeln("This paragraph will have a comment with a hyperlink.");

        // Create a comment node with required metadata.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // The comment must contain at least one paragraph.
        Paragraph commentParagraph = new Paragraph(doc);
        comment.AppendChild(commentParagraph);

        // Insert a hyperlink into the comment's paragraph.
        DocumentBuilder commentBuilder = new DocumentBuilder(doc);
        commentBuilder.MoveTo(commentParagraph);
        commentBuilder.InsertHyperlink(
            "Aspose.Words Documentation",
            "https://docs.aspose.com/words/net/",
            false);

        // Attach the comment to the first paragraph of the main story.
        Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph != null)
        {
            firstParagraph.AppendChild(comment);
        }

        // Prepare PDF save options (default options will export comments as annotations).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as PDF.
        string pdfPath = "CommentWithHyperlink.pdf";
        doc.Save(pdfPath, pdfOptions);

        // Verify that the hyperlink URL is present in the generated PDF file.
        bool linkFound = false;
        if (File.Exists(pdfPath))
        {
            byte[] pdfBytes = File.ReadAllBytes(pdfPath);
            string pdfText = Encoding.UTF8.GetString(pdfBytes);
            linkFound = pdfText.Contains("https://docs.aspose.com/words/net/");
        }

        // Output verification result.
        Console.WriteLine(linkFound
            ? "Hyperlink verified in the PDF."
            : "Hyperlink not found in the PDF.");
    }
}
