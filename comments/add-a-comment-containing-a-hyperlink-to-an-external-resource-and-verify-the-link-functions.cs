using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold the comment.
        builder.Writeln("This paragraph will have a comment with a hyperlink.");

        // Create a comment node with metadata.
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
        commentBuilder.InsertHyperlink("https://www.example.com", "Visit Example", false);

        // Attach the comment to the first paragraph of the document body.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Save the document as PDF.
        string pdfPath = "CommentWithHyperlink.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the hyperlink URL is present in the generated PDF.
        bool hyperlinkFound = false;
        if (File.Exists(pdfPath))
        {
            byte[] pdfBytes = File.ReadAllBytes(pdfPath);
            string pdfText = Encoding.ASCII.GetString(pdfBytes);
            hyperlinkFound = pdfText.Contains("https://www.example.com");
        }

        Console.WriteLine(hyperlinkFound
            ? "Hyperlink verified in PDF."
            : "Hyperlink not found in PDF.");
    }
}
